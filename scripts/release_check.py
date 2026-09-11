from __future__ import annotations

import argparse
import hashlib
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from pathlib import PurePosixPath


FORBIDDEN_SEGMENTS = {
    ".git",
    ".pytest_cache",
    ".venv",
    "venv",
    "site-packages",
    "__pycache__",
    "build",
    "dist",
}
SENSITIVE_PATTERNS = (
    re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----"),
    re.compile(r"(?i)\b(?:api[_-]?key|access[_-]?token|secret[_-]?key)\s*[:=]"),
    re.compile(r"(?i)[A-Z]:\\Users\\"),
)


def fail(message: str) -> None:
    raise SystemExit(f"RELEASE_CHECK=FAIL\n- {message}")


def read_pyproject_version(path: Path) -> str:
    if not path.is_file():
        fail(f"缺少 pyproject.toml：{path}")
    text = path.read_text(encoding="utf-8")
    project_section = re.search(r"(?ms)^\[project\]\s*(.*?)(?=^\[|\Z)", text)
    if project_section is None:
        fail(f"pyproject.toml 缺少 [project]：{path}")
    match = re.search(r"(?m)^\s*version\s*=\s*[\"']([^\"']+)[\"']\s*$", project_section.group(1))
    if match is None:
        fail(f"pyproject.toml 缺少 [project].version：{path}")
    return match.group(1)


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest().upper()


def _scan_release_content(relative: str, text: str) -> None:
    if any(part in FORBIDDEN_SEGMENTS for part in PurePosixPath(relative).parts):
        fail(f"发布包包含禁止目录/文件：{relative}")
    for pattern in SENSITIVE_PATTERNS:
        if pattern.search(text):
            fail(f"发布包疑似包含敏感信息：{relative}")
    if re.search(r"(?i)[A-Z]:\\|/Users/|/home/", text):
        fail(f"发布包包含绝对本地路径：{relative}")


def check_package(package: Path, version: str) -> list[str]:
    if not package.is_dir():
        fail(f"发布目录不存在：{package}")
    for relative in ("SKILL.md", "VERSION", "requirements.txt", "pyproject.toml"):
        if not package.joinpath(relative).is_file():
            fail(f"发布目录缺少：{relative}")
    if package.joinpath("VERSION").read_text(encoding="utf-8").strip() != version:
        fail("发布目录 VERSION 与当前 VERSION 不一致")
    if read_pyproject_version(package / "pyproject.toml") != version.removeprefix("V"):
        fail("发布目录 VERSION 与 pyproject.toml version 不一致")

    files: list[str] = []
    for path in sorted(package.rglob("*")):
        relative = path.relative_to(package)
        if path.is_dir():
            continue
        relative_name = relative.as_posix()
        files.append(relative_name)
        text = path.read_text(encoding="utf-8", errors="ignore")
        _scan_release_content(relative_name, text)
    if "SKILL.md" not in files:
        fail("发布包根目录没有 SKILL.md")
    return files


def check_zip_contents(zip_path: Path, package: Path, package_files: list[str]) -> None:
    try:
        with zipfile.ZipFile(zip_path) as archive:
            zip_files: dict[str, bytes] = {}
            for info in archive.infolist():
                if info.is_dir():
                    continue
                relative = info.filename.replace("\\", "/")
                if relative in zip_files:
                    fail(f"ZIP 包含重复文件：{relative}")
                data = archive.read(info)
                zip_files[relative] = data
                _scan_release_content(relative, data.decode("utf-8", errors="ignore"))

            package_set = set(package_files)
            zip_set = set(zip_files)
            if zip_set != package_set:
                missing = sorted(package_set - zip_set)
                extra = sorted(zip_set - package_set)
                fail(f"ZIP 文件清单不一致：缺失={missing or '无'}；多余={extra or '无'}")

            for relative in package_files:
                package_hash = sha256_bytes((package / Path(*relative.split("/"))).read_bytes())
                zip_hash = sha256_bytes(zip_files[relative])
                if package_hash != zip_hash:
                    fail(f"ZIP 文件内容与发布目录不一致：{relative}")
    except zipfile.BadZipFile as exc:
        fail(f"ZIP 无法打开：{exc}")


def main() -> int:
    parser = argparse.ArgumentParser(description="Independent release checks for Zeyang Excel Cleaner FREE")
    parser.add_argument("--package-dir", type=Path, required=True)
    parser.add_argument("--zip", dest="zip_path", type=Path, required=True)
    args = parser.parse_args()

    root = Path(__file__).resolve().parents[1]
    version = (root / "VERSION").read_text(encoding="utf-8").strip()
    if not re.fullmatch(r"V\d+\.\d+\.\d+", version):
        fail("VERSION 不是普通语义化版本")
    pyproject_version = read_pyproject_version(root / "pyproject.toml")
    if pyproject_version != version.removeprefix("V"):
        fail("VERSION 与 pyproject.toml version 不一致")
    if not (root / "SKILL.md").is_file() or not (root / "requirements.txt").is_file():
        fail("根目录缺少 SKILL.md 或 requirements.txt")

    tests = subprocess.run([sys.executable, "-m", "pytest", "-q"], cwd=root, text=True, capture_output=True)
    if tests.returncode != 0:
        fail("自动测试未通过\n" + tests.stdout + tests.stderr)
    files = check_package(args.package_dir, version)
    if not args.zip_path.is_file():
        fail(f"ZIP 不存在：{args.zip_path}")
    if args.zip_path.stat().st_size >= 5 * 1024 * 1024:
        fail("ZIP 超过 5 MB")
    with zipfile.ZipFile(args.zip_path) as archive:
        bad = archive.testzip()
        if bad:
            fail(f"ZIP 校验失败：{bad}")
    check_zip_contents(args.zip_path, args.package_dir, files)

    print("RELEASE_CHECK=PASS")
    print(f"VERSION={version}")
    print(f"TESTS={tests.stdout.strip().splitlines()[-1]}")
    print(f"PACKAGE_FILES={len(files)}")
    print(f"ZIP_SIZE={args.zip_path.stat().st_size}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
