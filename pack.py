import argparse
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
DIST_DIR = BASE_DIR / "dist"
BUILD_DIR = BASE_DIR / "build"
STAGING_DIST_DIR = BUILD_DIR / "release-dist"
STAGING_PORTABLE_DIR = BUILD_DIR / "EasyChat_Momo_Portable"
SPEC_PATH = BASE_DIR / "wechat_gui_momo.spec"
EXE_PATH = STAGING_DIST_DIR / "wechat_gui_momo.exe"
PORTABLE_DIR = DIST_DIR / "EasyChat_Momo_Portable"
PORTABLE_ZIP = BASE_DIR / "wechat_gui_momo_portable.zip"


def run_build():
    try:
        import PyInstaller  # noqa: F401
    except ImportError as exc:
        raise SystemExit(
            "PyInstaller is not installed. Run:\n"
            f'"{sys.executable}" -m pip install -r requirements-build.txt'
        ) from exc

    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--distpath",
        str(STAGING_DIST_DIR),
        str(SPEC_PATH),
    ]
    subprocess.run(command, cwd=BASE_DIR, check=True)
    if not EXE_PATH.is_file():
        raise SystemExit(f"Build completed without the expected executable: {EXE_PATH}")


def build_portable_package():
    if STAGING_PORTABLE_DIR.exists():
        shutil.rmtree(STAGING_PORTABLE_DIR)
    STAGING_PORTABLE_DIR.mkdir(parents=True)

    files = [
        EXE_PATH,
        BASE_DIR / "run_momo.vbs",
        BASE_DIR / "wechat_config_momo.json",
        BASE_DIR / "README.md",
    ]
    for source in files:
        shutil.copy2(source, STAGING_PORTABLE_DIR / source.name)

    if PORTABLE_ZIP.exists():
        PORTABLE_ZIP.unlink()
    with zipfile.ZipFile(PORTABLE_ZIP, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(STAGING_PORTABLE_DIR.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(STAGING_PORTABLE_DIR.parent))

    try:
        if PORTABLE_DIR.exists():
            shutil.rmtree(PORTABLE_DIR)
        shutil.copytree(STAGING_PORTABLE_DIR, PORTABLE_DIR)
        shutil.copy2(EXE_PATH, DIST_DIR / EXE_PATH.name)
    except PermissionError:
        print("Existing dist files are in use; the portable ZIP was still updated.")


def main():
    parser = argparse.ArgumentParser(description="Build EasyChat Momo")
    parser.add_argument(
        "--portable",
        action="store_true",
        help="also create wechat_gui_momo_portable.zip",
    )
    args = parser.parse_args()

    run_build()
    print(f"Executable: {EXE_PATH}")
    if args.portable:
        build_portable_package()
        print(f"Portable package: {PORTABLE_ZIP}")


if __name__ == "__main__":
    main()
