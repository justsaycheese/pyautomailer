import subprocess
import shutil
import zipfile
from pathlib import Path

REPO_DIR = Path(__file__).resolve().parent
DIST_DIR = REPO_DIR / "dist"
BUILD_DIR = REPO_DIR / "build"
EXE_NAME = "automailer_verZ.exe"
EXE_BASENAME = EXE_NAME.rsplit('.', 1)[0]
ZIP_NAME = REPO_DIR / "release.zip"


def run_pyinstaller():
    cmd = [
        "pyinstaller",
        "--clean",
        "--onefile",
        "--name",
        EXE_BASENAME,
        str(REPO_DIR / "automailer_verZ.py"),
    ]
    subprocess.run(cmd, check=True)

    produced = DIST_DIR / EXE_BASENAME
    exe_with_ext = DIST_DIR / EXE_NAME
    if produced.exists() and not exe_with_ext.exists():
        produced.rename(exe_with_ext)


def prepare_directories():
    (REPO_DIR / "embed").mkdir(exist_ok=True)
    (REPO_DIR / "attachment").mkdir(exist_ok=True)


def create_zip():
    with zipfile.ZipFile(ZIP_NAME, "w", zipfile.ZIP_DEFLATED) as zf:
        exe_path = DIST_DIR / EXE_NAME
        if exe_path.exists():
            zf.write(exe_path, EXE_NAME)
        for folder in ["embed", "attachment"]:
            folder_path = REPO_DIR / folder
            for p in folder_path.glob("*"):
                zf.write(p, f"{folder}/{p.name}")
            if not any(folder_path.iterdir()):
                # store empty folder
                info = zipfile.ZipInfo(f"{folder}/")
                zf.writestr(info, "")
        msg_file = REPO_DIR / "sample.msg"
        if msg_file.exists():
            zf.write(msg_file, "sample.msg")


def clean_build_artifacts():
    shutil.rmtree(DIST_DIR, ignore_errors=True)
    shutil.rmtree(BUILD_DIR, ignore_errors=True)
    spec_file = REPO_DIR / f"{EXE_NAME.rsplit('.',1)[0]}.spec"
    if spec_file.exists():
        spec_file.unlink()
    if ZIP_NAME.exists():
        ZIP_NAME.unlink()


if __name__ == "__main__":
    clean_build_artifacts()
    prepare_directories()
    run_pyinstaller()
    create_zip()
    print(f"Created {ZIP_NAME}")
