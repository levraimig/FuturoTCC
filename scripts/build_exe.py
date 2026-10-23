r"""
Script para gerar o executável standalone (.exe) com PyInstaller
Uso: .\venv\Scripts\python.exe scripts\build_exe.py
"""
import shutil
import subprocess
import sys
from pathlib import Path


def main() -> int:
    raiz = Path(__file__).resolve().parent.parent
    entrada = raiz / "__main__.py"

    comando = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--windowed",
        "--name",
        "LabimedDownload",
        "--collect-all",
        "customtkinter",
        "--collect-all",
        "selenium",
        "--collect-all",
        "webdriver_manager",
        str(entrada),
    ]

    print("Gerando executavel com PyInstaller...")
    resultado = subprocess.run(comando, cwd=str(raiz))
    if resultado.returncode == 0:
        exe_dist = raiz / "dist" / "LabimedDownload.exe"
        exe_raiz = raiz / "LabimedDownload.exe"
        try:
            shutil.copy2(exe_dist, exe_raiz)
        except OSError:
            pass
        shutil.rmtree(raiz / "build", ignore_errors=True)
        spec_file = raiz / "LabimedDownload.spec"
        if spec_file.exists():
            spec_file.unlink(missing_ok=True)
        print(f"\n[SUCESSO] Executavel gerado em: {exe_raiz}")
    return resultado.returncode


if __name__ == "__main__":
    raise SystemExit(main())
