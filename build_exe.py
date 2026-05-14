import os
import shutil
import subprocess
import sys
from pathlib import Path


def get_desktop_path():
    return str(Path.home() / "Desktop")


def build_executable():
    print("Gerando executavel Windows...")

    hidden_imports = [
        "sistema",
        "sistema_plus",
        "usuarios_db",
        "gerenciamento_usuarios",
        "sistema_online_offline",
        "banco_offline",
        "openpyxl",
        "zipfile",
        "sqlite3",
        "json",
        "datetime",
        "tkinter",
    ]

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--windowed",
        "--name=Planilhas",
        "--add-data=banco_plus.db;.",
        "--add-data=banco.db;.",
        "--add-data=usuarios.db;.",
        "--collect-all=openpyxl",
        "--distpath=dist",
        "--workpath=build",
        "--specpath=.",
        "--noconfirm",
        "--clean",
        "--noupx",
    ]

    for hidden_import in hidden_imports:
        cmd.extend(["--hidden-import", hidden_import])

    if os.path.exists("icon.ico"):
        cmd.extend(["--icon", "icon.ico"])

    cmd.append("menu_principal.py")

    try:
        subprocess.run(cmd, capture_output=False, text=True, check=True)
        print("Executavel gerado com sucesso.")

        exe_path = os.path.join("dist", "Planilhas.exe")
        desktop_path = get_desktop_path()
        desktop_exe = os.path.join(desktop_path, "Planilhas.exe")

        if os.path.exists(exe_path):
            shutil.copy2(exe_path, desktop_exe)
            print(f"Copiado para Desktop: {desktop_exe}")

        return True
    except subprocess.CalledProcessError as e:
        print(f"Erro ao gerar executavel: {e}")
        return False


if __name__ == "__main__":
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("PyInstaller nao encontrado. Instalando...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)

    if os.path.exists("build"):
        shutil.rmtree("build")

    success = build_executable()
    print("Concluido." if success else "Falhou.")
    input("Pressione Enter para sair...")
