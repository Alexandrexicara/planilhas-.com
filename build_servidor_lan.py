"""
Gera o executavel PLANILHAS_SERVIDOR.exe

Uso:
    python build_servidor_lan.py

Resultado:
    dist/PLANILHAS_SERVIDOR.exe   - executavel portatil
    Area de Trabalho do usuario   - copia automatica
"""
import os
import shutil
import subprocess
import sys
from pathlib import Path


def get_desktop_path():
    desktop = Path.home() / "Desktop"
    if desktop.exists():
        return str(desktop)
    desktop_pt = Path.home() / "Área de Trabalho"
    if desktop_pt.exists():
        return str(desktop_pt)
    return str(Path.home())


def build():
    print("=" * 60)
    print("  GERANDO PLANILHAS_SERVIDOR.exe")
    print("=" * 60)

    hidden_imports = [
        "app",
        "sistema",
        "sistema_plus",
        "usuarios_db",
        "web_access_db",
        "planilhas_paths",
        "gerenciamento_usuarios",
        "openpyxl",
        "flask",
        "werkzeug",
        "werkzeug.security",
        "jinja2",
        "sqlite3",
        "PIL",
        "PIL.Image",
    ]

    add_data = [
        "--add-data=templates;templates",
        "--add-data=static;static",
    ]
    # Adicionar bancos se existirem
    for db in ("banco.db", "banco_plus.db", "usuarios.db", "acesso_web.db"):
        if os.path.exists(db):
            add_data.append(f"--add-data={db};.")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--console",                           # Mostra janela com o link
        "--name=Planilhas",
        "--collect-all=openpyxl",
        "--collect-all=flask",
        "--collect-all=werkzeug",
        "--collect-all=jinja2",
        "--distpath=dist",
        "--workpath=build",
        "--specpath=.",
        "--noconfirm",
        "--clean",
        "--noupx",
    ]
    cmd.extend(add_data)
    for h in hidden_imports:
        cmd.extend(["--hidden-import", h])

    if os.path.exists("icon.ico"):
        cmd.extend(["--icon", "icon.ico"])

    cmd.append("iniciar_servidor_lan.py")

    print("Executando PyInstaller (pode demorar 1-2 minutos)...\n")
    try:
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as e:
        print(f"\n[ERRO] PyInstaller falhou: {e}")
        return False

    exe_path = os.path.join("dist", "Planilhas.exe")
    if not os.path.exists(exe_path):
        print("\n[ERRO] Executavel nao foi gerado.")
        return False

    # Copia para a Area de Trabalho
    desktop = get_desktop_path()
    destino = os.path.join(desktop, "Planilhas.exe")
    try:
        shutil.copy2(exe_path, destino)
        print(f"\n[OK] Copiado para: {destino}")
    except Exception as e:
        print(f"[AVISO] Nao foi possivel copiar para Desktop: {e}")

    print("\n" + "=" * 60)
    print("  CONCLUIDO!")
    print("=" * 60)
    print(f"  Executavel: {exe_path}")
    print(f"  Atalho:     {destino}")
    print()
    print("  COMO USAR:")
    print("    1. De duplo clique em Planilhas.exe")
    print("    2. O link compartilhavel sera COPIADO automaticamente")
    print("    3. Sera salvo em 'Link.txt' na mesma pasta do .exe")
    print("    4. Cole (Ctrl+V) no WhatsApp/email da sua equipe")
    print("    5. Todos devem estar na MESMA rede WiFi")
    print("=" * 60)
    return True


if __name__ == "__main__":
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("PyInstaller nao encontrado. Instalando...")
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "pyinstaller"],
            check=True,
        )

    if os.path.exists("build"):
        shutil.rmtree("build", ignore_errors=True)

    ok = build()
    input("\nPressione Enter para sair...")
    sys.exit(0 if ok else 1)
