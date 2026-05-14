"""
build_gerar_link.py - Empacota gerar_link.py em Gerar_Link.exe (PyInstaller).

Saida: dist/Gerar_Link.exe
Coloque esse Gerar_Link.exe na MESMA pasta do Planilhas.exe.
Quando o usuario der duplo clique, o Link.txt sera (re)gerado/atualizado
ao lado, com o IP correto da rede WiFi atual.
"""
import os
import shutil
import subprocess
import sys
from pathlib import Path

AQUI = Path(__file__).parent.resolve()
ENTRY = AQUI / "gerar_link.py"
NOME = "Gerar_Link"


def main():
    if not ENTRY.exists():
        print(f"[ERRO] Nao encontrei {ENTRY}")
        sys.exit(1)

    # Limpa builds anteriores deste utilitario
    for d in ("build_gerar_link", f"dist/{NOME}.exe"):
        p = AQUI / d
        if p.is_dir():
            shutil.rmtree(p, ignore_errors=True)
        elif p.is_file():
            try:
                p.unlink()
            except Exception:
                pass

    icon = AQUI / "icon.ico"

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--console",  # mostra janela preta com mensagem (usuario deve VER o feedback)
        "--name", NOME,
        "--workpath", "build_gerar_link",
        "--specpath", "build_gerar_link",
        "--noconfirm",
    ]
    if icon.exists():
        cmd += ["--icon", str(icon)]
    cmd.append(str(ENTRY))

    print("\n=== Empacotando Gerar_Link.exe ===")
    print(" ".join(cmd))
    r = subprocess.run(cmd, cwd=str(AQUI))
    if r.returncode != 0:
        print("\n[ERRO] PyInstaller falhou.")
        sys.exit(r.returncode)

    final = AQUI / "dist" / f"{NOME}.exe"
    if final.exists():
        print(f"\n[OK] Gerado: {final}")
        # Copia para a Area de Trabalho para facilitar
        try:
            desktop = Path.home() / "Desktop"
            if not desktop.exists():
                desktop = Path.home() / "Area de Trabalho"
            if desktop.exists():
                destino = desktop / f"{NOME}.exe"
                shutil.copy2(final, destino)
                print(f"[OK] Copiado tambem para: {destino}")
        except Exception as e:
            print(f"[AVISO] Nao foi possivel copiar para Desktop: {e}")
    else:
        print(f"\n[ERRO] Saida nao encontrada em {final}")
        sys.exit(1)


if __name__ == "__main__":
    main()
