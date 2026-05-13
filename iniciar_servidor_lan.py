"""
Entry point do executavel PLANILHAS_SERVIDOR.exe.

Ao rodar:
  1. Define modo rede local (PLANILHAS_LAN=1)
  2. Detecta o IP da rede WiFi/LAN
  3. Monta o link compartilhavel (ex: http://192.168.0.15:5000/)
  4. Salva esse link em "Link_Planilhas.txt" na Area de Trabalho
  5. Copia o link para a area de transferencia (Ctrl+V)
  6. Abre o navegador local do usuario
  7. Inicia o servidor Flask em 0.0.0.0:5000
"""
import os
import sys
import socket
import time
import threading
import webbrowser
import subprocess
from pathlib import Path


def _get_local_ip():
    """Descobre o IP da rede local (WiFi/LAN)."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def _copiar_para_clipboard(texto):
    """Copia texto para a area de transferencia do Windows."""
    try:
        subprocess.run(
            "clip",
            input=texto.encode("utf-8"),
            shell=True,
            check=False,
        )
        return True
    except Exception:
        return False


def _pasta_do_exe():
    """Retorna a pasta onde o .exe (ou o .py) esta rodando."""
    if getattr(sys, "frozen", False):
        # Rodando como .exe (PyInstaller)
        return Path(sys.executable).parent
    return Path(__file__).parent


def _salvar_link_desktop(link):
    """Salva o link em 'Link.txt' na pasta do .exe E na Area de Trabalho."""
    conteudo = (
        "=== LINK PARA COMPARTILHAR COM SUA EQUIPE ===\n\n"
        f"{link}\n\n"
        "Envie este endereco pelo WhatsApp, email ou Telegram.\n"
        "Todos devem estar conectados na MESMA rede WiFi.\n"
        "Mantenha este PC ligado enquanto a equipe usa o sistema.\n\n"
        "CREDENCIAIS PADRAO:\n"
        "  Email: admin@planilhas.com\n"
        "  Senha: admin123\n"
    )
    caminhos_salvos = []

    # 1) Salvar ao lado do .exe (sempre atualiza ao dar duplo clique)
    try:
        arq_exe = _pasta_do_exe() / "Link.txt"
        arq_exe.write_text(conteudo, encoding="utf-8")
        caminhos_salvos.append(str(arq_exe))
    except Exception as e:
        print(f"[AVISO] Nao foi possivel salvar Link.txt ao lado do exe: {e}")

    # 2) Salvar tambem na Area de Trabalho (conveniencia)
    try:
        desktop = Path.home() / "Desktop"
        if not desktop.exists():
            desktop = Path.home() / "Área de Trabalho"
        if desktop.exists():
            arq_desk = desktop / "Link_Planilhas.txt"
            arq_desk.write_text(conteudo, encoding="utf-8")
            caminhos_salvos.append(str(arq_desk))
    except Exception as e:
        print(f"[AVISO] Nao foi possivel salvar no Desktop: {e}")

    return caminhos_salvos


def _liberar_firewall(port):
    """Tenta liberar a porta no firewall do Windows (ignora erros)."""
    try:
        subprocess.run(
            [
                "netsh", "advfirewall", "firewall", "add", "rule",
                f"name=Planilhas LAN {port}",
                "dir=in", "action=allow", "protocol=TCP",
                f"localport={port}",
            ],
            capture_output=True,
            shell=False,
            check=False,
        )
    except Exception:
        pass


def _abrir_navegador(url, delay=3.0):
    """Abre o navegador local apos alguns segundos."""
    def _worker():
        time.sleep(delay)
        try:
            webbrowser.open(url)
        except Exception:
            pass
    threading.Thread(target=_worker, daemon=True).start()


def _banner(ip_local, port, arquivos_link):
    larg = 64
    linha = "=" * larg
    print("\n" + linha)
    print(" >> SERVIDOR PLANILHAS.COM - MODO REDE LOCAL (WiFi) ".center(larg))
    print(linha)
    print()
    print(f"  LINK PARA COMPARTILHAR:  http://{ip_local}:{port}/")
    print(f"  LINK NESTE PC:           http://127.0.0.1:{port}/")
    print()
    print("  >> O link ja foi COPIADO para a area de transferencia")
    print("     (Ctrl+V para colar no WhatsApp / Email)")
    if arquivos_link:
        print("  >> Tambem foi salvo em:")
        for a in arquivos_link:
            print(f"       - {a}")
    print()
    print("  CREDENCIAIS PADRAO:")
    print("     Email: admin@planilhas.com")
    print("     Senha: admin123")
    print()
    print("  Mantenha esta janela ABERTA enquanto a equipe usa o sistema.")
    print("  Feche a janela para desligar o servidor.")
    print(linha + "\n")


def main():
    # 1) Ativa modo LAN ANTES de importar o app.py
    os.environ["PLANILHAS_LAN"] = "1"
    os.environ.setdefault("PORT", "5000")
    port = int(os.environ["PORT"])

    # 2) Descobre IP da rede local
    ip_local = _get_local_ip()
    link_lan = f"http://{ip_local}:{port}/comecar?next=/executar-sistema"

    # 3) Libera firewall (silencioso)
    _liberar_firewall(port)

    # 4) Salva link ao lado do exe + Desktop, copia para clipboard
    arquivos = _salvar_link_desktop(link_lan)
    _copiar_para_clipboard(link_lan)

    # 5) Mostra banner informativo
    _banner(ip_local, port, arquivos)

    # 6) Abre navegador local automaticamente
    _abrir_navegador(f"http://127.0.0.1:{port}/", delay=3.0)

    # 7) Importa e inicia o app.py Flask
    try:
        from app import app as flask_app
        flask_app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)
    except OSError as e:
        print(f"\n[ERRO] Nao foi possivel iniciar na porta {port}: {e}")
        print("A porta pode estar em uso por outra instancia do servidor.")
        input("\nPressione Enter para sair...")
    except Exception as e:
        import traceback
        print(f"\n[ERRO FATAL] {e}")
        traceback.print_exc()
        input("\nPressione Enter para sair...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nServidor encerrado pelo usuario.")
