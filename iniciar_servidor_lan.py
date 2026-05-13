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
import json
import urllib.request
import urllib.error
from pathlib import Path

DEFAULT_SERVER = "https://planilhas-1.onrender.com"


def _ler_licenca():
    """Lê licenca.txt da pasta do exe. Retorna (token, server_url) ou (None, None)."""
    arq = _pasta_do_exe() / "licenca.txt"
    if not arq.exists():
        return None, None
    token = None
    server = None
    try:
        for linha in arq.read_text(encoding="utf-8").splitlines():
            linha = linha.strip()
            if linha.startswith("#") or not linha:
                continue
            if linha.upper().startswith("TOKEN="):
                token = linha.split("=", 1)[1].strip()
            elif linha.upper().startswith("SERVER="):
                server = linha.split("=", 1)[1].strip()
    except Exception as e:
        print(f"[AVISO] Erro lendo licenca.txt: {e}")
    return token, (server or DEFAULT_SERVER)


def _hardware_id():
    """Gera ID único estavel deste PC (CPU+MAC+hostname+disco)."""
    import hashlib
    import platform
    import uuid
    partes = []
    try:
        partes.append(str(uuid.getnode()))
    except Exception:
        partes.append("nomac")
    try:
        partes.append(platform.node() or "noname")
    except Exception:
        partes.append("noname")
    try:
        partes.append(platform.machine() or "noarch")
    except Exception:
        partes.append("noarch")
    try:
        partes.append(platform.processor() or "nocpu")
    except Exception:
        partes.append("nocpu")
    try:
        if os.name == "nt":
            r = subprocess.run(["vol", "C:"], capture_output=True, text=True,
                               timeout=3, shell=True)
            partes.append(r.stdout.strip())
    except Exception:
        pass
    return hashlib.sha256("|".join(partes).encode("utf-8")).hexdigest()


def _post_json(url, payload, timeout=10):
    """POST JSON simples sem dependencia externa."""
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url, data=data,
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _validar_licenca():
    """Valida licenca antes de iniciar o servidor.
    Retorna True se OK, False se bloqueado (encerra).
    """
    token, server = _ler_licenca()
    if not token:
        print("\n" + "!" * 64)
        print(" LICENCA NAO ENCONTRADA ".center(64, "!"))
        print("!" * 64)
        print("\n  Falta o arquivo 'licenca.txt' na mesma pasta do Planilhas.exe.")
        print("  Acesse o painel online > Licença Desktop e baixe o arquivo.")
        print("  Coloque-o ao lado do Planilhas.exe e rode novamente.\n")
        input("Pressione Enter para sair...")
        return False

    hw = _hardware_id()
    print(f"[LICENCA] Verificando ativacao em {server} ...")

    # Tenta verificar primeiro; se ainda nao ativada, ativa agora
    try:
        r = _post_json(
            f"{server.rstrip('/')}/api/licenca/verificar",
            {"token": token, "hardware_id": hw}
        )
    except Exception as e:
        print(f"[AVISO] Sem internet ou servidor indisponivel: {e}")
        print("        Continuando em modo offline (sera revalidado depois).")
        return True  # tolerancia se sem internet

    if r.get("ok"):
        print("[LICENCA] OK - este PC esta autorizado.\n")
        return True

    motivo = (r.get("motivo") or "").lower()

    # Tenta primeira ativacao
    if "nao ativada" in motivo or "não ativada" in motivo:
        print("[LICENCA] Primeira execucao. Ativando neste PC...")
        try:
            ar = _post_json(
                f"{server.rstrip('/')}/api/licenca/ativar",
                {"token": token, "hardware_id": hw}
            )
        except Exception as e:
            print(f"[ERRO] Falha ao ativar: {e}")
            input("Pressione Enter para sair...")
            return False
        if ar.get("ok"):
            print("[LICENCA] Ativada com sucesso neste PC!\n")
            return True
        motivo = ar.get("motivo", "Falha na ativacao")

    # Bloqueado
    print("\n" + "!" * 64)
    print(" LICENCA BLOQUEADA ".center(64, "!"))
    print("!" * 64)
    print(f"\n  Motivo: {motivo}")
    print("\n  Possiveis causas:")
    print("    - Voce esta usando o exe de outra empresa")
    print("    - O exe ja foi ativado em outro PC")
    print("    - O administrador precisa resetar a licenca no painel online\n")
    input("Pressione Enter para sair...")
    return False


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
    # 0) VALIDA LICENÇA ANTES DE QUALQUER COISA (anti-pirataria)
    if not _validar_licenca():
        return

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
