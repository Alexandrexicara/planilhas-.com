"""
Gerar_Link.exe - Utilitario de duplo clique.

O que faz quando o usuario da DUPLO CLIQUE:
  1. Descobre o IP da rede WiFi/LAN deste PC
  2. Monta o link compartilhavel: http://IP:5000/comecar?next=/executar-sistema
  3. Salva esse link em "Link.txt" na MESMA pasta deste arquivo (do lado do Planilhas.exe)
  4. Copia o link para a area de transferencia (Ctrl+V)
  5. Abre o Link.txt no Bloco de Notas para o usuario ver
  6. Mostra mensagem rapida no console e fecha em 6 segundos

Use sempre que o IP da sua rede mudar (trocou de WiFi, reiniciou roteador etc).
"""
import os
import sys
import socket
import time
import subprocess
from pathlib import Path


PORT_PADRAO = 5000


def _pasta_atual():
    """Retorna a pasta onde o .exe (ou .py) esta rodando."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


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


def _copiar_clipboard(texto):
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


def _abrir_no_notepad(arquivo):
    """Abre o arquivo no Bloco de Notas (sem bloquear)."""
    try:
        subprocess.Popen(["notepad.exe", str(arquivo)], shell=False)
    except Exception:
        pass


def main():
    pasta = _pasta_atual()
    ip = _get_local_ip()
    link = f"http://{ip}:{PORT_PADRAO}/comecar?next=/executar-sistema"

    conteudo = (
        "=== LINK PARA COMPARTILHAR COM SUA EQUIPE ===\n\n"
        f"{link}\n\n"
        "Como usar:\n"
        "  1. Mande este link no WhatsApp / Email / Telegram para sua equipe.\n"
        "  2. Todos devem estar conectados na MESMA rede WiFi deste PC.\n"
        "  3. Mantenha o Planilhas.exe ABERTO neste PC enquanto a equipe usar.\n\n"
        "Se o IP mudar (trocou de WiFi / reiniciou roteador):\n"
        "  - Basta dar duplo clique novamente em Gerar_Link.exe\n"
        "    e este arquivo Link.txt sera atualizado automaticamente.\n\n"
        "CREDENCIAIS PADRAO (para o primeiro login):\n"
        "  Email: admin@planilhas.com\n"
        "  Senha: admin123\n"
    )

    arq = pasta / "Link.txt"
    try:
        arq.write_text(conteudo, encoding="utf-8")
    except Exception as e:
        print(f"[ERRO] Nao foi possivel salvar Link.txt: {e}")
        input("Pressione Enter para sair...")
        return

    _copiar_clipboard(link)

    larg = 64
    print("\n" + "=" * larg)
    print(" >> LINK GERADO COM SUCESSO ".center(larg))
    print("=" * larg)
    print(f"\n  IP detectado neste PC : {ip}")
    print(f"  Link compartilhavel    : {link}\n")
    print(f"  Salvo em               : {arq}")
    print("\n  >> O link ja foi COPIADO para a area de transferencia.")
    print("     Cole no WhatsApp/Email com Ctrl+V.\n")
    print("=" * larg)

    # Abre o Link.txt no Bloco de Notas para visualizacao rapida
    _abrir_no_notepad(arq)

    # Pequena pausa para o usuario ler antes de fechar
    print("\nFechando em 6 segundos...")
    try:
        time.sleep(6)
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print(f"\n[ERRO FATAL] {e}")
        traceback.print_exc()
        input("\nPressione Enter para sair...")
