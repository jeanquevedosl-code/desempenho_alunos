"""
Publica o painel (painel.py) na internet usando Cloudflare Tunnel.

O painel e o robo continuam rodando NESTE computador; o Cloudflare apenas abre
um tunel de saida, sem precisar mexer no roteador nem abrir portas.

Uso:
    python painel_online.py

O link publico e temporario (muda a cada vez que este script e iniciado) e fica
salvo em painel_url.txt. O painel principal funciona sem senha.
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from pathlib import Path

BASE = Path(__file__).resolve().parent
CLOUDFLARED = BASE / "tools" / "cloudflared.exe"
URL_FILE = BASE / "painel_url.txt"
PID_FILE = BASE / "painel_tunnel.pid"
RE_URL = re.compile(r"https://[a-z0-9-]+\.trycloudflare\.com")

try:  # console do Windows (cp1252) nao aceita emoji: grava com substituicao
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def _encerrar_tunel_antigo() -> None:
    """Mata o cloudflared de uma execucao anterior que ficou orfao."""
    if not PID_FILE.exists():
        return
    try:
        pid = int(PID_FILE.read_text(encoding="utf-8").strip() or "0")
    except Exception:
        pid = 0
    if pid:
        try:
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/F"],
                capture_output=True,
                timeout=15,
            )
        except Exception:
            pass
    try:
        PID_FILE.unlink()
    except Exception:
        pass


def _aguardar_servidor(porta: int, timeout: float = 20.0) -> bool:
    limite = time.time() + timeout
    while time.time() < limite:
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{porta}/saude", timeout=2):
                return True
        except Exception:
            time.sleep(0.4)
    return False


def _subir_cloudflared(porta: int) -> tuple[subprocess.Popen, str]:
    if not CLOUDFLARED.exists():
        raise SystemExit(
            f"❌ Não encontrei {CLOUDFLARED}.\n"
            "   Baixe com:\n"
            '   Invoke-WebRequest "https://github.com/cloudflare/cloudflared/releases/'
            'latest/download/cloudflared-windows-amd64.exe" -OutFile "tools\\cloudflared.exe"'
        )

    _encerrar_tunel_antigo()

    processo = subprocess.Popen(
        [
            str(CLOUDFLARED),
            "--no-autoupdate",
            "tunnel",
            "--url", f"http://127.0.0.1:{porta}",
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )

    link = ""
    limite = time.time() + 90
    while time.time() < limite:
        linha = processo.stdout.readline() if processo.stdout else ""
        if not linha:
            if processo.poll() is not None:
                break
            time.sleep(0.2)
            continue
        if not link:
            achado = RE_URL.search(linha)
            if achado:
                link = achado.group(0)
                break
    return processo, link


def main() -> None:
    parser = argparse.ArgumentParser(description="Publica o painel via Cloudflare Tunnel.")
    parser.add_argument("--porta", type=int, default=8000, help="porta local do painel (padrao: 8000)")
    parser.add_argument("--sem-navegador", action="store_true", help="não abrir o link no navegador")
    args = parser.parse_args()

    import painel  # inicia o worker da fila
    from waitress import serve

    def servidor() -> None:
        serve(painel.app, host="127.0.0.1", port=args.porta, threads=8, ident="cgd-painel")

    threading.Thread(target=servidor, daemon=True).start()

    if not _aguardar_servidor(args.porta):
        raise SystemExit(f"❌ O painel não respondeu em http://127.0.0.1:{args.porta} (porta ocupada?).")

    print("=" * 62)
    print("  Painel de desempenho CGD — publicado")
    print("=" * 62)
    print("  Acesso ao painel: sem senha")
    print("  Abrindo túnel do Cloudflare...")

    processo, link = _subir_cloudflared(args.porta)

    if not link:
        try:
            processo.terminate()
        except Exception:
            pass
        raise SystemExit("❌ O Cloudflare não devolveu um link. Verifique a internet e tente de novo.")

    URL_FILE.write_text(link + "\n", encoding="utf-8")
    try:
        PID_FILE.write_text(str(processo.pid), encoding="utf-8")
    except Exception:
        pass
    print("-" * 62)
    print(f"  LINK PÚBLICO: {link}")
    print(f"  (salvo em {URL_FILE.name} — vale enquanto este script estiver aberto)")
    print("-" * 62)
    print("  Deixe esta janela aberta. Ctrl+C encerra o painel e o link.")
    print("=" * 62)

    if not args.sem_navegador:
        try:
            webbrowser.open(link)
        except Exception:
            pass

    try:
        while True:
            if processo.poll() is not None:
                print("⚠️  O túnel do Cloudflare caiu. Reinicie o script.")
                break
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n⏹️  Encerrando...")
    finally:
        try:
            processo.terminate()
        except Exception:
            pass
        for arquivo in (URL_FILE, PID_FILE):
            try:
                arquivo.unlink()
            except Exception:
                pass


if __name__ == "__main__":
    main()
