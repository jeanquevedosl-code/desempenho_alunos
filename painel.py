"""
Painel web do extrator CGD.

Roda no seu PC sem senha no painel e permite que outra pessoa:
  - digite os nomes dos alunos;
  - dispare a geração do relatório (o robô roda aqui, com a sua sessão do CGD);
  - abra/baixe os relatórios prontos.

Uso:
    python painel.py            # só local (http://127.0.0.1:8000)
    python painel_online.py     # local + link público via Cloudflare Tunnel

"""

from __future__ import annotations

import json
import re
import secrets
import subprocess
import sys
import threading
import unicodedata
import hashlib
import os
from dataclasses import asdict
from datetime import datetime, timedelta
from pathlib import Path
from queue import Queue

from flask import (
    Flask,
    abort,
    redirect,
    render_template_string,
    request,
    send_from_directory,
    session,
    url_for,
)
from playwright.sync_api import sync_playwright

BASE = Path(__file__).resolve().parent
RELATORIOS = BASE / "relatorios"
CONFIG_FILE = BASE / "painel_config.json"
SENHA_FILE = BASE / "painel_senha.txt"
JOBS_FILE = BASE / "painel_jobs.json"

MAX_NOMES = 20
MAX_CHARS_NOME = 120
TIMEOUT_JOB = 3600

try:  # console do Windows (cp1252) nao aceita emoji: grava com substituicao
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


# ---------------------------------------------------------------------------
# Configuracao / senha
# ---------------------------------------------------------------------------

def carregar_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            dados = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            if isinstance(dados, dict) and dados.get("secret"):
                dados["exigir_senha"] = False
                return dados
        except Exception:
            pass
    cfg = {"secret": secrets.token_hex(32), "exigir_senha": False}
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    except OSError:
        pass
    return cfg


def senha_atual() -> str:
    if SENHA_FILE.exists():
        senha = SENHA_FILE.read_text(encoding="utf-8").strip()
        if senha:
            return senha
    senha = "cgd-" + secrets.token_hex(3)
    try:
        SENHA_FILE.write_text(senha + "\n", encoding="utf-8")
    except OSError:
        pass
    return senha


def slug(texto: str, limite: int = 40) -> str:
    s = unicodedata.normalize("NFKD", texto or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", s).strip("_")
    return (s or "alunos")[:limite]


app = Flask(__name__)
_CONFIG = carregar_config()
app.secret_key = _CONFIG["secret"]
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    PERMANENT_SESSION_LIFETIME=timedelta(hours=12),
    MAX_CONTENT_LENGTH=64 * 1024,
    JSON_SORT_KEYS=False,
    EXIGIR_SENHA=bool(_CONFIG.get("exigir_senha", False)),
)


@app.after_request
def _cabecalhos(resposta):
    resposta.headers["X-Frame-Options"] = "DENY"
    resposta.headers["X-Content-Type-Options"] = "nosniff"
    resposta.headers["Referrer-Policy"] = "same-origin"
    if request.endpoint in {"login", "painel"}:
        resposta.headers["Cache-Control"] = "no-store"
    return resposta


# ---------------------------------------------------------------------------
# Jobs
# ---------------------------------------------------------------------------

LOCK = threading.Lock()
FILA: Queue[str] = Queue()
JOBS: dict[str, dict] = {}
CREDENTIALS: dict[str, dict[str, str]] = {}


def carregar_jobs() -> None:
    if not JOBS_FILE.exists():
        return
    try:
        dados = json.loads(JOBS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return
    for job in dados:
        if job.get("status") in {"fila", "rodando"}:
            job["status"] = "interrompido"
            job["log"] = (job.get("log") or "") + "\nO painel foi reiniciado antes de terminar."
        JOBS[job["id"]] = job


def salvar_jobs() -> None:
    with LOCK:
        dados = list(JOBS.values())
    try:
        JOBS_FILE.write_text(
            json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        pass


def agora() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def _worker() -> None:
    while True:
        job_id = FILA.get()
        with LOCK:
            job = JOBS.get(job_id)
            if job is None:
                FILA.task_done()
                continue
            job["status"] = "rodando"
            job["inicio"] = agora()
        salvar_jobs()

        try:
            RELATORIOS.mkdir(parents=True, exist_ok=True)
        except OSError:
            job["status"] = "erro"
            job["erro"] = "Geração indisponível nesta hospedagem. Use o painel local."
            salvar_jobs()
            CREDENTIALS.pop(job_id, None)
            FILA.task_done()
            continue
        carimbo = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        saida = RELATORIOS / f"desempenho_{slug(job['nomes'][0])}_{carimbo}.xlsx"
        credenciais = CREDENTIALS.get(job_id, {})
        chave_sessao = hashlib.sha256(
            credenciais.get("email", "").encode("utf-8")
        ).hexdigest()[:16]
        pasta_sessoes = BASE / "sessoes"
        pasta_sessoes.mkdir(parents=True, exist_ok=True)
        session_file = pasta_sessoes / f"{chave_sessao}.json"

        comando = [
            sys.executable,
            str(BASE / "cgd_extrator.py"),
            "--alunos", *job["nomes"],
            "--html", "--json",
            "--sem-login-manual",
            "--saida", str(saida),
            "--session-file", str(session_file),
        ]

        log = ""
        sucesso = False
        try:
            env = os.environ.copy()
            env["CGD_EMAIL"] = credenciais.get("email", "")
            env["CGD_SENHA"] = credenciais.get("senha", "")
            resultado = subprocess.run(
                comando,
                cwd=str(BASE),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=TIMEOUT_JOB,
                env=env,
            )
            log = (resultado.stdout or "") + (resultado.stderr or "")
            sucesso = saida.exists()
            if not sucesso and resultado.returncode != 0:
                log += f"\n(código de saída {resultado.returncode})"
        except subprocess.TimeoutExpired:
            log = f"Tempo esgotado ({TIMEOUT_JOB // 60} min)."
        except Exception as exc:  # noqa: BLE001
            log = f"Falha ao executar o extrator: {exc}"

        html = saida.with_suffix(".html")
        with LOCK:
            job = JOBS.get(job_id) or {}
            job["fim"] = agora()
            job["log"] = "\n".join(log.strip().splitlines()[-80:])
            if sucesso:
                job["status"] = "ok"
                job["arquivo"] = saida.name
                job["html"] = html.name if html.exists() else ""
                job["aviso"] = (
                    "Nenhum contrato encontrado para um dos nomes."
                    if "Nenhum contrato" in log or "nenhum contrato" in log.lower()
                    else ""
                )
            else:
                job["status"] = "erro"
                if "sessão do cgd expirada" in (log or "").lower():
                    job["erro"] = (
                        "Sessão do CGD expirada. No PC do painel, rode "
                        "`python cgd_extrator.py` e entre no site para renovar."
                    )
                elif "nenhum contrato" in (log or "").lower():
                    job["erro"] = "Nenhum contrato encontrado para os nomes informados."
                else:
                    job["erro"] = "O relatório não foi gerado. Veja o log."
        salvar_jobs()
        FILA.task_done()


def iniciar_worker() -> None:
    if getattr(app, "_worker_iniciado", False):
        return
    app._worker_iniciado = True  # type: ignore[attr-defined]
    threading.Thread(target=_worker, daemon=True).start()


carregar_jobs()
iniciar_worker()


# ---------------------------------------------------------------------------
# Acesso
# ---------------------------------------------------------------------------

TENTATIVAS: dict[str, list[float]] = {}


def _bloqueado(ip: str) -> bool:
    import time

    agora_ts = time.time()
    marcas = [t for t in TENTATIVAS.get(ip, []) if agora_ts - t < 300]
    TENTATIVAS[ip] = marcas
    return len(marcas) >= 8


def _registrar_falha(ip: str) -> None:
    import time

    TENTATIVAS.setdefault(ip, []).append(time.time())


def requer_login(funcao):
    from functools import wraps

    @wraps(funcao)
    def wrapper(*args, **kwargs):
        if app.config.get("EXIGIR_SENHA") and not session.get("ok"):
            return redirect(url_for("login"))
        return funcao(*args, **kwargs)

    return wrapper


# ---------------------------------------------------------------------------
# Templates
# ---------------------------------------------------------------------------

ESTILO = """
  :root { color-scheme: light; }
  * { box-sizing: border-box; }
  body { margin: 0; font: 15px/1.55 -apple-system, "Segoe UI", Roboto, Arial, sans-serif;
         color: #1f2430; background: #f4f6f9; }
  header { background: #2F5496; color: #fff; padding: 20px 28px;
           display: flex; align-items: center; justify-content: space-between; gap: 16px; }
  header h1 { margin: 0; font-size: 20px; }
  header p { margin: 2px 0 0; opacity: .85; font-size: 13px; }
  main { max-width: 1000px; margin: 0 auto; padding: 24px; }
  .card { background: #fff; border: 1px solid #e2e5ea; border-radius: 10px; padding: 18px 20px; margin-bottom: 18px; }
  .card h2 { margin: 0 0 10px; font-size: 16px; }
  textarea { width: 100%; min-height: 110px; padding: 10px 12px; border: 1px solid #cfd5df;
             border-radius: 8px; font: 14px/1.5 inherit; resize: vertical; }
  button, .botao { background: #2F5496; color: #fff; border: 0; border-radius: 8px;
                   padding: 9px 16px; font-size: 14px; font-weight: 600; cursor: pointer;
                   text-decoration: none; display: inline-block; }
  button:hover, .botao:hover { background: #274680; }
  .secundario { background: #eef1f7; color: #2F5496; }
  .secundario:hover { background: #e2e7f2; }
  .dica { color: #5b6472; font-size: 13px; margin: 8px 0 0; }
  table { width: 100%; border-collapse: collapse; font-size: 14px; }
  th { text-align: left; font-size: 12px; text-transform: uppercase; letter-spacing: .04em;
       color: #3c4658; padding: 8px; border-bottom: 1px solid #e2e5ea; }
  td { padding: 10px 8px; border-bottom: 1px solid #eef0f4; vertical-align: middle; }
  .badge { display: inline-block; padding: 2px 9px; border-radius: 999px; font-size: 12px; font-weight: 600; }
  td.acoes { white-space: nowrap; }
  td.acoes .botao { margin: 2px 4px 2px 0; }
  th.acoes, td.quando { white-space: nowrap; }
  .badge.ok { background: #d8f3e0; color: #14663a; }
  .badge.rodando { background: #dbe7fb; color: #1d3f7a; }
  .badge.fila { background: #fdf1c7; color: #755b00; }
  .badge.erro { background: #fbdada; color: #8a1f1f; }
  .badge.interrompido { background: #eceef2; color: #5b6472; }
  details summary { cursor: pointer; color: #2F5496; font-size: 13px; }
  pre { background: #0f172a; color: #d7e2f3; padding: 12px; border-radius: 8px;
        overflow: auto; max-height: 260px; font-size: 12px; }
  .vazio { color: #7a8290; font-size: 14px; }
  .login { max-width: 360px; margin: 12vh auto; }
  .login h1 { font-size: 18px; margin: 0 0 4px; }
  .login p { color: #5b6472; font-size: 13px; margin: 0 0 16px; }
  .login input { width: 100%; padding: 10px 12px; border: 1px solid #cfd5df; border-radius: 8px;
                 font-size: 15px; margin-bottom: 12px; }
  .erro-box { background: #fdf2f2; border: 1px solid #f3caca; color: #8a1f1f;
              padding: 10px 12px; border-radius: 8px; font-size: 14px; margin-bottom: 12px; }
"""

LOGIN_HTML = """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Entrar — Desempenho CGD</title><style>""" + ESTILO + """</style></head>
<body>
<div class="login">
  <div class="card">
    <h1>Desempenho de alunos</h1>
    <p>Área restrita. Informe a senha do painel.</p>
    {% if erro %}<div class="erro-box">{{ erro }}</div>{% endif %}
    <form method="post" action="{{ url_for('login') }}">
      <input type="password" name="senha" placeholder="Senha" autofocus autocomplete="current-password">
      <button type="submit" style="width:100%">Entrar</button>
    </form>
  </div>
</div>
</body></html>
"""

PAINEL_HTML = """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Desempenho de alunos</title><style>""" + ESTILO + """</style></head>
<body>
<header>
  <div>
    <h1>Desempenho de alunos</h1>
    <p>Gere o relatório digitando os nomes — a coleta roda neste computador.</p>
  </div>
  {% if com_senha %}
  <form method="post" action="{{ url_for('sair') }}">
    <button class="secundario" type="submit">Sair</button>
  </form>
  {% endif %}
</header>
<main>
  <div class="card">
    <h2>Novo relatório</h2>
    <form method="post" action="{{ url_for('gerar') }}">
      <input type="email" name="cgd_email" placeholder="E-mail do CGD" autocomplete="email" required>
      <input type="password" name="cgd_senha" placeholder="Senha do CGD" autocomplete="current-password" required>
      <textarea name="nomes" placeholder="Um nome por linha&#10;Ex.:&#10;Pedro Henrique Almeida de Oliveira&#10;Maria Souza Lima">{{ nomes }}</textarea>
      <p class="dica">Aceita até {{ max_nomes }} nomes por vez. Cada aluno leva alguns segundos.</p>
      <button type="submit">Gerar relatório</button>
    </form>
  </div>

  <div class="card">
    <h2>Gerados agora</h2>
    <div id="jobs">{{ jobs|safe }}</div>
  </div>

  <div class="card">
    <h2>Relatórios salvos</h2>
    <div id="salvos">{{ salvos|safe }}</div>
  </div>
</main>
<script>
async function atualizar() {
  try {
    const r = await fetch("{{ url_for('api_jobs') }}", {cache: "no-store"});
    if (!r.ok) return;
    const dados = await r.json();
    document.getElementById("jobs").innerHTML = dados.jobs;
    document.getElementById("salvos").innerHTML = dados.salvos;
    if (dados.ativo) setTimeout(atualizar, 3000);
  } catch (e) {}
}
setTimeout(atualizar, 2500);
</script>
</body></html>
"""

JOBS_HTML = """
{% if jobs %}
<table>
  <thead><tr><th class="quando">Quando</th><th>Alunos</th><th>Status</th><th class="acoes">Arquivos</th></tr></thead>
  <tbody>
  {% for job in jobs %}
    <tr>
      <td class="quando">{{ job.criado }}</td>
      <td>{{ job.nomes|join(", ") }}</td>
      <td><span class="badge {{ job.status }}">{{ rotulo.get(job.status, job.status) }}</span>
        {% if job.aviso %}<div class="dica">{{ job.aviso }}</div>{% endif %}</td>
      <td class="acoes">
        {% if job.status == 'ok' %}
          {% if job.html %}<a class="botao secundario" href="{{ url_for('ver_relatorio', nome=job.html) }}" target="_blank">Abrir relatório</a>{% endif %}
          {% if job.html %}<a class="botao secundario" href="{{ url_for('editar_observacoes', nome=job.html) }}">Observações</a>{% endif %}
          {% if job.html %}<a class="botao" href="{{ url_for('baixar_pdf', nome=job.html) }}">Baixar PDF</a>{% endif %}
          <a class="botao" href="{{ url_for('baixar', nome=job.arquivo) }}">Excel</a>
        {% elif job.status == 'erro' %}
          <span class="dica">{{ job.erro or "Falhou" }}</span>
        {% else %}
          <span class="dica">aguarde…</span>
        {% endif %}
        {% if job.log %}<details><summary>log</summary><pre>{{ job.log }}</pre></details>{% endif %}
      </td>
    </tr>
  {% endfor %}
  </tbody>
</table>
{% else %}
<p class="vazio">Nenhum relatório gerado nesta sessão.</p>
{% endif %}
"""

SALVOS_HTML = """
{% if arquivos %}
<table>
  <thead><tr><th>Arquivo</th><th class="quando">Gerado em</th><th class="acoes"></th></tr></thead>
  <tbody>
  {% for a in arquivos %}
    <tr>
      <td>{{ a.titulo }}</td>
      <td class="quando">{{ a.data }}</td>
      <td class="acoes">
        {% if a.html %}<a class="botao secundario" href="{{ url_for('ver_relatorio', nome=a.html) }}" target="_blank">Abrir</a>{% endif %}
        {% if a.html %}<a class="botao secundario" href="{{ url_for('editar_observacoes', nome=a.html) }}">Observações</a>{% endif %}
        {% if a.html %}<a class="botao" href="{{ url_for('baixar_pdf', nome=a.html) }}">PDF</a>{% endif %}
        {% if a.xlsx %}<a class="botao secundario" href="{{ url_for('baixar', nome=a.xlsx) }}">Excel</a>{% endif %}
      </td>
    </tr>
  {% endfor %}
  </tbody>
</table>
{% else %}
<p class="vazio">Nenhum relatório salvo ainda.</p>
{% endif %}
"""

ROTULOS = {
    "fila": "Na fila",
    "rodando": "Gerando…",
    "ok": "Pronto",
    "erro": "Erro",
    "interrompido": "Interrompido",
}


# ---------------------------------------------------------------------------
# Renderizacao
# ---------------------------------------------------------------------------

def _titulo_salvo(nome: str) -> str:
    """Nome legivel para a lista: 'Fulano de Tal' (ou a data, se nao houver nome)."""
    texto = nome
    if texto.startswith("desempenho_"):
        texto = texto[len("desempenho_"):]
    achado = re.match(r"^(.*?)_(\d{4}-\d{2}-\d{2})_(\d{6})$", texto)
    if achado:
        nomes, data, hora = achado.groups()
        quando = f"{data[8:10]}/{data[5:7]}/{data[:4]} {hora[:2]}:{hora[2:4]}"
        bonito = nomes.replace("_", " ").strip()
        return bonito or quando
    return texto.replace("_", " ").strip()


def listar_salvos() -> list[dict]:
    if not RELATORIOS.exists():
        return []
    itens = []
    vistos = set()
    arquivos = sorted(
        (p for p in RELATORIOS.iterdir() if p.is_file() and p.suffix in {".html", ".xlsx"}),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    for arquivo in arquivos:
        base = arquivo.with_suffix("")
        if base.name in vistos:
            continue
        vistos.add(base.name)
        html = arquivo.with_suffix(".html")
        xlsx = arquivo.with_suffix(".xlsx")
        itens.append(
            {
                "titulo": _titulo_salvo(base.name),
                "data": datetime.fromtimestamp(arquivo.stat().st_mtime).strftime("%d/%m/%Y %H:%M"),
                "html": html.name if html.exists() else "",
                "xlsx": xlsx.name if xlsx.exists() else "",
            }
        )
    return itens[:30]


def render_jobs() -> str:
    with LOCK:
        jobs = sorted(
            (dict(j) for j in JOBS.values()),
            key=lambda j: j.get("criado", ""),
            reverse=True,
        )
    return render_template_string(JOBS_HTML, jobs=jobs[:15], rotulo=ROTULOS)


def render_salvos() -> str:
    return render_template_string(SALVOS_HTML, arquivos=listar_salvos())


def tem_job_ativo() -> bool:
    with LOCK:
        return any(j.get("status") in {"fila", "rodando"} for j in JOBS.values())


# ---------------------------------------------------------------------------
# Rotas
# ---------------------------------------------------------------------------

@app.get("/login")
def login():
    if not app.config.get("EXIGIR_SENHA"):
        return redirect(url_for("painel"))
    return render_template_string(LOGIN_HTML, erro=None)


@app.post("/login")
def login_post():
    if not app.config.get("EXIGIR_SENHA"):
        return redirect(url_for("painel"))

    ip = request.headers.get("CF-Connecting-IP") or request.remote_addr or "?"
    if _bloqueado(ip):
        return render_template_string(LOGIN_HTML, erro="Muitas tentativas. Aguarde 5 minutos."), 429

    senha = (request.form.get("senha") or "").strip()
    if senha and secrets.compare_digest(senha, senha_atual()):
        session.clear()
        session.permanent = True
        session["ok"] = True
        return redirect(url_for("painel"))

    _registrar_falha(ip)
    return render_template_string(LOGIN_HTML, erro="Senha incorreta."), 401


@app.post("/sair")
def sair():
    session.clear()
    return redirect(url_for("login"))


@app.get("/")
@requer_login
def painel():
    return render_template_string(
        PAINEL_HTML,
        nomes="",
        max_nomes=MAX_NOMES,
        com_senha=app.config.get("EXIGIR_SENHA"),
        jobs=render_jobs(),
        salvos=render_salvos(),
    )


@app.post("/gerar")
@requer_login
def gerar():
    texto = request.form.get("nomes") or ""
    nomes = []
    for linha in texto.splitlines():
        nome = re.sub(r"\s+", " ", linha).strip()
        if not nome:
            continue
        if len(nome) > MAX_CHARS_NOME or len(nome) < 3:
            continue
        if nome not in nomes:
            nomes.append(nome)

    if not nomes:
        return redirect(url_for("painel"))
    nomes = nomes[:MAX_NOMES]
    cgd_email = (request.form.get("cgd_email") or "").strip()
    cgd_senha = (request.form.get("cgd_senha") or "").strip()
    if not cgd_email or not cgd_senha:
        return redirect(url_for("painel"))

    job_id = secrets.token_hex(4)
    CREDENTIALS[job_id] = {"email": cgd_email, "senha": cgd_senha}
    with LOCK:
        JOBS[job_id] = {
            "id": job_id,
            "nomes": nomes,
            "status": "fila",
            "criado": agora(),
            "inicio": "",
            "fim": "",
            "arquivo": "",
            "html": "",
            "log": "",
            "aviso": "",
            "erro": "",
        }
    salvar_jobs()
    FILA.put(job_id)
    return redirect(url_for("painel"))


@app.get("/api/jobs")
@requer_login
def api_jobs():
    return {"jobs": render_jobs(), "salvos": render_salvos(), "ativo": tem_job_ativo()}


def _arquivo_seguro(nome: str) -> Path:
    caminho = (RELATORIOS / nome).resolve()
    if caminho.parent != RELATORIOS.resolve() or not caminho.is_file():
        abort(404)
    if caminho.suffix.lower() not in {".xlsx", ".html"}:
        abort(404)
    return caminho


@app.get("/ver/<path:nome>")
@requer_login
def ver_relatorio(nome: str):
    html = _arquivo_seguro(nome)
    documento = html.read_text(encoding="utf-8")
    nome_base = html.stem
    barra = f"""
<nav class="acoes-relatorio" aria-label="Acoes do relatorio"
     style="display:flex;gap:8px;align-items:center;margin:14px 32px 0">
  <a href="{url_for('editar_observacoes', nome=html.name)}"
     style="background:#fff;color:#2F5496;border:1px solid #cfd5df;border-radius:8px;padding:8px 12px;text-decoration:none;font-weight:600">Observações</a>
  <a href="{url_for('baixar', nome=nome_base + '.xlsx')}"
     style="background:#fff;color:#2F5496;border:1px solid #cfd5df;border-radius:8px;padding:8px 12px;text-decoration:none;font-weight:600">Baixar Excel</a>
  <a href="{url_for('baixar_pdf', nome=html.name)}"
     style="background:#fff;color:#2F5496;border:1px solid #cfd5df;border-radius:8px;padding:8px 12px;text-decoration:none;font-weight:600">Baixar PDF</a>
</nav>
"""
    if "</header>" in documento:
        documento = documento.replace("</header>", "</header>" + barra, 1)
    else:
        documento = documento.replace("<body>", "<body>" + barra, 1)
    resposta = app.make_response(documento)
    resposta.headers["Content-Type"] = "text/html; charset=utf-8"
    resposta.headers["Cache-Control"] = "no-store"
    return resposta


@app.get("/baixar/<path:nome>")
@requer_login
def baixar(nome: str):
    _arquivo_seguro(nome)
    return send_from_directory(RELATORIOS, nome, as_attachment=True)


OBSERVACOES_HTML = """
<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Observações — Desempenho CGD</title><style>""" + ESTILO + """
.observacao { width: 100%; min-height: 72px; }
.curso { margin-bottom: 16px; padding-bottom: 16px; border-bottom: 1px solid #eef0f4; }
.curso h3 { margin: 0 0 4px; font-size: 15px; }
.curso p { margin: 0 0 8px; color: #5b6472; font-size: 13px; }
</style></head><body>
<header><div><h1>Observações do relatório</h1><p>{{ titulo }}</p></div></header>
<main><div class="card">
<form method="post">
{% for item in cursos %}
<div class="curso">
  <h3>{{ item.aluno }} — {{ item.curso }}</h3>
  <textarea class="observacao" name="observacao_{{ loop.index0 }}" placeholder="Digite uma observação">{{ item.observacao }}</textarea>
</div>
{% endfor %}
<button type="submit">Salvar observações</button>
<a class="botao secundario" href="{{ url_for('ver_relatorio', nome=nome_html) }}">Voltar ao relatório</a>
</form></div></main></body></html>
"""


def _dados_relatorio(nome_html: str) -> tuple[Path, Path, dict]:
    html = _arquivo_seguro(nome_html)
    if html.suffix.lower() != ".html":
        abort(404)
    dados = html.with_suffix(".json")
    if not dados.is_file():
        abort(404)
    try:
        payload = json.loads(dados.read_text(encoding="utf-8"))
    except Exception:
        abort(404)
    if not isinstance(payload, dict) or not isinstance(payload.get("cursos"), list):
        abort(404)
    return html, dados, payload


@app.get("/observacoes/<path:nome>")
@requer_login
def editar_observacoes(nome: str):
    html, _, payload = _dados_relatorio(nome)
    return render_template_string(
        OBSERVACOES_HTML,
        titulo=html.stem.replace("_", " "),
        nome_html=html.name,
        cursos=payload["cursos"],
    )


@app.post("/observacoes/<path:nome>")
@requer_login
def salvar_observacoes_relatorio(nome: str):
    html, dados, payload = _dados_relatorio(nome)
    from cgd_extrator import Curso, salvar_excel, salvar_html

    cursos = []
    for indice, item in enumerate(payload["cursos"]):
        curso = Curso(**item)
        curso.observacao = (request.form.get(f"observacao_{indice}") or "").strip()
        cursos.append(curso)

    payload["cursos"] = [asdict(curso) for curso in cursos]
    dados.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    salvar_html(cursos, payload.get("erros", []), str(html))
    pdf = html.with_suffix(".pdf")
    if pdf.exists():
        pdf.unlink()
    xlsx = html.with_suffix(".xlsx")
    salvar_excel(cursos, payload.get("erros", []), str(xlsx))
    return redirect(url_for("ver_relatorio", nome=html.name))


@app.get("/pdf/<path:nome>")
@requer_login
def baixar_pdf(nome: str):
    html, _, _ = _dados_relatorio(nome)
    pdf = html.with_suffix(".pdf")
    if not pdf.is_file():
        try:
            with sync_playwright() as playwright:
                browser = playwright.chromium.launch(headless=True)
                page = browser.new_page()
                page.goto(html.as_uri(), wait_until="load")
                page.pdf(
                    path=str(pdf),
                    format="A4",
                    landscape=True,
                    print_background=True,
                    margin={"top": "12mm", "right": "8mm", "bottom": "12mm", "left": "8mm"},
                )
                browser.close()
        except Exception as exc:
            abort(500, description=f"Não foi possível gerar o PDF: {exc}")
    return send_from_directory(RELATORIOS, pdf.name, as_attachment=True)


@app.get("/saude")
def saude():
    return {"ok": True, "hora": agora()}


# ---------------------------------------------------------------------------
# Execucao
# ---------------------------------------------------------------------------

def main() -> None:
    print("=" * 58)
    print("  Painel de desempenho CGD")
    print("=" * 58)
    print("  Acesso ao painel: sem senha")
    print("  Local: http://127.0.0.1:8000")
    print("  Para publicar: python painel_online.py")
    print("=" * 58)

    from waitress import serve

    serve(app, host="127.0.0.1", port=8000, threads=8, ident="cgd-painel")


if __name__ == "__main__":
    main()
