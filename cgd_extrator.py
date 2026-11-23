"""
CGD - Extrator de desempenho de alunos (Playwright)

Uso:
    python cgd_extrator.py                            # modo interativo
    python cgd_extrator.py --alunos "Ana Silva" "Joao Souza"
    python cgd_extrator.py --lista alunos.txt --saida relatorio.xlsx
    python cgd_extrator.py --alunos "Ana" --debug-html --nova-sessao

Requisitos:
    pip install playwright openpyxl
    playwright install chromium

O login e feito MANUALMENTE na primeira execucao (Cloudflare).
A sessao fica salva em cgd_session.json e e reaproveitada nas proximas.
"""

from __future__ import annotations

import argparse
import difflib
import html
import json
import os
import re
import sys
import time
import unicodedata
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, urljoin

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

BASE_URL = "https://app.cgd.com.br"
SESSION_FILE = Path("cgd_session.json")
TIMEOUT_NAV = 45000
TIMEOUT_LOGIN = 300000  # 5 minutos para o login manual

# Evita crash de UnicodeEncodeError quando a saida e redirecionada (ex.: > log.txt).
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------

def limpar(texto: str | None) -> str:
    """Remove quebras de linha e espacos repetidos."""
    return re.sub(r"\s+", " ", texto or "").strip()


def normalizar(texto: str | None) -> str:
    """minusculas, sem acentos, sem pontuacao e com espacos unicos."""
    s = unicodedata.normalize("NFKD", texto or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def nome_arquivo_seguro(texto: str, limite: int = 60) -> str:
    s = unicodedata.normalize("NFKD", texto or "")
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", s).strip("_")
    return (s or "aluno")[:limite]


def combina_nome(nome: str, texto: str) -> bool:
    """True se todos os tokens do nome aparecem no texto (ignora ordem/acentos)."""
    tokens = [t for t in normalizar(nome).split() if len(t) > 1]
    if not tokens:
        return False
    alvo = normalizar(texto)
    return all(t in alvo for t in tokens)


def esperar(page, segundos: float) -> None:
    try:
        page.wait_for_timeout(int(segundos * 1000))
    except Exception:
        pass


def ir_para(page, url: str, tentativas: int = 2) -> bool:
    """Navega com retry. Retorna False se nao conseguiu abrir."""
    for tentativa in range(tentativas):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=TIMEOUT_NAV)
            esperar(page, 1.5)
            return True
        except Exception as exc:  # noqa: BLE001
            if tentativa == tentativas - 1:
                print(f"   ⚠️  Falha ao abrir {url}: {exc}")
                return False
            esperar(page, 1.0)
    return False


# ---------------------------------------------------------------------------
# Modelo de dados
# ---------------------------------------------------------------------------

@dataclass
class Curso:
    aluno: str
    contrato_id: str
    curso: str
    tipo: str = ""
    inicio: str = ""
    expiracao: str = ""
    ultimo_acesso: str = ""
    email: str = ""
    status: str = ""
    secao: str = ""
    desempenho_url: str = ""
    # dados extraidos da pagina de desempenho
    progresso: str = ""
    itens: str = ""
    nota: str = ""
    situacao: str = ""
    acessos: str = ""
    tempo: str = ""
    # comparativo de carga horaria (preenchido a partir de carga_horaria.csv)
    carga_ideal: str = ""
    cumprimento: str = ""
    diferenca: str = ""
    situacao_carga: str = ""
    tempo_horas: float = 0.0
    carga_horas: float = 0.0
    # observacao escrita pelo professor (painel / observacoes.json)
    observacao: str = ""
    detalhes: list = field(default_factory=list)


# ---------------------------------------------------------------------------
# Navegador / login
# ---------------------------------------------------------------------------

def iniciar_navegador(p, headless: bool, lento: int, renovar_sessao: bool, session_file: Path):
    args = ["--disable-blink-features=AutomationControlled", "--start-maximized"]

    browser = None
    ultimo_erro = None
    # Tenta o Chrome instalado (mais parecido com o navegador real, ajuda no Cloudflare).
    for canal in ("chrome", None):
        try:
            opcoes = {"headless": headless, "slow_mo": lento, "args": args}
            if canal:
                opcoes["channel"] = canal
            browser = p.chromium.launch(**opcoes)
            break
        except Exception as exc:  # noqa: BLE001
            ultimo_erro = exc

    if browser is None:
        raise RuntimeError(f"Não foi possível abrir o navegador: {ultimo_erro}")

    contexto_opcoes = {}
    if session_file.exists() and not renovar_sessao:
        contexto_opcoes["storage_state"] = str(session_file)
    if not headless:
        contexto_opcoes["no_viewport"] = True

    context = browser.new_context(**contexto_opcoes)
    context.set_default_timeout(30000)
    page = context.new_page()
    return browser, context, page


def _estado_pagina(page) -> str:
    """Classifica a pagina: '' (ok), 'cloudflare', 'negado' ou 'vazia'."""
    try:
        titulo = (page.title() or "").lower()
        texto = page.locator("body").inner_text() or ""
    except Exception:
        return "carregando"
    if "just a moment" in titulo or "verificando" in titulo or "attention required" in titulo:
        return "cloudflare"
    if "acesso negado" in titulo or "acesso negado" in texto.lower():
        return "negado"
    if len(texto.strip()) < 40:
        return "vazia"
    return ""


def _esperar_app(page, timeout_ms: int, intervalo: float = 1.5) -> bool:
    """Espera ate a aplicacao carregar de verdade (sai do Cloudflare/login)."""
    limite = time.time() + timeout_ms / 1000
    while time.time() < limite:
        if not _estado_pagina(page) and "/login" not in (page.url or "").lower():
            return True
        esperar(page, intervalo)
    return False


def garantir_login(
    context,
    page,
    permitir_login_manual: bool = True,
    email: str = "",
    senha: str = "",
    session_file: Path = SESSION_FILE,
) -> None:
    """Reaproveita a sessao salva; se estiver expirada, espera o login manual."""
    # /contratos responde vazio mesmo logado; /dashboard é a rota confiável.
    ir_para(page, f"{BASE_URL}/dashboard", tentativas=1)

    # Cloudflare pode exigir alguns segundos antes de liberar.
    if _esperar_app(page, timeout_ms=20000):
        print("✅ Sessão válida, seguindo sem login manual.")
        return

    if not permitir_login_manual and not (email and senha):
        estado = _estado_pagina(page) or "sem resposta"
        raise SystemExit(
            "❌ Sessão do CGD expirada ou bloqueada "
            f"({estado}). Rode o extrator no PC do painel para entrar de novo."
        )

    print("\n⚠️  Login do CGD necessário.")
    print("   O navegador será usado para autenticar esta conta.\n")
    try:
        page.goto(f"{BASE_URL}/login", wait_until="domcontentloaded", timeout=TIMEOUT_NAV)
    except Exception:
        pass

    # Aviso dentro da propria pagina de login.
    try:
        page.evaluate(
            "() => { if (document.getElementById('aviso-robo')) return;"
            " const d = document.createElement('div'); d.id = 'aviso-robo';"
            " d.textContent = 'Robô do extrator: faça login normalmente. "
            "O processamento continua sozinho depois.';"
            " d.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:2147483647;"
            "background:#2F5496;color:#fff;padding:10px;font:14px sans-serif;text-align:center';"
            " document.body.appendChild(d); }"
        )
    except Exception:
        pass

    if email and senha:
        try:
            email_locator = page.locator(
                "input[type=email], input[name*=email i], input[name*=usuario i], "
                "input[name*=login i], input[placeholder*=email i]"
            ).first
            senha_locator = page.locator(
                "input[type=password], input[name*=senha i], input[name*=password i]"
            ).first
            if email_locator.count() and senha_locator.count():
                email_locator.fill(email)
                senha_locator.fill(senha)
                submit = page.locator(
                    "button[type=submit], input[type=submit]"
                ).first
                if submit.count():
                    submit.click()
                else:
                    senha_locator.press("Enter")
                print("   Credenciais enviadas; aguardando validação do CGD...")
            else:
                print("   Campos de login não encontrados; faça o login manualmente.")
        except Exception as exc:
            print(f"   Não foi possível preencher o login automaticamente: {exc}")

    print("   ⏳ Aguardando login (até 5 minutos)...")
    if not _esperar_app(page, timeout_ms=TIMEOUT_LOGIN):
        estado = _estado_pagina(page) or "login não concluído"
        raise SystemExit(f"❌ Tempo esgotado ({estado}). Rode novamente.")

    esperar(page, 1.5)
    try:
        context.storage_state(path=str(session_file))
        print("✅ Login detectado! Sessão salva com segurança no servidor.")
    except Exception:
        print("✅ Login detectado!")


def revelar_abas(page) -> None:
    """Tenta abrir abas/colapsaveis que escondem a tabela 'Fazer'."""
    for seletor in (
        "a:has-text('Fazer')",
        "button:has-text('Fazer')",
        "[role='tab']:has-text('Fazer')",
    ):
        try:
            loc = page.locator(seletor).first
            if loc.count() and loc.is_visible():
                texto = limpar(loc.inner_text()).lower()
                if texto in {"fazer", "fazer cursos", "cursos a fazer", "a fazer"}:
                    loc.click(timeout=2500)
                    esperar(page, 1.0)
                    return
        except Exception:
            continue


# ---------------------------------------------------------------------------
# Busca de contratos
# ---------------------------------------------------------------------------

def buscar_contratos(page, nome: str) -> list[str]:
    """Busca o aluno em /alunos e devolve os ids de contrato encontrados."""
    if not ir_para(page, f"{BASE_URL}/alunos?q={quote(nome)}"):
        return []

    estado = _estado_pagina(page)
    if estado:
        print(f"   ❌ A página de alunos não carregou ({estado}).")
        if estado in {"negado", "vazia"}:
            print("   💡 Sessão expirada: rode com --nova-sessao e faça login de novo.")
        return []

    links = page.locator("a[href*='/contratos/']")
    try:
        total = links.count()
    except Exception:
        total = 0

    achados: dict[str, str] = {}
    for i in range(total):
        link = links.nth(i)
        href = link.get_attribute("href") or ""
        achado = re.search(r"/contratos/(?:cursos-ead/)?(\d+)", href)
        if not achado:
            continue
        try:
            texto_linha = link.evaluate(
                "el => { const tr = el.closest('tr'); return tr ? tr.innerText : (el.innerText || ''); }"
            )
        except Exception:
            texto_linha = ""
        achados.setdefault(achado.group(1), texto_linha)

    if not achados:
        print("   ⚠️  A busca não retornou nenhum contrato para esse nome.")
        return []

    confirmados = [cid for cid, txt in achados.items() if combina_nome(nome, txt)]
    if confirmados:
        return sorted(confirmados)

    # A tabela pode nao mostrar o nome do aluno (so curso/datas).
    if len(achados) <= 15:
        print(
            f"   ⚠️  Nome não aparece nas linhas; usando {len(achados)} contrato(s) "
            "do resultado da busca."
        )
        return sorted(achados)

    print(
        f"   ⚠️  Busca retornou {len(achados)} contratos sem confirmar o nome. "
        "Ajuste o nome digitado (mais completo) e tente de novo."
    )
    return []


# ---------------------------------------------------------------------------
# Pagina do contrato (cursos)
# ---------------------------------------------------------------------------

POSICIONAL = {
    "curso": 0,
    "tipo": 1,
    "inicio": 2,
    "expiracao": 3,
    "acesso": 4,
    "email": 5,
    "status": 6,
}


def _headers_da_tabela(tabela) -> list[str]:
    headers: list[str] = []
    try:
        ths = tabela.locator("thead th")
        n = ths.count()
        if n:
            headers = [limpar(ths.nth(j).inner_text()) for j in range(n)]
    except Exception:
        pass
    return headers


def _achar_coluna(headers: list[str], *termos: str) -> int | None:
    """Casa cabecalhos por palavra (evita 'acao' casar com 'expiracao')."""
    for i, header in enumerate(headers):
        hn = normalizar(header)
        for termo in termos:
            if re.search(rf"\b{re.escape(termo)}", hn):
                return i
    return None


def _mapear_colunas(headers: list[str]) -> dict[str, int | None]:
    return {
        "curso": _achar_coluna(headers, "curso", "disciplina", "treinamento", "conteudo"),
        "tipo": _achar_coluna(headers, "tipo", "modalidade", "categoria"),
        "inicio": _achar_coluna(headers, "inicio", "iniciado", "abertura", "matricula"),
        "expiracao": _achar_coluna(headers, "expira", "validade", "vencimento", "termino", "fim"),
        "acesso": _achar_coluna(headers, "acesso", "ultimo"),
        "email": _achar_coluna(headers, "email", "mail"),
        "status": _achar_coluna(headers, "status", "situa", "acoes", "acao"),
    }


def _pegar(valores: list[str], mapa: dict[str, int | None], campo: str) -> str:
    indice = mapa.get(campo)
    if indice is None:
        indice = POSICIONAL.get(campo)
    if indice is None or indice >= len(valores):
        return ""
    return valores[indice]


def _status_da_linha(texto: str, secao: str) -> str:
    t = normalizar(texto)
    if secao == "Fazer":
        return "Pendente"

    # 1) status escrito na propria celula (badge/texto)
    if re.search(r"\bexpirad", t):
        return "Expirado"
    if re.search(r"\bsuspens", t):
        return "Suspenso"
    if re.search(r"\bbloquead", t):
        return "Bloqueado"
    if re.search(r"\bativ[oa]\b", t):
        return "Ativo"
    if re.search(r"\bconcluid|\bfinalizad", t):
        return "Concluído"

    # 2) inferencia pelo link de acao disponivel:
    #    "Desbloquear" => aluno bloqueado | "Bloquear" => aluno ativo.
    if "desbloquear" in t:
        return "Bloqueado"
    if "bloquear" in t:
        return "Ativo"
    return "Em andamento"


def coletar_cursos(page, aluno: str, contrato_id: str) -> list[Curso]:
    cursos: list[Curso] = []
    tabelas = page.locator("table")
    try:
        total_tabelas = tabelas.count()
    except Exception:
        total_tabelas = 0

    for ti in range(total_tabelas):
        tabela = tabelas.nth(ti)
        try:
            if not tabela.is_visible():
                continue
        except Exception:
            continue

        try:
            meta = tabela.evaluate(
                "el => { const p = el.closest('.tab-pane');"
                " return {id: (el.id || ''), pane: (p && p.id) || ''}; }"
            )
        except Exception:
            meta = {"id": "", "pane": ""}
        marcador = f"{meta.get('id', '')} {meta.get('pane', '')}".lower()
        secao = "Fazer" if "fazer" in marcador else "Andamento"

        headers = _headers_da_tabela(tabela)
        mapa = _mapear_colunas(headers)

        linhas = tabela.locator("tbody tr")
        try:
            if linhas.count() == 0:
                linhas = tabela.locator("tr")
            total_linhas = linhas.count()
        except Exception:
            total_linhas = 0

        registros: list[tuple[list[str], object]] = []
        for k in range(total_linhas):
            tr = linhas.nth(k)
            try:
                tds = tr.locator("td")
                n = tds.count()
            except Exception:
                continue
            if n < 2:
                continue
            valores = [limpar(tds.nth(j).inner_text()) for j in range(n)]
            if not any(valores):
                continue
            registros.append((valores, tr))

        if not registros:
            continue

        hjoin = normalizar(" ".join(headers))
        tem_link_desempenho = False
        for _, tr in registros:
            try:
                if tr.locator("a[href*='desempenho']").count():
                    tem_link_desempenho = True
                    break
            except Exception:
                continue

        # Ignora tabelas que claramente nao sao de cursos.
        if "curso" not in hjoin and "tipo" not in hjoin and not tem_link_desempenho:
            continue

        for valores, tr in registros:
            nome_curso = _pegar(valores, mapa, "curso") or valores[0]
            if not nome_curso:
                continue

            texto_status = _pegar(valores, mapa, "status") or " ".join(valores)

            desempenho_url = ""
            try:
                links = tr.locator("a[href*='desempenho']")
                if links.count():
                    href = links.first.get_attribute("href") or ""
                    if href:
                        desempenho_url = urljoin(f"{BASE_URL}/", href)
            except Exception:
                pass

            cursos.append(
                Curso(
                    aluno=aluno,
                    contrato_id=contrato_id,
                    curso=nome_curso,
                    tipo=_pegar(valores, mapa, "tipo"),
                    inicio=_pegar(valores, mapa, "inicio"),
                    expiracao=_pegar(valores, mapa, "expiracao"),
                    ultimo_acesso=_pegar(valores, mapa, "acesso"),
                    email=_pegar(valores, mapa, "email"),
                    status=_status_da_linha(texto_status, secao),
                    secao=secao,
                    desempenho_url=desempenho_url,
                )
            )

    # remove duplicados mantendo a ordem
    vistos = set()
    unicos: list[Curso] = []
    for curso in cursos:
        chave = (curso.contrato_id, curso.curso, curso.tipo, curso.secao)
        if chave in vistos:
            continue
        vistos.add(chave)
        unicos.append(curso)
    return unicos


# ---------------------------------------------------------------------------
# Pagina de desempenho
# ---------------------------------------------------------------------------

RE_SITUACAO_ROTULO = re.compile(
    r"(?:situa[çc][ãa]o|status)(?:\s+do\s+[a-zà-ú]+|\s+da\s+[a-zà-ú]+)?\s*[:\-]\s*([^\n\r]{3,40})",
    re.IGNORECASE,
)


def _rotulo_situacao(texto: str) -> str:
    """Converte frases/palavras conhecidas no rotulo canonico."""
    t = normalizar(texto)

    def tem(padrao: str, negar: str | None = None) -> bool:
        for achado in re.finditer(padrao, t):
            if negar and re.search(negar, achado.group(0)):
                continue
            return True
        return False

    if tem(r"certificad\w*.{0,40}(emitido|disponivel|liberado)",
           negar=r"\b(nao|sem|aguardando)\b"):
        return "Certificado emitido"
    if tem(r"\bconcluid[oa]s?\b", negar=r"\bnao\b"):
        return "Concluído"
    if tem(r"\bfinalizad[oa]s?\b", negar=r"\bnao\b"):
        return "Concluído"
    if tem(r"\baprovad[oa]s?\b", negar=r"\bnao\b"):
        return "Aprovado"
    if tem(r"\breprovad[oa]s?\b"):
        return "Reprovado"
    if tem(r"\bcursando\b"):
        return "Cursando"
    if tem(r"em andamento"):
        return "Em andamento"
    if tem(r"\bpendente\b"):
        return "Pendente"
    if tem(r"\bexpirad[oa]s?\b"):
        return "Expirado"
    if tem(r"\bbloquead[oa]s?\b"):
        return "Bloqueado"
    return ""


def _situacao_do_texto(texto: str, sem_tabelas: str = "") -> str:
    """Detecta a situacao. Prioriza rotulo explicito ("Situacao: X")."""
    rotulo = RE_SITUACAO_ROTULO.search(texto or "")
    if rotulo:
        situacao = _rotulo_situacao(rotulo.group(1))
        if situacao:
            return situacao
    # Sem rotulo: usa o texto fora das tabelas para nao pegar status de modulo.
    return _rotulo_situacao(sem_tabelas or texto or "")


def texto_principal(page) -> dict[str, str]:
    """Texto visivel da area de conteudo: completo e sem tabelas."""
    for seletor in ("main", "#content", ".content-wrapper", ".container-fluid", "body"):
        try:
            loc = page.locator(seletor).first
            if not (loc.count() and loc.is_visible()):
                continue
            dados = loc.evaluate(
                "el => {"
                "  const clone = el.cloneNode(true);"
                "  clone.querySelectorAll('table').forEach(t => t.remove());"
                "  return {completo: el.innerText || '', semTabelas: clone.textContent || ''};"
                "}"
            )
            completo = limpar((dados or {}).get("completo", ""))
            if completo:
                return {
                    "completo": completo,
                    "sem_tabelas": limpar((dados or {}).get("semTabelas", "")),
                }
        except Exception:
            continue
    return {"completo": "", "sem_tabelas": ""}


def extrair_tabelas(page, limite_linhas: int = 300) -> list[dict]:
    """Extrai todas as tabelas visiveis (cabecalho + linhas).

    Le tudo em uma unica chamada de JavaScript: e rapido e imune a linhas que
    somem do DOM enquanto o Playwright le celula por celula.
    """
    try:
        dados = page.evaluate(
            """(limite) => [...document.querySelectorAll('table')].map(t => {
                 const visivel = !!(t.getClientRects().length);
                 const celulas = (tr) => [...tr.children]
                   .filter(c => c.tagName === 'TD' || c.tagName === 'TH')
                   .map(c => (c.innerText || '').replace(/\\s+/g, ' ').trim());
                 let cabecalho = t.tHead && t.tHead.rows.length
                   ? celulas(t.tHead.rows[t.tHead.rows.length - 1]) : [];
                 const corpo = [...t.rows]
                   .filter(r => !(t.tHead && t.tHead.contains(r)))
                   .map(celulas);
                 if (!cabecalho.length && corpo.length && t.rows[0] &&
                     [...t.rows[0].children].some(c => c.tagName === 'TH')) {
                   cabecalho = corpo.shift();
                 }
                 return {visivel, cabecalho, linhas: corpo.slice(0, limite)};
               })""",
            limite_linhas,
        )
    except Exception:
        return []

    saida: list[dict] = []
    for i, tabela in enumerate(dados or []):
        if not tabela.get("visivel"):
            continue
        headers = [h for h in (tabela.get("cabecalho") or []) if h]
        linhas = [l for l in (tabela.get("linhas") or []) if any(l)]
        if not headers and linhas:
            headers = linhas.pop(0)
        if linhas:
            saida.append(
                {
                    "indice": i,
                    "headers": headers,
                    "linhas": linhas[:limite_linhas],
                }
            )
    return saida


def _indice_coluna(headers: list[str], *termos: str) -> int | None:
    for i, header in enumerate(headers):
        hn = normalizar(header)
        for termo in termos:
            if re.search(rf"\b{re.escape(termo)}", hn):
                return i
    return None


def _tabela_etapas(tabelas: list[dict]) -> dict | None:
    """Tabela 'Etapa | Situacao | Notas | ...' da pagina de desempenho."""
    for tabela in tabelas:
        headers = [normalizar(x) for x in tabela.get("headers", [])]
        if any("etapa" in h for h in headers) and any("situa" in h for h in headers):
            return tabela
    return None


def _contar_acessos(tabelas: list[dict]) -> int:
    """Linhas do 'Resumo da frequencia' (data + curso)."""
    for tabela in tabelas:
        headers = [normalizar(x) for x in tabela.get("headers", [])]
        if any(h == "data" for h in headers) and any("curso" in h for h in headers):
            return len(tabela.get("linhas", []))
    return 0


def analisar_desempenho_tabelas(tabelas: list[dict], texto: str = "", sem_tabelas: str = "") -> dict[str, str]:
    """Le progresso/itens/nota a partir da tabela de etapas do curso."""
    info = {"progresso": "", "itens": "", "nota": "", "situacao": "", "acessos": "", "tempo": ""}

    tabela = _tabela_etapas(tabelas)
    if tabela is None:
        # Pagina em outro formato: nao inventa numeros, so tenta a situacao.
        info["situacao"] = _situacao_do_texto(texto, sem_tabelas)
        return info

    headers = tabela.get("headers", [])
    i_situacao = _indice_coluna(headers, "situa")
    i_nota = _indice_coluna(headers, "nota")
    i_tempo = _indice_coluna(headers, "tempo")

    total = 0
    concluidos = 0
    em_andamento = 0
    linha_total: list[str] | None = None

    for valores in tabela.get("linhas", []):
        if not valores:
            continue
        if normalizar(valores[0]) == "total":
            linha_total = valores
            continue
        situacao = ""
        if i_situacao is not None and i_situacao < len(valores):
            situacao = limpar(valores[i_situacao])
        if not situacao:
            continue  # linha de modulo/cabecalho de etapa
        total += 1
        ns = normalizar(situacao)
        if "conclu" in ns:
            concluidos += 1
        elif "andamento" in ns:
            em_andamento += 1

    if not total:
        info["situacao"] = _situacao_do_texto(texto, sem_tabelas)
        return info

    info["itens"] = f"{concluidos}/{total}"
    info["progresso"] = f"{concluidos * 100 / total:.2f}".replace(".", ",") + "%"

    if linha_total:
        if i_nota is not None and i_nota < len(linha_total):
            info["nota"] = limpar(linha_total[i_nota])
        if i_tempo is not None and i_tempo < len(linha_total):
            info["tempo"] = limpar(linha_total[i_tempo])

    if concluidos == total:
        info["situacao"] = "Concluído"
    elif concluidos or em_andamento:
        info["situacao"] = "Em andamento"
    else:
        info["situacao"] = "Não iniciado"

    acessos = _contar_acessos(tabelas)
    if acessos:
        info["acessos"] = str(acessos)
    return info


def coletar_desempenho(context, curso: Curso, debug_dir: Path | None, erros: list[dict]) -> None:
    if not curso.desempenho_url:
        return

    page = context.new_page()
    try:
        page.goto(curso.desempenho_url, wait_until="domcontentloaded", timeout=TIMEOUT_NAV)
        esperar(page, 2.5)
        try:
            page.wait_for_load_state("networkidle", timeout=8000)
        except Exception:
            pass

        textos = texto_principal(page)
        tabelas = extrair_tabelas(page, limite_linhas=2000)
        info = analisar_desempenho_tabelas(tabelas, textos["completo"], textos["sem_tabelas"])
        curso.progresso = info["progresso"]
        curso.itens = info["itens"]
        curso.nota = info["nota"]
        curso.situacao = info["situacao"]
        curso.acessos = info["acessos"]
        curso.tempo = info["tempo"]

        curso.detalhes = [
            {
                "tabela": f"Tabela {t['indice'] + 1}",
                "headers": t["headers"],
                "linhas": t["linhas"],
            }
            for t in tabelas
        ]

        if not curso.itens:
            print("      ⚠️  Tabela de etapas não encontrada nesta página.")

        if debug_dir:
            debug_dir.mkdir(parents=True, exist_ok=True)
            arquivo = (
                debug_dir
                / f"{nome_arquivo_seguro(curso.aluno)}__"
                f"{nome_arquivo_seguro(curso.curso)}.html"
            )
            try:
                arquivo.write_text(page.content(), encoding="utf-8")
            except Exception:
                pass
    except Exception as exc:  # noqa: BLE001
        erros.append(
            {
                "aluno": curso.aluno,
                "curso": curso.curso,
                "etapa": "desempenho",
                "erro": str(exc),
            }
        )
        print(f"      ⚠️  Erro ao ler desempenho: {exc}")
    finally:
        try:
            page.close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Orquestracao
# ---------------------------------------------------------------------------

def nome_do_aluno(page) -> str:
    """Nome do aluno no cabecalho (ex.: 'Fulano 12 anos Contrato EAD')."""
    try:
        textos = page.locator("h3").all_inner_texts()
    except Exception:
        return ""
    for bruto in textos:
        texto = limpar(bruto)
        if not texto or "novidades" in texto.lower():
            continue
        achado = re.match(r"(.+?)\s+\d{1,2}\s+anos?\b", texto, re.IGNORECASE)
        if achado:
            return limpar(achado.group(1))
    return ""


def _salvar_debug(pagina, nome_arquivo: str) -> None:
    pasta = Path("debug_html")
    try:
        pasta.mkdir(parents=True, exist_ok=True)
        (pasta / nome_arquivo).write_text(pagina.content(), encoding="utf-8")
    except Exception:
        pass


def processar_contrato(context, page, nome: str, contrato_id: str, args, cursos: list[Curso], erros: list[dict]) -> None:
    """Le um contrato pelo id (pula a busca) e coleta cursos + desempenho."""
    url = f"{BASE_URL}/contratos/cursos-ead/{contrato_id}"
    if not ir_para(page, url):
        erros.append({"aluno": nome, "curso": "", "etapa": "contrato", "erro": f"falha ao abrir {url}"})
        return

    estado = _estado_pagina(page)
    if estado:
        print(f"   ❌ Página do contrato não carregou ({estado}).")
        erros.append({"aluno": nome, "curso": "", "etapa": "contrato", "erro": f"página não carregou ({estado})"})
        return

    revelar_abas(page)
    if args.debug_html:
        _salvar_debug(page, f"contrato_{contrato_id}.html")

    nome_aluno = nome_do_aluno(page) or nome
    if nome_aluno != nome:
        print(f"   👤 Aluno: {nome_aluno}")

    do_contrato = coletar_cursos(page, nome_aluno, contrato_id)
    if not do_contrato:
        print("   ⚠️  Nenhum curso lido nesta página de contrato.")
        erros.append({"aluno": nome_aluno, "curso": "", "etapa": "cursos", "erro": "nenhuma tabela de cursos reconhecida"})
        return

    com_desempenho = sum(1 for c in do_contrato if c.desempenho_url)
    print(
        f"   📚 {len(do_contrato)} curso(s) "
        f"({sum(1 for c in do_contrato if c.secao == 'Andamento')} em andamento, "
        f"{sum(1 for c in do_contrato if c.secao == 'Fazer')} a fazer, "
        f"{com_desempenho} com link de desempenho)"
    )

    for curso in do_contrato:
        if curso.desempenho_url:
            print(f"      ↳ lendo desempenho: {curso.curso[:60]}")
            coletar_desempenho(
                context, curso, Path("debug_html") if args.debug_html else None, erros
            )
        cursos.append(curso)


def processar_aluno(context, page, nome: str, args, cursos: list[Curso], erros: list[dict]) -> None:
    print(f"\n🔎 Buscando '{nome}'...")
    contratos = buscar_contratos(page, nome)
    if not contratos:
        print("   ❌ Nenhum contrato encontrado.")
        erros.append({"aluno": nome, "curso": "", "etapa": "busca", "erro": "contrato não encontrado"})
        return

    print(f"   Contrato(s): {', '.join(contratos)}")
    for contrato_id in contratos:
        processar_contrato(context, page, nome, contrato_id, args, cursos, erros)


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def _pct(valor: str) -> float | None:
    achado = re.search(r"(\d{1,3})(?:[.,](\d+))?", valor or "")
    if not achado:
        return None
    numero = float(f"{achado.group(1)}.{achado.group(2)}" if achado.group(2) else achado.group(1))
    return numero if 0 <= numero <= 100 else None


STATUS_CORES = {
    "Expirado": "FFC7CE",
    "Bloqueado": "FFC7CE",
    "Suspenso": "FFC7CE",
    "Reprovado": "FFC7CE",
    "Ativo": "C6EFCE",
    "Em andamento": "C6EFCE",
    "Concluído": "C6EFCE",
    "Aprovado": "C6EFCE",
    "Certificado emitido": "C6EFCE",
    "Pendente": "FFEB9C",
    "Cursando": "FFEB9C",
}


def _cor_progresso(valor: str) -> str | None:
    numero = _pct(valor)
    if numero is None:
        return None
    if numero >= 80:
        return "C6EFCE"
    if numero >= 40:
        return "FFEB9C"
    return "FFC7CE"


def _cor_carga(situacao: str) -> str | None:
    if situacao == "Cumprida":
        return "C6EFCE"
    if situacao in {"Quase", "Abaixo"}:
        return "FFEB9C" if situacao == "Quase" else "FFC7CE"
    return None


# ---------------------------------------------------------------------------
# Carga horaria ideal + observacoes do professor
# ---------------------------------------------------------------------------

CARGA_FILE = Path("carga_horaria.csv")
OBS_FILE = Path("observacoes.json")


def horas_do_tempo(texto: str) -> float | None:
    """'57:11:00' -> 57,18 horas. Aceita tambem '12:30'."""
    achado = re.match(r"^\s*(\d{1,3}):(\d{2})(?::(\d{2}))?\s*$", texto or "")
    if not achado:
        return None
    horas, minutos, segundos = (
        int(achado.group(1)),
        int(achado.group(2)),
        int(achado.group(3) or 0),
    )
    if minutos > 59 or segundos > 59:
        return None
    return horas + minutos / 60 + segundos / 3600


def _brl_horas(valor: float) -> str:
    return f"{valor:.1f}".replace(".", ",") + " h"


def carregar_carga_horaria(caminho: Path | str | None = None) -> dict[str, float]:
    """Le 'curso;horas' de um CSV. Linhas vazias ou com # sao ignoradas."""
    arquivo = Path(caminho) if caminho else CARGA_FILE
    if not arquivo.exists():
        return {}

    tabela: dict[str, float] = {}
    for linha in arquivo.read_text(encoding="utf-8-sig").splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#"):
            continue
        partes = [p.strip() for p in re.split(r"[;\t]", linha)]
        if len(partes) < 2:
            continue
        nome, bruto = partes[0], partes[1]
        bruto = re.sub(r"(?i)\s*h(oras?)?\s*$", "", bruto).replace(",", ".")
        try:
            horas = float(bruto)
        except ValueError:
            continue
        if not nome or normalizar(nome) == "curso" or horas <= 0:
            continue
        tabela[nome] = horas
    return tabela


def casar_carga(curso_nome: str, tabela: dict[str, float]) -> tuple[str, float] | None:
    """Acha a carga do curso aceitando nome parcial/erro de digitacao."""
    if not tabela or not curso_nome:
        return None

    alvo = normalizar(curso_nome)
    normalizada = {normalizar(nome): (nome, horas) for nome, horas in tabela.items()}

    if alvo in normalizada:
        return normalizada[alvo]

    # nome contido no outro ("Word 2019" <-> "Word 2019 - Avançado")
    contidos = [
        (nome, horas)
        for chave, (nome, horas) in normalizada.items()
        if chave and (chave in alvo or alvo in chave)
    ]
    if contidos:
        contidos.sort(key=lambda item: abs(len(normalizar(item[0])) - len(alvo)))
        return contidos[0]

    # parecido o bastante (erro de digitacao, plural, sigla)
    melhor = max(
        normalizada.items(),
        key=lambda item: difflib.SequenceMatcher(None, item[0], alvo).ratio(),
        default=None,
    )
    if melhor and difflib.SequenceMatcher(None, melhor[0], alvo).ratio() >= 0.82:
        return melhor[1]
    return None


def _rotulo_carga(percentual: float | None) -> str:
    if percentual is None:
        return ""
    if percentual >= 100:
        return "Cumprida"
    if percentual >= 75:
        return "Quase"
    if percentual > 0:
        return "Abaixo"
    return "Sem registro"


def aplicar_carga_horaria(cursos: list[Curso], tabela: dict[str, float]) -> int:
    """Preenche carga ideal, cumprimento, diferenca e situacao. Devolve quantos casaram."""
    casados = 0
    for curso in cursos:
        achado = casar_carga(curso.curso, tabela)
        if not achado:
            continue
        nome_tabela, ideal = achado
        curso.carga_horas = float(ideal)
        curso.carga_ideal = _brl_horas(float(ideal))

        horas = horas_do_tempo(curso.tempo)
        curso.tempo_horas = round(horas, 2) if horas is not None else 0.0
        if horas is None:
            curso.cumprimento = ""
            curso.situacao_carga = _rotulo_carga(None)
            continue

        percentual = horas * 100 / ideal
        curso.cumprimento = f"{percentual:.0f}%"
        curso.situacao_carga = _rotulo_carga(percentual)
        diferenca = ideal - horas
        curso.diferenca = (
            f"-{_brl_horas(abs(diferenca))}" if diferenca < 0 else _brl_horas(diferenca)
        )
        casados += 1
    return casados


def _chave_observacao(aluno: str, curso: str) -> str:
    return f"{normalizar(aluno)}|{normalizar(curso)}"


def chave_observacao(aluno: str, curso: str) -> str:
    """Chave publica usada pelo painel para guardar a observacao."""
    return _chave_observacao(aluno, curso)


def carregar_observacoes(caminho: Path | str | None = None) -> dict[str, str]:
    arquivo = Path(caminho) if caminho else OBS_FILE
    if not arquivo.exists():
        return {}
    try:
        dados = json.loads(arquivo.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if isinstance(dados, dict):
        return {str(k): str(v) for k, v in dados.items() if v}
    return {}


def salvar_observacoes(dados: dict[str, str], caminho: Path | str | None = None) -> None:
    arquivo = Path(caminho) if caminho else OBS_FILE
    limpo = {k: v for k, v in dados.items() if (v or "").strip()}
    arquivo.write_text(
        json.dumps(limpo, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def aplicar_observacoes(cursos: list[Curso], dados: dict[str, str]) -> None:
    if not dados:
        return
    for curso in cursos:
        curso.observacao = dados.get(_chave_observacao(curso.aluno, curso.curso), "")


def aplicar_extras(
    cursos: list[Curso],
    carga: dict[str, float] | None = None,
    observacoes: dict[str, str] | None = None,
) -> None:
    """Aplica carga horaria + observacoes (usado na extracao e ao regerar relatorios)."""
    aplicar_carga_horaria(cursos, carga if carga is not None else carregar_carga_horaria())
    aplicar_observacoes(
        cursos, observacoes if observacoes is not None else carregar_observacoes()
    )


def _ajustar_larguras(ws, maximo: int = 45) -> None:
    for coluna in ws.columns:
        letra = get_column_letter(coluna[0].column)
        maior = 0
        for celula in coluna:
            texto = "" if celula.value is None else str(celula.value)
            maior = max(maior, max((len(linha) for linha in texto.split("\n")), default=0))
        ws.column_dimensions[letra].width = min(maior + 4, maximo)


def _cabecalho(ws, headers: list[str]) -> None:
    ws.append(headers)
    preenchimento = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    fonte = Font(bold=True, color="FFFFFF")
    for coluna in range(1, len(headers) + 1):
        celula = ws.cell(row=1, column=coluna)
        celula.fill = preenchimento
        celula.font = fonte
        celula.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"


def salvar_excel(cursos: list[Curso], erros: list[dict], arquivo: str) -> None:
    wb = Workbook()
    # --- Aba principal -----------------------------------------------------
    ws = wb.active
    ws.title = "Desempenho"
    headers = [
        "Aluno", "Contrato", "Curso", "Tipo", "Início", "Expiração",
        "Último Acesso", "E-mail", "Status", "Seção",
        "Progresso", "Aulas", "Nota", "Situação", "Acessos", "Tempo",
        "Carga Ideal", "Cumprimento", "Diferença", "Situação da Carga",
        "Observações", "Desempenho (URL)",
    ]
    _cabecalho(ws, headers)

    col_url = headers.index("Desempenho (URL)") + 1
    col_status = headers.index("Status") + 1
    col_progresso = headers.index("Progresso") + 1
    col_cumprimento = headers.index("Cumprimento") + 1
    col_situacao_carga = headers.index("Situação da Carga") + 1
    col_obs = headers.index("Observações") + 1

    for curso in cursos:
        ws.append([
            curso.aluno, curso.contrato_id, curso.curso, curso.tipo,
            curso.inicio, curso.expiracao, curso.ultimo_acesso, curso.email,
            curso.status, curso.secao, curso.progresso, curso.itens,
            curso.nota, curso.situacao, curso.acessos, curso.tempo,
            curso.carga_ideal, curso.cumprimento, curso.diferenca,
            curso.situacao_carga, curso.observacao, curso.desempenho_url,
        ])
        linha = ws.max_row

        cor_status = STATUS_CORES.get(curso.status)
        if cor_status:
            ws.cell(row=linha, column=col_status).fill = PatternFill(
                start_color=cor_status, end_color=cor_status, fill_type="solid"
            )

        cor_prog = _cor_progresso(curso.progresso)
        if cor_prog:
            ws.cell(row=linha, column=col_progresso).fill = PatternFill(
                start_color=cor_prog, end_color=cor_prog, fill_type="solid"
            )

        cor_carga = _cor_carga(curso.situacao_carga)
        if cor_carga:
            for coluna in (col_cumprimento, col_situacao_carga):
                ws.cell(row=linha, column=coluna).fill = PatternFill(
                    start_color=cor_carga, end_color=cor_carga, fill_type="solid"
                )

        celula_obs = ws.cell(row=linha, column=col_obs)
        celula_obs.alignment = Alignment(vertical="top", wrap_text=True)

        if curso.desempenho_url:
            celula = ws.cell(row=linha, column=col_url)
            celula.hyperlink = curso.desempenho_url
            celula.font = Font(color="0563C1", underline="single")

    ws.auto_filter.ref = ws.dimensions
    _ajustar_larguras(ws)
    ws.column_dimensions[get_column_letter(col_obs)].width = 46

    # --- Resumo por aluno --------------------------------------------------
    ws_resumo = wb.create_sheet("Resumo")
    _cabecalho(
        ws_resumo,
        [
            "Aluno", "Cursos", "Em andamento", "A fazer", "Com desempenho",
            "Progresso médio", "Carga ideal (h)", "Tempo total (h)", "Cumprimento médio",
        ],
    )

    ordem: list[str] = []
    for curso in cursos:
        if curso.aluno not in ordem:
            ordem.append(curso.aluno)

    for aluno in ordem:
        do_aluno = [c for c in cursos if c.aluno == aluno]
        percentuais = [p for p in (_pct(c.progresso) for c in do_aluno) if p is not None]
        media = f"{sum(percentuais) / len(percentuais):.0f}%" if percentuais else ""
        carga_total = sum(c.carga_horas for c in do_aluno)
        tempo_total = sum(c.tempo_horas for c in do_aluno)
        cumprimentos = [
            p for p in (_pct(c.cumprimento) for c in do_aluno) if p is not None
        ]
        cumprimento = (
            f"{sum(cumprimentos) / len(cumprimentos):.0f}%" if cumprimentos else ""
        )
        ws_resumo.append([
            aluno,
            len(do_aluno),
            sum(1 for c in do_aluno if c.secao == "Andamento"),
            sum(1 for c in do_aluno if c.secao == "Fazer"),
            sum(1 for c in do_aluno if c.desempenho_url),
            media,
            round(carga_total, 1) if carga_total else "",
            round(tempo_total, 1) if tempo_total else "",
            cumprimento,
        ])

    ws_resumo.auto_filter.ref = ws_resumo.dimensions
    _ajustar_larguras(ws_resumo)

    # --- Detalhes ----------------------------------------------------------
    linhas_detalhe = []
    for curso in cursos:
        for tabela in curso.detalhes:
            cabecalho = " | ".join(tabela["headers"])
            limite = 400
            for linha in tabela["linhas"][:limite]:
                linhas_detalhe.append([
                    curso.aluno,
                    curso.curso,
                    tabela["tabela"],
                    cabecalho,
                    " | ".join(linha),
                ])
            if len(tabela["linhas"]) > limite:
                linhas_detalhe.append([
                    curso.aluno, curso.curso, tabela["tabela"], cabecalho,
                    f"... {len(tabela['linhas']) - limite} linha(s) omitida(s) (veja o HTML/JSON)",
                ])

    if linhas_detalhe:
        ws_det = wb.create_sheet("Detalhes")
        _cabecalho(ws_det, ["Aluno", "Curso", "Tabela", "Cabeçalho", "Linha"])
        for linha in linhas_detalhe:
            ws_det.append(linha)
        ws_det.auto_filter.ref = ws_det.dimensions
        _ajustar_larguras(ws_det, maximo=60)

    # --- Erros -------------------------------------------------------------
    if erros:
        ws_err = wb.create_sheet("Erros")
        _cabecalho(ws_err, ["Aluno", "Curso", "Etapa", "Erro"])
        for erro in erros:
            ws_err.append([
                erro.get("aluno", ""), erro.get("curso", ""),
                erro.get("etapa", ""), erro.get("erro", ""),
            ])
        _ajustar_larguras(ws_err, maximo=80)

    Path(arquivo).parent.mkdir(parents=True, exist_ok=True)
    wb.save(arquivo)
    print(f"\n✅ Planilha salva: {arquivo}")


def _html_escape(texto: object) -> str:
    return html.escape("" if texto is None else str(texto), quote=True)


def _cartao_aluno(aluno: str, do_aluno: list[Curso]) -> str:
    percentuais = [p for p in (_pct(c.progresso) for c in do_aluno) if p is not None]
    media = f"{sum(percentuais) / len(percentuais):.0f}%" if percentuais else "—"
    return (
        '<div class="card">'
        f'<h2>{_html_escape(aluno)}</h2>'
        f'<p class="meta">{len(do_aluno)} curso(s) · média {media}</p>'
        '</div>'
    )


def _linha_curso_html(curso: Curso, mostrar_aluno: bool = True) -> str:
    numero = _pct(curso.progresso)
    largura = 0 if numero is None else max(0, min(100, numero))
    classe_barra = "ok" if largura >= 80 else ("mid" if largura >= 40 else "low")
    barra = (
        '<div class="bar"><span class="' + classe_barra + '" '
        f'style="width:{largura:.0f}%"></span></div>'
        f'<small>{_html_escape(curso.progresso) or "—"}</small>'
    )
    link = (
        f'<a href="{_html_escape(curso.desempenho_url)}" target="_blank" rel="noopener">abrir</a>'
        if curso.desempenho_url else ""
    )
    celula_aluno = f"<td>{_html_escape(curso.aluno)}</td>" if mostrar_aluno else ""
    if curso.carga_ideal:
        percentual = _pct(curso.cumprimento)
        classe = "ok" if (percentual or 0) >= 100 else ("mid" if (percentual or 0) >= 75 else "bad")
        celula_carga = (
            '<td class="carga">'
            f'<strong>{_html_escape(curso.tempo) or "—"}</strong>'
            f'<span> de {_html_escape(curso.carga_ideal)}</span>'
            f'<br><span class="badge {classe}">{_html_escape(curso.situacao_carga)}</span>'
            f' <small>{_html_escape(curso.cumprimento)}</small>'
            '</td>'
        )
    else:
        celula_carga = (
            f'<td class="carga" title="Sem carga horária cadastrada em carga_horaria.csv">'
            f'{_html_escape(curso.tempo) or "—"}</td>'
        )
    celula_obs = (
        f'<td class="obs">{_html_escape(curso.observacao)}</td>'
        if curso.observacao
        else '<td class="obs vazio">—</td>'
    )
    return (
        "<tr>"
        f"{celula_aluno}"
        f"<td>{_html_escape(curso.curso)}</td>"
        f'<td><span class="badge {_classe_status(curso.status)}">{_html_escape(curso.status)}</span></td>'
        f"<td>{barra}</td>"
        f"<td>{_html_escape(curso.itens) or '—'}</td>"
        f"<td>{_html_escape(curso.nota) or '—'}</td>"
        f"<td>{_html_escape(curso.situacao) or '—'}</td>"
        f"<td>{_html_escape(curso.acessos) or '—'}</td>"
        f"{celula_carga}"
        f"{celula_obs}"
        f"<td>{link}</td>"
        "</tr>"
    )


def _classe_status(status: str) -> str:
    if status in {"Ativo", "Em andamento", "Concluído", "Aprovado", "Certificado emitido"}:
        return "ok"
    if status in {"Expirado", "Bloqueado", "Suspenso", "Reprovado"}:
        return "bad"
    return "mid"


def salvar_html(cursos: list[Curso], erros: list[dict], arquivo: str) -> None:
    """Relatorio autonomo (HTML) para abrir no navegador."""
    alunos: list[str] = []
    for curso in cursos:
        if curso.aluno not in alunos:
            alunos.append(curso.aluno)

    cards = "".join(
        _cartao_aluno(aluno, [c for c in cursos if c.aluno == aluno]) for aluno in alunos
    )
    mostrar_aluno = len(alunos) > 1
    linhas = "\n".join(_linha_curso_html(c, mostrar_aluno) for c in cursos)
    th_aluno = "<th>Aluno</th>" if mostrar_aluno else ""

    bloco_erros = ""
    if erros:
        itens = "".join(
            f"<li><strong>{_html_escape(e.get('aluno'))}</strong> — "
            f"{_html_escape(e.get('etapa'))}: {_html_escape(e.get('erro'))}</li>"
            for e in erros
        )
        bloco_erros = f'<section class="erros"><h2>Avisos</h2><ul>{itens}</ul></section>'

    gerado = datetime.now().strftime("%d/%m/%Y %H:%M")
    documento = f"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Desempenho de alunos — CGD</title>
<style>
  :root {{ color-scheme: light; }}
  * {{ box-sizing: border-box; }}
  body {{ margin: 0; font: 15px/1.5 -apple-system, "Segoe UI", Roboto, Arial, sans-serif;
         color: #1f2430; background: #f5f6f8; }}
  header {{ background: #2F5496; color: #fff; padding: 28px 32px; }}
  header h1 {{ margin: 0 0 4px; font-size: 24px; }}
  header p {{ margin: 0; opacity: .85; font-size: 14px; }}
  main {{ padding: 24px 32px 48px; }}
  .cards {{ display: flex; gap: 14px; flex-wrap: wrap; margin-bottom: 24px; }}
  .card {{ background: #fff; border: 1px solid #e2e5ea; border-radius: 10px;
           padding: 14px 18px; min-width: 200px; }}
  .card h2 {{ margin: 0; font-size: 16px; }}
  .card .meta {{ margin: 4px 0 0; color: #5b6472; font-size: 13px; }}
  .tabela {{ background: #fff; border: 1px solid #e2e5ea; border-radius: 10px; overflow: auto; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
  th {{ background: #eef1f7; text-align: left; padding: 10px 8px; font-size: 12px;
        text-transform: uppercase; letter-spacing: .04em; color: #3c4658; white-space: nowrap; }}
  td {{ padding: 10px 8px; border-top: 1px solid #eef0f4; vertical-align: middle; white-space: nowrap; }}
  td.carga {{ white-space: nowrap; font-size: 13px; }}
  td.carga strong {{ font-weight: 600; }}
  td.carga span {{ color: #5b6472; }}
  td.obs {{ white-space: normal; min-width: 200px; max-width: 340px; font-size: 13px; }}
  td.vazio {{ color: #b0b7c3; }}
  tr:hover td {{ background: #fafbfd; }}
  .badge {{ display: inline-block; padding: 2px 9px; border-radius: 999px;
            font-size: 12px; font-weight: 600; }}
  .badge.ok {{ background: #d8f3e0; color: #14663a; }}
  .badge.mid {{ background: #fdf1c7; color: #755b00; }}
  .badge.bad {{ background: #fbdada; color: #8a1f1f; }}
  .bar {{ width: 90px; height: 7px; background: #eceef2; border-radius: 999px; overflow: hidden; }}
  .bar span {{ display: block; height: 100%; }}
  .bar span.ok {{ background: #2f9e5f; }}
  .bar span.mid {{ background: #d6a518; }}
  .bar span.low {{ background: #d9534f; }}
  small {{ color: #5b6472; }}
  a {{ color: #2F5496; }}
  .erros {{ margin-top: 22px; background: #fff7f7; border: 1px solid #f3caca;
            border-radius: 10px; padding: 14px 18px; }}
  .erros h2 {{ font-size: 15px; margin: 0 0 8px; }}
  .erros ul {{ margin: 0; padding-left: 20px; }}
  .acoes-relatorio {{ display: flex; align-items: center; gap: 8px;
                      margin: 14px 32px 0; }}
  .acoes-relatorio a, .acoes-relatorio button {{ background: #fff; color: #2F5496;
                      border: 1px solid #cfd5df; border-radius: 8px; padding: 8px 12px;
                      font: 600 13px/1.2 -apple-system, "Segoe UI", sans-serif;
                      text-decoration: none; cursor: pointer; }}
  @media print {{
    .acoes-relatorio {{ display: none; }}
  }}
</style>
</head>
<body>
<header>
  <h1>Desempenho de alunos</h1>
  <p>Gerado em {gerado} · {len(alunos)} aluno(s) · {len(cursos)} curso(s)</p>
</header>
<nav class="acoes-relatorio" id="acoes-relatorio" aria-label="Acoes do relatorio">
  <a id="acao-observacoes" href="#">Observações</a>
  <a id="acao-excel" href="#" download>Baixar Excel</a>
  <a id="acao-pdf" href="#" download>Baixar PDF</a>
</nav>
<main>
  <section class="cards">{cards}</section>
  <div class="tabela">
    <table>
      <thead><tr>
        {th_aluno}<th>Curso</th><th>Status</th>
        <th>Progresso</th><th>Aulas</th><th>Nota</th><th>Situação</th><th>Acessos</th>
        <th>Carga horária</th><th>Observações</th><th></th>
      </tr></thead>
      <tbody>{linhas}</tbody>
    </table>
  </div>
  {bloco_erros}
</main>
<script>
(() => {{
  const caminho = window.location.pathname;
  if (!window.location.protocol.startsWith("http")) {{
    document.getElementById("acoes-relatorio").style.display = "none";
    return;
  }}
  const nomeHtml = decodeURIComponent(caminho.split("/").pop());
  const nomeBase = nomeHtml.replace(/\.html$/i, "");
  document.getElementById("acao-observacoes").href = "/observacoes/" + encodeURIComponent(nomeHtml);
  document.getElementById("acao-excel").href = "/baixar/" + encodeURIComponent(nomeBase + ".xlsx");
  document.getElementById("acao-pdf").href = "/pdf/" + encodeURIComponent(nomeHtml);
}})();
</script>
</body>
</html>
"""
    Path(arquivo).parent.mkdir(parents=True, exist_ok=True)
    Path(arquivo).write_text(documento, encoding="utf-8")
    print(f"✅ Relatório HTML salvo: {arquivo}")


def salvar_json(cursos: list[Curso], erros: list[dict], arquivo: str) -> None:
    payload = {
        "gerado_em": datetime.now().isoformat(timespec="seconds"),
        "cursos": [asdict(c) for c in cursos],
        "erros": erros,
    }
    Path(arquivo).parent.mkdir(parents=True, exist_ok=True)
    Path(arquivo).write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"✅ JSON salvo: {arquivo}")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extrai o desempenho de alunos do portal CGD."
    )
    parser.add_argument("--alunos", nargs="+", default=[], help="nomes dos alunos")
    parser.add_argument("--contratos", nargs="+", help="ids de contrato (pula a busca por nome)")
    parser.add_argument("--lista", help="arquivo .txt/.csv com um nome por linha")
    parser.add_argument("--saida", help="arquivo Excel de saída")
    parser.add_argument("--headless", action="store_true", help="rodar sem janela (não recomendado no 1º login)")
    parser.add_argument("--lento", type=int, default=0, help="slow_mo do Playwright em ms")
    parser.add_argument("--nova-sessao", action="store_true", help="ignorar cgd_session.json e logar de novo")
    parser.add_argument("--debug-html", action="store_true", help="salvar o HTML das páginas de desempenho")
    parser.add_argument("--json", action="store_true", help="gerar também um .json com tudo")
    parser.add_argument("--html", action="store_true", help="gerar também um relatório .html")
    parser.add_argument(
        "--carga-horaria",
        help="CSV 'curso;horas' com a carga ideal (padrão: carga_horaria.csv)",
    )
    parser.add_argument(
        "--observacoes",
        help="JSON com as observações do professor (padrão: observacoes.json)",
    )
    parser.add_argument("--pausar", action="store_true", help="esperar Enter antes de fechar o navegador")
    parser.add_argument(
        "--sem-login-manual",
        action="store_true",
        help="não esperar login manual: falha na hora se a sessão expirou (usado pelo painel)",
    )
    parser.add_argument("--session-file", help="arquivo da sessão Playwright desta conta")
    return parser.parse_args()


def coletar_nomes(args) -> list[str]:
    nomes = [limpar(n) for n in args.alunos if limpar(n)]

    if args.lista:
        caminho = Path(args.lista)
        if not caminho.exists():
            raise SystemExit(f"❌ Arquivo não encontrado: {caminho}")
        for linha in caminho.read_text(encoding="utf-8-sig").splitlines():
            linha = limpar(linha)
            if linha and not linha.startswith("#"):
                nomes.append(linha)

    if not nomes:
        print("=" * 52)
        print("  CGD - Extrator de Desempenho de Alunos")
        print("=" * 52)
        entrada = input("\n👤 Nome(s) do(s) aluno(s) (separe por ';'): ")
        nomes = [limpar(n) for n in entrada.split(";") if limpar(n)]

    # remove duplicados preservando a ordem
    unicos: list[str] = []
    for nome in nomes:
        if nome not in unicos:
            unicos.append(nome)
    return unicos


def arquivo_padrao(primeiro: str) -> str:
    data = datetime.now().strftime("%Y-%m-%d_%H%M")
    return str(Path("relatorios") / f"desempenho_{nome_arquivo_seguro(primeiro)}_{data}.xlsx")


def main() -> None:
    args = parse_args()
    session_file = Path(args.session_file) if args.session_file else SESSION_FILE
    email = os.environ.get("CGD_EMAIL", "").strip()
    senha = os.environ.get("CGD_SENHA", "")
    contratos_diretos = [c.strip() for c in (args.contratos or []) if c.strip()]

    if contratos_diretos:
        nomes = [f"Contrato {cid}" for cid in contratos_diretos]
    else:
        nomes = coletar_nomes(args)
        if not nomes:
            raise SystemExit("❌ Informe pelo menos um aluno (ou use --contratos).")

    arquivo = args.saida or arquivo_padrao(nomes[0])

    carga_tabela = carregar_carga_horaria(args.carga_horaria)
    observacoes = carregar_observacoes(args.observacoes)
    if carga_tabela:
        print(f"⏱️  {len(carga_tabela)} carga(s) horária(s) carregada(s) de "
              f"{args.carga_horaria or CARGA_FILE}.")

    resultados: list[Curso] = []
    erros: list[dict] = []

    with sync_playwright() as p:
        if args.headless:
            print("⚠️  Modo headless é bloqueado pelo Cloudflare do CGD; use sem --headless.")
        browser, context, page = iniciar_navegador(
            p, args.headless, args.lento, args.nova_sessao, session_file
        )
        try:
            permitir_manual = not args.sem_login_manual
            garantir_login(
                context,
                page,
                permitir_login_manual=permitir_manual,
                email=email,
                senha=senha,
                session_file=session_file,
            )
            if contratos_diretos:
                for cid in contratos_diretos:
                    try:
                        processar_contrato(
                            context, page, f"Contrato {cid}", cid, args, resultados, erros
                        )
                    except KeyboardInterrupt:
                        print("\n⏹️  Interrompido pelo usuário.")
                        break
                    except Exception as exc:  # noqa: BLE001
                        print(f"   ❌ Erro inesperado: {exc}")
                        erros.append({"aluno": f"Contrato {cid}", "curso": "", "etapa": "geral", "erro": str(exc)})
            for nome in ([] if contratos_diretos else nomes):
                try:
                    processar_aluno(context, page, nome, args, resultados, erros)
                except KeyboardInterrupt:
                    print("\n⏹️  Interrompido pelo usuário.")
                    break
                except Exception as exc:  # noqa: BLE001
                    print(f"   ❌ Erro inesperado: {exc}")
                    erros.append({"aluno": nome, "curso": "", "etapa": "geral", "erro": str(exc)})
        finally:
            if args.pausar:
                try:
                    input("\nPressione Enter para fechar...")
                except Exception:
                    pass
            try:
                browser.close()
            except Exception:
                pass

    if not resultados and not erros:
        print("\n⚠️  Nada foi extraído.")
        return

    aplicar_extras(resultados, carga_tabela, observacoes)
    if not observacoes and args.observacoes:
        print(f"⚠️  Nenhuma observação encontrada em {args.observacoes}.")

    salvar_excel(resultados, erros, arquivo)
    if args.json:
        salvar_json(resultados, erros, str(Path(arquivo).with_suffix(".json")))
    if args.html:
        salvar_html(resultados, erros, str(Path(arquivo).with_suffix(".html")))

    andamento = sum(1 for c in resultados if c.secao == "Andamento")
    a_fazer = sum(1 for c in resultados if c.secao == "Fazer")
    com_perf = sum(1 for c in resultados if c.progresso or c.itens or c.nota)
    com_carga = sum(1 for c in resultados if c.carga_ideal)
    print(f"\n📋 Resumo: {len(resultados)} curso(s) | {andamento} em andamento | "
          f"{a_fazer} a fazer | {com_perf} com dados de desempenho")
    if com_carga:
        print(f"⏱️  Carga horária comparada em {com_carga} curso(s).")
    elif carga_tabela:
        print("⚠️  Carga horária cadastrada, mas nenhum curso casou pelo nome.")
    if erros:
        print(f"⚠️  {len(erros)} erro(s) — veja a aba 'Erros'.")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as exc:  # noqa: BLE001
        print(f"\n❌ Erro: {exc}")
        sys.exit(1)
