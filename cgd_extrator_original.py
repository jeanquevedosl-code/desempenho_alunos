"""
CGD - Extrator de desempenho de alunos (Playwright)
Uso: python cgd_extrator.py
Requisitos: pip install playwright openpyxl
            playwright install chromium
"""

from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import getpass
import re

BASE_URL = "https://app.cgd.com.br"


def login(page, email, senha):
    """Abre o navegador e espera o usuário logar manualmente (Cloudflare)."""
    page.goto(f"{BASE_URL}/login", wait_until="domcontentloaded")

    print("\n⚠️  O navegador foi aberto.")
    print("   Faça login MANUALMENTE no navegador (digite e-mail/senha).")
    print("   O script continuará automaticamente após o login.\n")

    # Aguarda o usuário logar (até 5 minutos)
    page.wait_for_url("**/contratos**", timeout=300000)
    print("✅ Login detectado! Continuando...")    

def buscar_contrato_por_nome(page, nome_aluno):
    page.goto(f"{BASE_URL}/contratos", timeout=15000, wait_until="domcontentloaded")
    page.wait_for_timeout(2000)  # espera 2s pra página renderizar

    search_input = page.locator("input[placeholder*='usca'], input[name*='busca'], input[name*='search'], input[name*='nome']").first
    if search_input.is_visible():
        search_input.fill(nome_aluno)
        search_input.press("Enter")
        page.wait_for_timeout(2000)

    links = page.locator("a[href*='/contratos/cursos-ead/']")
    if links.count() == 0:
        return None

    href = links.first.get_attribute("href")
    match = re.search(r"/contratos/cursos-ead/(\d+)", href)
    return match.group(1) if match else None   


def parse_andamento(page):
    """Extrai dados da tabela 'Andamento'."""
    cursos = []
    # Tabela com colunas: Curso, Tipo, Início, Expiração, Acesso, E-mail, Ações
    rows = page.locator("table.table-condensed table-striped table-bordered tbody tr")
    count = rows.count()

    for i in range(count):
        row = rows.nth(i)
        tds = row.locator("td")
        if tds.count() >= 6:
            curso = tds.nth(0).inner_text().strip()
            tipo = tds.nth(1).inner_text().strip()
            inicio = tds.nth(2).inner_text().strip()
            expiracao = tds.nth(3).inner_text().strip()
            acesso = tds.nth(4).inner_text().strip()
            email = tds.nth(5).inner_text().strip()

            # Status
            acoes = tds.nth(6)
            texto_acoes = acoes.inner_text()
            if "Expirado" in texto_acoes:
                status = "Expirado"
            elif "Bloquear" in texto_acoes:
                status = "Bloqueado"
            elif "Desbloquear" in texto_acoes:
                status = "Ativo"
            else:
                status = "Em andamento"

            # Link de desempenho
            link_perf = acoes.locator("a[href*='/desempenho/']").first
            perf_url = ""
            if link_perf.is_visible():
                perf_url = link_perf.get_attribute("href") or ""

            cursos.append({
                "curso": curso,
                "tipo": tipo,
                "inicio": inicio,
                "expiracao": expiracao,
                "ultimo_acesso": acesso,
                "email": email,
                "status": status,
                "desempenho_url": perf_url,
                "secao": "Andamento",
            })
    return cursos


def parse_fazer(page):
    """Extrai dados da tabela 'Fazer'."""
    cursos = []
    table = page.locator("#table_cursos_fazer")
    if not table.is_visible():
        return cursos

    rows = table.locator("tbody tr")
    count = rows.count()

    for i in range(count):
        row = rows.nth(i)
        tds = row.locator("td")
        if tds.count() >= 2:
            curso = tds.nth(0).inner_text().strip()
            tipo = tds.nth(1).inner_text().strip()
            cursos.append({
                "curso": curso,
                "tipo": tipo,
                "inicio": "",
                "expiracao": "",
                "ultimo_acesso": "",
                "email": "",
                "status": "Pendente",
                "desempenho_url": "",
                "secao": "Fazer",
            })
    return cursos


def salvar_excel(aluno_nome, cursos, arquivo=None):
    """Salva os dados em arquivo Excel formatado."""
    if not arquivo:
        arquivo = f"desempenho_{aluno_nome.replace(' ', '_').lower()}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Desempenho"

    headers = ["Curso", "Tipo", "Início", "Expiração", "Último Acesso", "E-mail", "Status", "Seção"]
    ws.append(headers)

    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    for col in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    status_colors = {
        "Expirado": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
        "Bloqueado": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
        "Ativo": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        "Em andamento": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        "Pendente": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
    }

    for c in cursos:
        ws.append([c["curso"], c["tipo"], c["inicio"], c["expiracao"],
                   c["ultimo_acesso"], c["email"], c["status"], c["secao"]])
        fill = status_colors.get(c["status"])
        if fill:
            ws.cell(row=ws.max_row, column=7).fill = fill

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 45)

    wb.save(arquivo)
    print(f"\n✅ Planilha salva: {arquivo}")


def main():
    print("=" * 50)
    print("  CGD - Extrator de Desempenho (Playwright)")
    print("=" * 50)

    email = input("\n📧 E-mail: ")
    senha = getpass.getpass("🔑 Senha: ")
    nome_aluno = input("\n👤 Nome do aluno: ").strip()

    if not nome_aluno:
        print("❌ Informe o nome do aluno.")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(
        headless=False,
        args=["--disable-blink-features=AutomationControlled"]
   )
        context = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36"
        )
        page = context.new_page()   

        # 1. Login
        print("\n⏳ Fazendo login...")
        login(page, email, senha)

        # 2. Buscar contrato
        print(f"⏳ Buscando contrato de '{nome_aluno}'...")
        contrato_id = buscar_contrato_por_nome(page, nome_aluno)

        if not contrato_id:
            print("❌ Nenhum contrato encontrado.")
            print("💡 Dica: abra o navegador manualmente e me diga qual URL você usa pra achar um aluno.")
            browser.close()
            return

        print(f"   Contrato: {contrato_id}")

        # 3. Abrir página do contrato
        page.goto(f"{BASE_URL}/contratos/cursos-ead/{contrato_id}", wait_until="domcontentloaded")
        page.wait_for_timeout(2000)   

        # 4. Parsear tabelas
        andamento = parse_andamento(page)
        fazer = parse_fazer(page)
        todos = andamento + fazer

        if not todos:
            print("⚠️ Nenhum curso encontrado.")
            browser.close()
            return

        # 5. Salvar Excel
        salvar_excel(nome_aluno, todos)
        print(f"\n📋 Resumo: {len(andamento)} em andamento, {len(fazer)} pendentes")

        # Mantém navegador aberto pra inspeção (opcional)
        input("\nPressione Enter para fechar...")
        browser.close()


if __name__ == "__main__":
    main()   