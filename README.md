# CGD — Extrator de desempenho de alunos

Automação em Python (Playwright) que entra no portal **app.cgd.com.br**, localiza o
contrato do aluno, lê os cursos das abas **Andamento** e **Fazer**, abre o
**Desempenho** de cada curso e gera uma planilha Excel consolidada.

## Instalação

```bash
pip install playwright openpyxl
playwright install chromium
```

## Uso

```bash
# 1) um aluno (modo interativo pede o nome)
python cgd_extrator.py

# 2) vários alunos de uma vez
python cgd_extrator.py --alunos "Maria Silva" "João Souza"

# 3) lista em arquivo (um nome por linha, aceita .txt ou .csv)
python cgd_extrator.py --lista alunos.txt --saida relatorio.xlsx

# 4) guardar o HTML das páginas de desempenho (para ajustar o parser)
python cgd_extrator.py --alunos "Maria Silva" --debug-html

# 5) gerar também o JSON com tudo (útil para outros relatórios)
python cgd_extrator.py --alunos "Maria Silva" --json

# 6) gerar relatório HTML (abre no navegador, bom para apresentar)
python cgd_extrator.py --alunos "Maria Silva" --html

# 7) por id de contrato (sem buscar por nome)
python cgd_extrator.py --contratos 638795 --html
```

Opções: `--headless`, `--lento 200` (slow_mo em ms), `--nova-sessao`,
`--pausar` (espera Enter antes de fechar o navegador), `--sem-login-manual`
(falha na hora se a sessão expirou — usado pelo painel).

Por padrão os arquivos vão para a pasta `relatorios/`.

### Fluxo pelo chat

Você pode simplesmente mandar os nomes na conversa: o extrator roda com
`--headless --json --html` e o relatório é aberto aqui no painel, com a planilha
`.xlsx` em `relatorios/` para baixar.

### Login

O site usa **Cloudflare**, então o navegador abre em janela visível e o login é
manual (e-mail, senha e o desafio do Cloudflare). O script detecta quando você
entrou, salva a sessão em `cgd_session.json` e segue sozinho.

Depois disso, as próximas execuções reaproveitam a sessão e rodam sem nenhuma
interação — até ela expirar (aí a janela de login aparece de novo). Para forçar
um novo login use `--nova-sessao`. Não coloque `cgd_session.json` em repositório
ou rede compartilhada: ele contém os cookies da sua sessão.

## Painel na web (para outra pessoa usar)

O robô continua rodando **neste PC** (é ele que tem a sessão do CGD e o Chrome
que passa pelo Cloudflare). O painel é só a porta de entrada.

```bash
# só neste computador
python painel.py            # http://127.0.0.1:8000

# publicado na internet com link HTTPS (duplo clique no .bat também funciona)
python painel_online.py     # imprime o link e abre no navegador
.\INICIAR_PAINEL_ONLINE.bat # mesma coisa, sem precisar do terminal
```

- **Acesso:** o painel principal é aberto sem senha.
- **Login CGD:** cada pessoa informa seu e-mail e senha do CGD no formulário.
  As credenciais são usadas em memória para o job e não são gravadas.
- **Link público:** `painel_online.py` cria um túnel de saída do Cloudflare
  (`tools/cloudflared.exe`), sem abrir porta no roteador. O endereço é do tipo
  `https://algo-algo.trycloudflare.com`, muda a cada execução e só vale
  enquanto a janela do script estiver aberta (fica anotado em `painel_url.txt`).
- **Deixar no ar ao ligar o PC:** crie um atalho do `INICIAR_PAINEL_ONLINE.bat`
  dentro de `shell:startup` (Win+R → `shell:startup`).
- **O que a pessoa faz:** digita os nomes (um por linha, até 20), clica em
  *Gerar relatório*, acompanha o status e abre/baixa o HTML e o Excel.
- **Fila:** um trabalho por vez (o Chrome é único). O histórico fica em
  `painel_jobs.json`; se o painel reiniciar no meio, o trabalho é marcado como
  interrompido.
- **Sessão expirada:** o painel não consegue fazer login por você. Se aparecer
  "Sessão do CGD expirada", rode `python cgd_extrator.py` neste PC e entre no
  site; aí os jobs voltam a funcionar.

Enquanto um relatório é gerado, o Chrome abre **na tela deste PC** (é o único
jeito de passar pelo Cloudflare) — quem está usando o link não vê isso, só o
status "Gerando…". Cada aluno leva cerca de 1 minuto.

⚠️ **LGPD:** os relatórios têm nome, e-mail e notas de menores. Use senha forte,
compartilhe o link só com quem precisa e feche a janela do `painel_online.py`
quando terminar (o link morre na hora).

## Publicação no GitHub e Vercel

O painel pode ser publicado no Vercel como interface web, mas a geração real exige
que o processo Flask/Chrome permaneça em execução neste computador. A plataforma
serverless não dispõe da sessão autenticada nem do navegador usado pelo CGD.
Para uso público completo, mantenha `python painel_online.py` ativo neste PC.

## Saída (Excel)

| Aba | Conteúdo |
| --- | --- |
| `Desempenho` | uma linha por curso: aluno, contrato, curso, tipo, datas, e-mail, status da matrícula, **progresso** (itens concluídos/total), **nota**, **situação do curso**, **acessos**, **tempo estudado** e link |
| `Resumo` | totais por aluno e progresso médio |
| `Detalhes` | tabelas lidas na página de desempenho (etapas, aulas e quizzes) |
| `Erros` | o que falhou, por aluno/curso |

Cores: status verde (ativo/concluído), vermelho (expirado/bloqueado/reprovado),
amarelo (pendente); o progresso segue a mesma lógica (≥80% verde, ≥40% amarelo).

## Como o desempenho é lido

A busca acontece em `/alunos?q=<nome>`, que devolve a linha do aluno com o link
`/contratos/<id>`. Nessa página ficam as tabelas **Cursos em andamento** e
**Cursos a fazer** (id `table_cursos_fazer`), com a coluna **Ações**
("Expirado"/"Bloqueado"/"Desempenho").

O link de cada curso abre `/contratos/cursos-ead/<id>/desempenho/`, onde existe a
tabela `#dataTable` com **Etapa | Situação | Notas | Data de realização | Tempo (h)**:

- **progresso** = atividades com situação "Concluído" ÷ total de atividades;
- **itens** = `concluídas/total` (linhas de módulo são ignoradas);
- **nota** = valor da linha **Total** da coluna Notas (nota final do curso);
- **situação** = "Concluído" (100%), "Em andamento" (parcial) ou "Não iniciado";
- **acessos** = linhas do quadro "Resumo da frequência";
- **tempo** = total de horas da linha **Total**.

Se o site mudar, rode com `--debug-html`: o HTML de cada tela é salvo em
`debug_html/` para ajustar os seletores.

## Estrutura do código

```
login/sessão          → iniciar_navegador, garantir_login
busca do aluno        → buscar_contratos
cursos do contrato    → coletar_cursos (lê cabeçalhos, mapeia colunas)
desempenho do curso   → coletar_desempenho, analisar_desempenho_tabelas, extrair_tabelas
saída                 → salvar_excel, salvar_json, salvar_html
painel web            → painel.py (Flask + fila) / painel_online.py (Cloudflare Tunnel)
CLI                   → parse_args, coletar_nomes, main
```

## Correções em relação à versão original

- seletor inválido `table.table-condensed table-striped ...` (retornava 0 linhas) → tabelas lidas por cabeçalho;
- `IndexError` ao acessar a 7ª coluna quando a linha tinha 6;
- contrato do aluno agora é conferido pelo nome (não pega mais o primeiro link qualquer);
- "Desbloquear" não é mais confundido com "Bloquear"; coluna "Ações" não é mais confundida com "Expiração";
- a página de `/desempenho/` passou a ser realmente visitada (antes o link era salvo e ignorado);
- remoção do e-mail/senha sem uso (`getpass`), sessão reaproveitada, `try/finally` no navegador,
  nome de arquivo sanitizado, múltiplos alunos e aba de erros.

O arquivo `cgd_extrator_original.py` é o backup da versão anterior.
