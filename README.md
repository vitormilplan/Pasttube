# Painel de liberação de juntas — M.449

Site estático que cruza **Takeoff (linhas 100%) × SGJ × Lista de Pastas × Rede de Precedência** e entrega:

- indicadores (juntas previstas / liberadas / pendentes, linhas que podem ser executadas agora);
- **escada de prioridades** pela Rede de Precedência (quem pode executar agora e quem está aguardando etapa anterior);
- gráficos por etapa, por status da junta e por pasta;
- tabela de linhas com busca, filtros e detalhe das juntas pendentes;
- alertas automáticos (linha 100% sem juntas no SGJ, pasta diferente entre takeoff e lista, subsistema fora da rede, quantidade de juntas diferente entre takeoff e SGJ);
- exportação para Excel (abas RESUMO, PRIORIDADES e JUNTAS PENDENTES).

**Os arquivos são lidos só no navegador de quem usa.** Nada é enviado para servidor — o GitHub guarda apenas o código.

---

## Como subir no GitHub (sem precisar de git)

1. Entre em <https://github.com> → **New repository**.
   - Nome sugerido: `painel-juntas`.
   - Visibilidade: **Public** (o GitHub Pages gratuito só funciona em repositório público; os dados da obra **não** ficam no repositório, só o código).
2. Na página do repositório novo, clique em **uploading an existing file**.
3. Arraste **o conteúdo** desta pasta (o `index.html`, as pastas `css` e `js`, o `README.md` e o `.nojekyll`) e clique em **Commit changes**.
   > Arraste os arquivos e pastas que estão *dentro* de `painel-juntas`, não a pasta em si — o `index.html` precisa ficar na raiz do repositório.
4. Vá em **Settings → Pages**. Em *Build and deployment*, escolha **Source: Deploy from a branch**, **Branch: main**, pasta **/(root)** e **Save**.
5. Em 1–2 minutos aparece o link, no formato `https://SEU-USUARIO.github.io/painel-juntas/`.

### Sem GitHub Pages (repositório privado ou uso local)

Abra o `index.html` direto no Chrome/Edge (duplo clique). Funciona igual — só precisa de internet para baixar as duas bibliotecas (SheetJS e Chart.js).

---

## Como usar

1. **Importar** — arraste os três arquivos:
   | Arquivo | Aba | O que é usado |
   |---|---|---|
   | Takeoff de tubulação (.xlsb/.xlsx) | `PROG  TUB` | Linhas-mãe (sem spool, com “-” na coluna O). Linha = G, pasta = F, juntas previstas/realizadas = AO/AP. Entra quem tem realizadas/previstas = 100%. |
   | SGJ (Situação Geral de Juntas) | `SGJ` | Linha = C, junta = L, status = CP. Liberada = “Junta Liberada”. |
   | Lista de Pastas | `MILPLAN` | Achadas pelo cabeçalho: *Line Code*, *N° DA PASTA - PARA*, *Subsystem/ SOP*. |
2. **Processar** → abre o painel.
3. Clique num degrau da escada para ver as linhas daquela etapa. Na aba **Linhas**, filtre **Situação = PODE EXECUTAR** para ter a lista de trabalho.
4. **Exportar Excel** quando quiser mandar para alguém.

A comparação dos códigos de linha ignora espaços e trata `1.1/2"` = `1 1/2"` e `''` = `"`.

### Situação de cada linha

| Situação | Significado |
|---|---|
| **PODE EXECUTAR** | Tem juntas pendentes e nenhuma etapa anterior da rede está com pendência. |
| **AGUARDANDO PRIORIDADE N** | Tem juntas pendentes, mas a etapa N (anterior) ainda tem pendência. |
| **OK** | Todas as juntas liberadas no SGJ. |
| **SEM JUNTAS NO SGJ** | Linha 100% no takeoff, mas nenhuma junta encontrada no SGJ. |
| **VERIFICAR** | Tem pendência, mas a linha não tem subsistema ou o subsistema não está na rede. |

---

## Rede de Precedência

O PDF da rede é imagem, então ela fica transcrita numa tabela (`js/rede-padrao.js`, Rev02 de 16/11/2026).
Cada subsistema tem uma **etapa** — etapas menores vêm antes; mesma etapa = em paralelo.

- Para ajustar no dia a dia: aba **Rede de precedência** → editar → **Salvar** (fica salvo no navegador). Use **Exportar JSON** para mandar para outra pessoa e **Importar JSON** para carregar.
- Para mudar o padrão de todo mundo: edite `js/rede-padrao.js` no GitHub (ícone de lápis) e faça commit.

Subsistemas que aparecem nas linhas mas não estão na rede são listados nos alertas do painel.

## Se o layout de alguma planilha mudar

Aba **Configurações**: troque as letras das colunas, nome da aba, primeira linha de dados ou o texto do status liberado. Fica salvo no navegador.

## Estrutura

```
index.html           página
css/style.css        visual
js/core.js           regras de cruzamento (sem interface — dá para testar no Node)
js/rede-padrao.js    rede de precedência transcrita
js/app.js            interface, gráficos e exportação
```
