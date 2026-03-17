# Escola de Dons — Histórico de Versões

---

## v3.1.0 — 2026-03-17

### Correções
- **Relatório por Dons** (`admin.html`): a aba "Dons" ficava sempre vazia porque os campos `don1`, `don2`, `don3` nunca eram preenchidos. Corrigido em três frentes:
  - `form.html`: após calcular o resultado, envia automaticamente os top 3 dons (Alfa) para o backend via nova ação `saveDons`
  - `backend.gs`: adicionada ação `saveDons` no `doPost` e função correspondente que grava `don1/2/3` na planilha `usuarios`
  - `backend.gs`: `releaseAndSaveSheets` agora também salva `don1/2/3` ao liberar o resultado de cada pessoa
  - `admin.html`: `loadFullDons()` força recarregamento dos dados em vez de usar cache desatualizado
- **Bloqueio Geral — barra "Verificando..."** (`admin.html`): `loadBloqueioGeral()` não era chamada após o login. Corrigido — o status é carregado imediatamente ao entrar no painel
- **Bloqueio Geral — envio de respostas** (`form.html`): `sendCurrentInv()` não verificava `_READONLY`. Corrigido — envio é bloqueado mesmo se o formulário for acessado por outros meios
- **Bloqueio Geral — feedback visual** (`form.html`): adicionado banner laranja na tela home explicando que o sistema está em modo leitura e que o resultado pode ser visualizado normalmente
- **Backend**: removida entrada `case 'getAnswers'` duplicada no `doPost`

### Novidades
- **Backfill de dons** (`backend.gs`): função `backfillDons()` para preencher `don1/2/3` de todos os usuários já cadastrados antes da v3.1.0. Executar manualmente uma única vez via Google Apps Script
- **Versão do sistema**: número de versão exibido discretamente em `index.html` (rodapé), `form.html` (rodapé da tela home) e `admin.html` (barra superior)

---

## v3.0.0 — 2026 (versão inicial em produção)

### Sistema completo implantado
- Backend Google Apps Script com planilha Google Sheets como banco de dados
- Inventário Alfa (110 questões) e Inventário Ômega (138 questões)
- Cálculo automático de Top 3 dons por inventário e resultado combinado (Alfa + Ômega)
- Sistema de códigos únicos de 6 caracteres para acesso sem senha
- Cadastro com nome e telefone
- Salvamento incremental de respostas por inventário (respostas_alfa / respostas_omega)
- Resultado liberado manualmente pelo administrador
- Gravação automática nas abas "Resultado Alfa", "Resultado Ômega" e "Comparativo" ao liberar

### Painel Administrativo
- Login com senha
- Visão geral: total de pessoas, completos, parciais, liberados
- Gráfico "Top Dons Identificados" na visão geral
- Listagem de pessoas com busca e filtros (completos, parciais, sem respostas, liberados, pendentes)
- Liberar / bloquear resultado individualmente
- Liberar múltiplos resultados de uma vez (bulk release)
- Bloquear / desbloquear preenchimento individualmente por pessoa
- Bloqueio Geral: suspende novos cadastros e edição de formulários para todos
- Exportar dados para Excel (CSV com resultados completos)
- Ranking por Dom: tabela detalhada por don com pontuação por pessoa
- Aba "Dons": relatório agrupado por dom com lista de participantes
- Configurações: nome da igreja, logo, imagem de fundo, senha admin
- Upload de imagens direto para Google Drive

### Formulário do Participante
- Tela home com progresso visual de cada inventário
- Perguntas com escala 0–3 (Nunca / Raro / Às vezes / Sempre)
- Salvamento local (localStorage) como backup em caso de queda
- Resultado com Top 3 combinado, lista completa de dons ordenada, barra de pontuação
- Resultado bloqueado enquanto não liberado pelo admin (tela "em análise")

---

> Para atualizar a versão: altere o número em `index.html`, `form.html`, `admin.html` e adicione uma entrada neste arquivo.
