# Operational Metrics Engine (Google Sheets + Apps Script)

> Lightweight operational workflow engine built on Google Sheets + Apps Script.  
> Designed for teams that need structured metrics without a full BI stack.

A plug-and-play Google Sheets + Apps Script framework that converts raw tabular data into a structured, team-ready operational board with persistent history, validations, protections, and optional monthly metric freezing.

---

## 🚀 English Version

### Overview

The Operational Metrics Engine is a configurable workflow layer that transforms raw input data into a structured operational board for teams.

It provides:

- Persistent **HISTORY** by unique record ID  
- A protected **TEAM_BOARD** with editable fields and validations  
- Conditional formatting based on priority logic  
- Optional **monthly snapshot** mechanism  
- A basic **DASHBOARD** and execution **LOGS**

This solution is ideal for teams that need operational structure without implementing a full BI or CRM stack.

---

### Typical Use Cases

- Sales / SDR operations (lead management)
- Customer Support (ticket workflows)
- Customer Success (portfolio tracking)
- Collections / Financial operations
- Backoffice task management
- Internal quality control tracking

---

### How It Works (Pipeline)

1. Paste or import your dataset into **INPUT_DATA**
2. Configure mappings and rules in **CONFIG**
3. Run **Operational Engine → Daily Update**
4. The team operates inside **TEAM_BOARD**
5. Edits automatically sync back to **HISTORY**
6. (Optional) Run **Monthly Snapshot** to freeze previous-period metrics

---

### Included Sheets

- `ABOUT` — Quick instructions and license info
- `CONFIG` — Column mappings and operational rules
- `INPUT_DATA` — Raw dataset
- `HISTORY` — Persistent operational layer
- `TEAM_BOARD` — Editable team-facing board
- `DASHBOARD` — KPI summary view
- `LOGS` — Execution audit trail

---

### Configuration (Core Customization)

All customization happens inside the **CONFIG** sheet.

#### Column Mapping

- `MAP_ID` — Unique identifier column (e.g., lead_id, ticket_id, CNPJ)
- `MAP_OWNER` — Owner/assignee column
- `MAP_SUBJECT` — Subject or name column
- `MAP_CREATED_AT` — Created date column
- `MAP_PRIORITY` — Priority/score column (recommended 0–4 scale)

#### Operational Behavior

- `EDITABLE_FIELDS` — Editable fields (semicolon-separated)
- `STATUS_OPTIONS` — Allowed STATUS values (semicolon-separated)
- `PROTECT_NON_EDITABLE` — YES/NO
- `DAILY_OVERWRITE_OWNER` — YES/NO

#### Essential Columns

- `ESSENTIAL_COLUMNS` — Additional columns to include
- `ESSENTIAL_BY_HEADER_COLOR` — YES/NO
- `ESSENTIAL_COLOR_HEX`
- `COLOR_TOLERANCE`

---

### Installation

1. Create a new Google Sheet
2. Go to Extensions → Apps Script
3. Paste the contents of `Code.gs`
4. Save and run `Install Structure`
5. Configure the `CONFIG` sheet
6. Paste your dataset into `INPUT_DATA`
7. Run `Daily Update`

---

### Architecture Overview

The engine follows a layered architecture:

INPUT_DATA  
→ Validation & Normalization  
→ HISTORY (persistent state)  
→ TEAM_BOARD (operational layer)  
→ DASHBOARD + LOGS  

This structure separates raw data from operational logic and team interaction.

---

### License

This project is dual-licensed:

- AGPL-3.0 (open-source use) — see `LICENSE`
- Commercial License (for proprietary/closed-source distribution) — see `COMMERCIAL_LICENSE.md`

---

### Author

Eduardo Sousa

---

## 🇧🇷 Versão em Português

### Visão Geral

O Operational Metrics Engine é uma camada operacional configurável que transforma dados brutos em um board estruturado para equipes.

Ele oferece:

- **HISTORY** persistente por ID único  
- **TEAM_BOARD** protegido com campos editáveis e validações  
- Formatação condicional baseada em prioridade  
- Mecanismo opcional de **snapshot mensal**  
- **DASHBOARD** simples e **LOGS** de execução  

Ideal para equipes que precisam de organização operacional sem implementar uma stack completa de BI ou CRM.

---

### Casos de Uso

- Operação Comercial / SDR (gestão de leads)
- Suporte (fluxo de tickets)
- Customer Success (gestão de carteira)
- Cobrança / Operações financeiras
- Backoffice (gestão de tarefas)
- Controle interno de qualidade

---

### Como Funciona (Fluxo)

1. Cole ou importe os dados em **INPUT_DATA**
2. Configure os mapeamentos e regras na aba **CONFIG**
3. Execute **Operational Engine → Daily Update**
4. O time trabalha dentro do **TEAM_BOARD**
5. As edições são sincronizadas automaticamente para o **HISTORY**
6. (Opcional) Execute **Monthly Snapshot** para congelar métricas do período anterior

---

### Abas Incluídas

- `ABOUT` — Instruções rápidas e informações de licença
- `CONFIG` — Mapeamentos e regras operacionais
- `INPUT_DATA` — Base bruta
- `HISTORY` — Camada persistente
- `TEAM_BOARD` — Board operacional editável
- `DASHBOARD` — Resumo de indicadores
- `LOGS` — Auditoria de execuções

---

### Configuração

Toda a personalização é feita na aba **CONFIG**, sem necessidade de alterar o código principal.

#### Mapeamento de Colunas

- `MAP_ID` — Identificador único (ex.: lead_id, ticket_id, CNPJ)
- `MAP_OWNER` — Responsável
- `MAP_SUBJECT` — Nome ou assunto
- `MAP_CREATED_AT` — Data de criação
- `MAP_PRIORITY` — Prioridade/Score (escala recomendada 0–4)

#### Comportamento Operacional

- `EDITABLE_FIELDS` — Campos editáveis (separados por ;)
- `STATUS_OPTIONS` — Valores permitidos para STATUS (separados por ;)
- `PROTECT_NON_EDITABLE` — YES/NO
- `DAILY_OVERWRITE_OWNER` — YES/NO

#### Colunas Essenciais

- `ESSENTIAL_COLUMNS` — Colunas adicionais no board
- `ESSENTIAL_BY_HEADER_COLOR` — YES/NO
- `ESSENTIAL_COLOR_HEX`
- `COLOR_TOLERANCE`

---

### Licença

Licença dupla:

- AGPL-3.0 para uso open-source (ver `LICENSE`)
- Licença comercial para distribuição proprietária (ver `COMMERCIAL_LICENSE.md`)

---

### Autor

Eduardo Sousa
