# Cobre Seu Político — Contatos parlamentares automáticos

Coleta semanalmente os contatos **institucionais e públicos** de todos os
deputados federais e senadores em exercício, grava numa planilha do Google
Sheets e alimenta um dashboard no Looker Studio embutido no seu Notion.

Tudo roda sozinho no GitHub Actions — você não precisa deixar computador ligado.

---

## Visão geral da arquitetura

```
  APIs Dados Abertos            GitHub Actions            Google              Notion
  (Câmara + Senado)   ──────►   (roda toda segunda)  ──►  Sheets  ──►  Looker ──► embed
                                 atualizar_contatos.py     (base)      Studio    no site
```

---

## O que é coletado (e o que NÃO é)

**Coletado** — dados funcionais já publicados pelos próprios órgãos (LAI):
e-mail de gabinete (`@camara.leg.br` / `@senado.leg.br`), telefone e
localização do gabinete, endereço institucional, link do perfil oficial.

**Nunca coletado** — celular pessoal, e-mail particular, endereço residencial.
O script simplesmente não acessa esses campos.

---

## Passo a passo da configuração (uma vez só)

### 1. Criar a conta de serviço do Google (dá acesso à planilha)

1. Acesse <https://console.cloud.google.com/> e crie um projeto (ex.: "cobre-politico").
2. No menu, vá em **APIs e serviços → Biblioteca** e ative:
   - **Google Sheets API**
   - **Google Drive API**
3. Vá em **APIs e serviços → Credenciais → Criar credenciais → Conta de serviço**.
   Dê um nome e conclua.
4. Abra a conta de serviço criada → aba **Chaves → Adicionar chave → Criar nova
   chave → JSON**. Um arquivo `.json` será baixado. **Guarde-o.**
5. Abra esse JSON e copie o valor do campo `client_email`
   (algo como `cobre-politico@...iam.gserviceaccount.com`).

### 2. Preparar a planilha

Você tem duas opções:

- **Deixar o script criar sozinho** (mais simples): não faça nada agora. Na
  primeira execução, a planilha "Cobre Seu Político - Contatos" será criada
  automaticamente no Drive da conta de serviço. Depois você a compartilha com
  seu e-mail pessoal (role de editor) para conseguir vê-la.
- **Criar você mesmo**: crie uma planilha no Google Sheets com esse nome e
  **compartilhe com o `client_email`** da conta de serviço (role de Editor).

### 3. Subir o projeto no GitHub

1. Crie um repositório novo (pode ser privado).
2. Suba estes arquivos:
   - `atualizar_contatos.py`
   - `requirements.txt`
   - `.github/workflows/atualizar.yml`
3. No repositório, vá em **Settings → Secrets and variables → Actions →
   New repository secret**:
   - Nome: `GOOGLE_CREDS`
   - Valor: **cole todo o conteúdo do arquivo JSON** baixado no passo 1.

### 4. Testar

Vá na aba **Actions** do repositório → workflow "Atualizar contatos
parlamentares" → **Run workflow**. Em 1–3 minutos a planilha estará populada.
A partir daí ele roda sozinho toda segunda-feira de manhã.

---

## Montar o dashboard no Looker Studio

1. Acesse <https://lookerstudio.google.com/> → **Criar → Relatório**.
2. Conecte a fonte **Planilhas Google** → selecione "Cobre Seu Político - Contatos"
   → aba principal.
3. Adicione uma **Tabela** com as colunas: Nome, Partido, UF, E-mail oficial,
   Telefone gabinete, Página oficial.
4. Adicione **controles de filtro** (menu suspenso) para **UF** e **Partido** —
   assim o cidadão filtra o próprio estado.
5. **Compartilhar → Gerenciar acesso →** deixe "Qualquer pessoa com o link pode ver".
6. Copie o link do relatório.

## Embutir no Notion

Na página do seu site Notion, digite `/embed`, cole o link do Looker Studio e
ajuste a altura. O dashboard aparece dentro da página, já filtrável e sempre
sincronizado com a planilha.

> Dica: uma nota curta acima do painel — *"Use estes contatos com respeito e
> responsabilidade cívica"* — reforça a seriedade do projeto e o protege.

---

## Manutenção

- Para mudar o dia/horário, edite a linha `cron` em `.github/workflows/atualizar.yml`.
  (`0 9 * * 1` = segunda 09:00 UTC. Use <https://crontab.guru> para montar outros.)
- Se a estrutura de alguma API mudar, o script avisa no log da aba Actions.
- O plano gratuito do GitHub Actions cobre folgadamente uma execução semanal.
