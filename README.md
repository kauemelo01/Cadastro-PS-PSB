# 📋 Pastoral Social — Consulta de Cadastro

Aplicativo mobile para busca de registros do cadastro da Pastoral Social,
construído com **Streamlit** e hospedado gratuitamente no **Streamlit Community Cloud**.

---

## 🚀 Como configurar

### 1 · Prepare seu repositório no GitHub

Crie um repositório (pode ser privado) com esta estrutura:

```
meu-repo/
├── app.py
├── requirements.txt
├── README.md
└── cadastro.xlsx          ← seu arquivo de dados
```

### 2 · Configure a URL do arquivo

Abra `app.py` e substitua a linha:

```python
GITHUB_RAW_URL = (
    "https://raw.githubusercontent.com/SEU_USUARIO/SEU_REPO/main/cadastro.xlsx"
)
```

com o endereço **raw** do seu arquivo. Para encontrar este endereço:
1. Acesse o arquivo `cadastro.xlsx` no GitHub
2. Clique em **Raw**
3. Copie a URL da barra de endereços

**Formato:** `https://raw.githubusercontent.com/{usuario}/{repo}/{branch}/{arquivo}.xlsx`

> **Dica de segurança:** Se o seu repositório for privado, use o
> **Streamlit Secrets** em vez de deixar a URL no código.
> Veja a seção "Repositório Privado" abaixo.

### 3 · Faça o deploy no Streamlit Community Cloud

1. Acesse [share.streamlit.io](https://share.streamlit.io) e entre com sua conta GitHub
2. Clique em **New app**
3. Selecione seu repositório, branch (`main`) e o arquivo `app.py`
4. Clique em **Deploy**

Pronto! O app estará disponível em um link público do tipo `https://SEU_USUARIO-SEU_REPO-app-XXXX.streamlit.app`.

---

## 🔒 Repositório Privado (recomendado para dados sensíveis)

Se o repositório for privado, a URL raw não funcionará diretamente. Use o
[Streamlit Secrets](https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management):

1. No painel do Streamlit Cloud, acesse **Settings → Secrets**
2. Adicione:
   ```toml
   GITHUB_RAW_URL = "https://raw.githubusercontent.com/SEU_USUARIO/SEU_REPO/main/cadastro.xlsx"
   ```
3. Para repositórios privados, você precisará de um **Personal Access Token (PAT)**
   do GitHub com permissão `repo`. Substitua a URL pelo formato:
   ```
   https://SEU_TOKEN@raw.githubusercontent.com/SEU_USUARIO/SEU_REPO/main/cadastro.xlsx
   ```
   e guarde apenas nos Secrets, nunca no código.

---

## 📱 Funcionalidades

| Funcionalidade | Descrição |
|---|---|
| **Busca unificada** | Pesquisa simultânea em Nome, CPF, Número, Reserva 1 e 2 |
| **Busca parcial** | Funciona com fragmentos (ex: "silva", "123", "maria") |
| **Dados Tier 1** | Tipo, Status, CID 2026, Alerta para a Mesa, histórico de entregas mensais |
| **Hint Tier 2** | Informa que campos adicionais existem, sem exibi-los |
| **Cache** | Dados em cache por 5 minutos para evitar requisições excessivas |
| **Mobile-first** | Layout otimizado para telas pequenas |

---

## 🗂️ Estrutura do arquivo xlsx

O app espera que o arquivo tenha **duas linhas de cabeçalho**:

| Linha | Conteúdo |
|---|---|
| Linha 1 | Tags: `SEARCH`, `Tier 1` ou `Tier 2` |
| Linha 2 | Nomes das colunas |
| Linha 3+ | Dados dos registros |

**Campos SEARCH** (buscáveis): NUMERO, CPF, NOME, RESERVA 1, CPF RESERVA 1, RESERVA 2, CPF RESERVA 2

**Campos Tier 1** (visíveis para todos): TIPO, CID 2026, STATUS, ALERTA PARA A MESA, colunas mensais de entrega

**Campos Tier 2** (restritos): endereço, telefone, observações internas, dados pessoais, etc.

---

## 🔄 Atualizar os dados

Basta substituir o arquivo `cadastro.xlsx` no repositório GitHub.
O app recarrega os dados automaticamente a cada 5 minutos (ou ao recarregar a página).

---

## 🛠️ Desenvolvimento local

```bash
pip install -r requirements.txt
streamlit run app.py
```
