# 📋 Pastoral Social — Consulta de Cadastro

Aplicativo mobile para busca de registros do cadastro da Pastoral Social,
construído com **Streamlit** e hospedado gratuitamente no **Streamlit Community Cloud**.

---

## 📱 Funcionalidades

| Funcionalidade | Descrição |
|---|---|
| **Busca unificada** | Pesquisa simultânea em Nome, CPF, Número, Reserva 1 e 2 |
| **Busca parcial** | Funciona com fragmentos (ex: "silva", "123", "maria") |
| **Dados Tier 1** | Tipo, Status, CID 2026, Alerta para a Mesa, histórico de entregas mensais |
| **Dados Tier 2** | Informa que campos adicionais existem, sem exibi-los. WIP. |
| **Cache** | Dados em cache por 5 minutos para evitar requisições excessivas |
| **Mobile-first** | Layout otimizado para telas pequenas |

---

## 🗂️ Estrutura do arquivo .xlsx

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
streamlit run main.py
```