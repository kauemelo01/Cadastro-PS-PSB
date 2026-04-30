# 📋 Pastoral Social — Consulta de Cadastro

Aplicativo mobile para busca e gestão de registros do cadastro da Pastoral Social,
construído com **Streamlit** e hospedado no **Streamlit Community Cloud**.

---

## 📱 Funcionalidades

### Busca
| Funcionalidade | Descrição |
|---|---|
| **Busca por Número** | Campo dedicado com correspondência exata — buscar "1" retorna apenas o registro cujo número é "1" |
| **Busca geral** | Pesquisa simultânea em Nome, CPF, Reserva 1, CPF Reserva 1, Reserva 2 e CPF Reserva 2 |
| **Lógica OR** | Ambas as barras de busca podem ser usadas ao mesmo tempo; os resultados são a união dos dois conjuntos |
| **Busca parcial** | Funciona com fragmentos (ex: "silva", "maria") |
| **Insensível a acentos** | "Antonio" encontra "Antônio" e vice-versa |
| **Insensível a formatação de CPF** | "06904914840" encontra "069.049.148-40" e vice-versa |
| **Destaque de correspondências** | Os termos buscados são destacados em amarelo nos campos exibidos |
| **Prioridade de resultados** | Correspondências em Nome ou CPF aparecem primeiro na lista |

### Exibição dos registros
| Funcionalidade | Descrição |
|---|---|
| **Cabeçalho do card** | Nome, Número e CPF exibidos com destaque |
| **Campos Tier 1** | Tipo, Status, CID 2026, Alerta para a Mesa, Reserva 1, CPF Reserva 1, Reserva 2, CPF Reserva 2 |
| **Alerta para a Mesa** | Exibido em vermelho com texto ligeiramente maior quando presente |
| **Histórico de entregas** | Grade compacta com 6 colunas — células verdes para meses com entrega confirmada (OK), cinza para os demais |
| **Campos Tier 2** | Indica que campos adicionais existem sem exibi-los |

### Edição
| Funcionalidade | Descrição |
|---|---|
| **Marcar entrega do mês atual** | Botão em cada card para registrar a entrega do mês corrente; exige confirmação em popup antes de salvar |
| **Editar CID 2026** | Botão dedicado abre um popup para editar o campo CID diretamente no app |
| **Salvamento automático** | Alterações confirmadas são salvas imediatamente no arquivo `.xlsx` no Google Drive |
| **Overlay de sessão** | Edições ficam refletidas na tela instantaneamente, sem necessidade de recarregar |

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
