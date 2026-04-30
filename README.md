# 📋 Pastoral Social — Consulta de Cadastro

Aplicativo mobile para busca e gestão de registros do cadastro da Pastoral Social,
construído com **Streamlit** e hospedado no **Streamlit Community Cloud**.

---

## 📱 Funcionalidades

### Busca
| Funcionalidade | Descrição |
|---|---|
| **Busca por Número** | Localiza um registro pelo número exato |
| **Busca geral** | Pesquisa em Nome, CPF, Reserva 1 e 2 e seus CPFs |
| **Busca parcial** | Funciona com fragmentos (ex: "silva", "maria") |
| **Sem acento** | "Antonio" encontra "Antônio" e vice-versa |
| **CPF flexível** | `06904914840` encontra `069.049.148-40` e vice-versa |
| **Lógica OR** | As duas barras de busca podem ser usadas simultaneamente — os resultados são unidos |
| **Destaque** | Os termos buscados são destacados nos campos correspondentes dos resultados |
| **Ordenação** | Correspondências por Nome ou CPF aparecem primeiro |

### Visualização
| Funcionalidade | Descrição |
|---|---|
| **Dados Tier 1** | Tipo, Status, CID 2026, Reserva 1 e 2, Alerta para a Mesa |
| **Alerta** | Campo de alerta exibido em vermelho quando preenchido |
| **Histórico de entregas** | Grade compacta com os meses do ano; células verdes indicam entrega confirmada |
| **Dados Tier 2** | Informa que campos adicionais existem, sem exibi-los |

### Edição
| Funcionalidade | Descrição |
|---|---|
| **Marcar entrega do mês** | Botão por registro para confirmar a entrega do mês corrente, com popup de confirmação |
| **Editar CID 2026** | Botão por registro para adicionar ou alterar o CID, com popup de edição |
| **Salvamento automático** | Alterações confirmadas são salvas diretamente no arquivo `.xlsx` no Google Drive |
| **Atualização em tempo real** | O cache é invalidado após cada edição; o próximo carregamento reflete a alteração |
