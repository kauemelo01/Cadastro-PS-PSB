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
| **Modo escuro** | Cores adaptadas automaticamente ao tema do sistema |

### Modo Administrador
| Funcionalidade | Descrição |
|---|---|
| **Acesso restrito** | Link discreto no rodapé abre um popup de senha; a senha é lida exclusivamente dos Secrets (`ADMIN_PASSWORD`) |
| **Sem senha configurada** | Se o secret não estiver definido, o acesso administrativo fica indisponível — não há senha padrão |
| **Sessão** | Enquanto ativo, um selo "Modo administrador ativo" aparece no rodapé, com botão de saída |
| **Recursos protegidos** | Edição de dados e Cestas Extras só ficam disponíveis em modo administrador |

### Edição de registros
| Funcionalidade | Descrição |
|---|---|
| **Marcar entrega do mês atual** | Botão em cada card para registrar a entrega do mês corrente; exige confirmação em popup. Disponível a todos os usuários |
| **Editar Dados** *(admin)* | Popup para editar Tipo, Status, CID 2026, Reserva 1, CPF Res. 1, Reserva 2, CPF Res. 2 e Alerta |
| **Listas suspensas** | Tipo e Status são escolhidos em menus construídos a partir dos valores já existentes na planilha, mantendo a padronização |
| **Gravação seletiva** | Apenas os campos efetivamente alterados são gravados; o botão Salvar fica inativo quando nada mudou |
| **Salvamento automático** | Alterações confirmadas são salvas imediatamente no arquivo `.xlsx` no Google Drive, em uma única operação |
| **Overlay de sessão** | Edições ficam refletidas na tela instantaneamente, sem necessidade de recarregar |

### Cestas Extras *(admin)*
| Funcionalidade | Descrição |
|---|---|
| **Dar Cesta Extra** | Popup com os campos Número, Nome, CPF, Telefone e Motivo |
| **Numeração automática** | O Número é preenchido com o próximo valor da sequência (maior número existente + 1), mas permanece editável |
| **Campos obrigatórios** | Número, Nome e Motivo são exigidos; o Número aceita apenas dígitos |
| **Lista de Cestas Extras** | Tela dedicada com todos os registros, exibindo selo Entregue/Pendente e um resumo com totais |
| **Marcar como entregue** | Botão em cada registro, com popup de confirmação; grava `OK` na coluna F da aba Sheet2 |
| **Atualização automática** | A lista se atualiza sozinha a cada minuto |

---

## 🗂️ Estrutura do arquivo .xlsx

### Sheet1 — cadastro principal

O app espera que a primeira aba tenha **duas linhas de cabeçalho**:

| Linha | Conteúdo |
|---|---|
| Linha 1 | Tags: `SEARCH`, `Tier 1` ou `Tier 2` |
| Linha 2 | Nomes das colunas |
| Linha 3+ | Dados dos registros |

**Campos SEARCH** (buscáveis): NUMERO, CPF, NOME, RESERVA 1, CPF RESERVA 1, RESERVA 2, CPF RESERVA 2

**Campos Tier 1** (visíveis para todos): TIPO, CID 2026, STATUS, ALERTA PARA A MESA, colunas mensais de entrega

**Campos Tier 2** (restritos): endereço, telefone, observações internas, dados pessoais, etc.

As colunas mensais são reconhecidas automaticamente a partir de nomes de coluna com data e exibidas no formato `Mmm/AA` (ex: `Mai/26`).

### Sheet2 — cestas extras

Uma única linha de cabeçalho:

| Coluna | Conteúdo |
|---|---|
| A | NÚMERO |
| B | NOME |
| C | CPF |
| D | TELEFONE |
| E | MOTIVO |
| F | ENTREGUE (`OK` quando a cesta é entregue) |

A coluna F é criada automaticamente pelo app na primeira entrega registrada.
