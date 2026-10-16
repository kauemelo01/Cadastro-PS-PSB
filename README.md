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

### Ações em cada cadastro
Cada card exibe três botões lado a lado:

| Botão | Quem pode usar | Descrição |
|---|---|---|
| **✏️ Editar Dados** | Somente administrador | Edição dos campos do cadastro (veja *Edição de registros*) |
| **📦 Marcar mês** | Todos | Registra a entrega do mês atual, com confirmação em popup |
| **👶 Cadastrar Criança** | Todos | Vincula uma criança ao cadastro (veja *Cadastro de crianças*) |

Para usuários sem acesso administrativo, o botão Editar Dados aparece com um cadeado e fica desabilitado.

### Edição de registros
| Funcionalidade | Descrição |
|---|---|
| **Marcar entrega do mês atual** | Botão em cada card para registrar a entrega do mês corrente; exige confirmação em popup. Disponível a todos os usuários. Quando o mês já está marcado, o botão aparece desabilitado como "OK" |
| **Editar Dados** *(admin)* | Popup para editar Tipo, Status, CID 2026, Reserva 1, CPF Res. 1, Reserva 2, CPF Res. 2 e Alerta |
| **Listas suspensas** | Tipo e Status são escolhidos em menus construídos a partir dos valores já existentes na planilha, mantendo a padronização |
| **Gravação seletiva** | Apenas os campos efetivamente alterados são gravados; o botão Salvar fica inativo quando nada mudou |
| **Salvamento automático** | Alterações confirmadas são salvas imediatamente no arquivo `.xlsx` no Google Drive, em uma única operação |
| **Overlay de sessão** | Edições ficam refletidas na tela instantaneamente, sem necessidade de recarregar |

### Cadastro de crianças
O botão **Cadastrar Criança** abre um popup para registrar uma criança vinculada ao cadastro selecionado.

| Campo | Obrigatório | Descrição |
|---|---|---|
| **NÚMERO** | — | Fixo e somente leitura: é sempre o número do cadastro em que o botão foi clicado |
| **NOME** | Sim | Nome da criança |
| **SEXO** | Sim | Feminino, Masculino ou Não informado |
| **IDADE** | Sim | Idade em anos completos (número inteiro, a partir de 0) |
| **OBS.** | Não | Observação livre, em caixa de texto de várias linhas |

Abaixo dos campos há os botões **Salvar** e **Cancelar**. Após salvar, uma mensagem de confirmação aparece no topo da tela, e o formulário abre limpo na vez seguinte.

| Funcionalidade | Descrição |
|---|---|
| **Gravação em JSON** | Os dados são gravados em formato JSON na coluna Z (Tier 2), na linha do cadastro correspondente |
| **Verificação prévia** | Antes de gravar, o app lê o conteúdo atual da célula; se já houver JSON, a nova criança é acrescentada e todas as informações anteriores são preservadas |
| **Várias crianças** | Um mesmo cadastro pode receber quantas crianças forem necessárias no mesmo JSON |
| **Localização pelo NÚMERO** | A linha é localizada pelo NÚMERO no arquivo mais recente do Drive, e não pela posição na tela. Se o número não for encontrado, ou existir em mais de uma linha, nada é alterado |
| **Proteção do conteúdo existente** | Se a coluna Z contiver um valor que não é um JSON válido, o app não o sobrescreve e exibe um aviso |
| **Limite do Excel** | Se o JSON ultrapassar o limite de 32.767 caracteres por célula, o app informa o erro e não altera nenhum dado |
| **Compatibilidade** | JSON já existente em formato de lista é mantido nesse formato |

Os dados das crianças ficam armazenados em Tier 2 e **não são exibidos** nos cards.

Exemplo do conteúdo gravado na coluna Z:

```json
{"criancas":[
  {"NÚMERO":"7","NOME":"Ana","SEXO":"Feminino","IDADE":5,"OBS":"Alergia a amendoim"},
  {"NÚMERO":"7","NOME":"Beto","SEXO":"Masculino","IDADE":9,"OBS":""}
]}
```

*(O JSON é gravado em uma única linha na célula; a quebra acima serve apenas para leitura.)*

Crianças cadastradas antes da inclusão do campo OBS permanecem inalteradas, sem essa chave.

### Modo Administrador
| Funcionalidade | Descrição |
|---|---|
| **Acesso restrito** | Link discreto no rodapé abre um popup de senha; a senha é lida exclusivamente dos Secrets (`ADMIN_PASSWORD`) |
| **Sem senha configurada** | Se o secret não estiver definido, o acesso administrativo fica indisponível — não há senha padrão |
| **Sessão** | Enquanto ativo, um selo "Modo administrador ativo" aparece no rodapé, com botão de saída |
| **Recursos protegidos** | Editar Dados e Cestas Extras só ficam disponíveis em modo administrador. Marcar mês e Cadastrar Criança são abertos a todos |

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

A **coluna Z** é um campo Tier 2 reservado ao cadastro de crianças: ela precisa estar identificada como `Tier 2` na linha 1, e seu conteúdo é um JSON com a lista de crianças do cadastro (veja *Cadastro de crianças*).

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
