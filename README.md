# 📋 Pastoral Social — Consulta de Cadastro

Aplicativo mobile para consulta e gestão do cadastro de beneficiários da Pastoral
Social, construído com **Streamlit**. Lê e grava um arquivo `.xlsx` hospedado no
**Google Drive**, que permanece a fonte da verdade.

---

## 🔍 Busca

| Funcionalidade | Descrição |
|---|---|
| **Duas barras de busca** | Uma dedicada ao número do registro, outra para os demais campos |
| **Busca por número** | Correspondência **exata** — buscar "1" retorna apenas o registro nº 1 |
| **Busca geral** | Nome, CPF, Reserva 1, CPF Reserva 1, Reserva 2 e CPF Reserva 2 |
| **Lógica OR** | As duas barras podem ser usadas ao mesmo tempo; o resultado é a união |
| **Busca parcial** | Funciona com fragmentos ("silva", "maria") |
| **Insensível a acentos** | "Antonio" encontra "Antônio" e vice-versa |
| **Insensível a pontuação de CPF** | "06904914840" encontra "069.049.148-40" e vice-versa |
| **Destaque das correspondências** | Os termos buscados são realçados em amarelo nos campos exibidos |
| **Prioridade nos resultados** | Correspondências em Nome ou CPF aparecem primeiro |

Na busca por número, o destaque aparece somente no cabeçalho do número, para não
sugerir que o registro casou por outro campo. O nome das crianças não é realçado,
já que a coluna de crianças não é um campo de busca.

---

## 📇 Exibição dos registros

| Funcionalidade | Descrição |
|---|---|
| **Cabeçalho do card** | Nome, Nº e CPF |
| **Campos Tier 1** | Tipo, Status, CID 2026, Reserva 1, CPF Res. 1, Reserva 2, CPF Res. 2 |
| **Alerta para a Mesa** | Em vermelho e com fonte ligeiramente maior quando presente |
| **Histórico de entregas** | Grade compacta de 6 colunas; célula verde para o mês entregue, cinza para os demais |
| **Crianças** | Chips com emoji por sexo (👧 / 👦 / 🧒), nome e idade — aparece só quando há crianças |
| **Modo escuro** | Paleta própria para tema escuro, incluindo os campos de busca |
| **Mobile-first** | Layout otimizado para telas pequenas |

---

## ✏️ Edição

### Disponível para todos

| Funcionalidade | Descrição |
|---|---|
| **Marcar entrega do mês** | Registra "OK" no mês corrente, com popup de confirmação. Fica desabilitado quando já entregue |
| **Cadastrar Criança** | Nome, sexo, idade e observação. Aparece quando o registro não tem crianças |
| **Editar Crianças (N)** | Carrega as crianças cadastradas para editar, remover ou acrescentar — tudo numa única gravação |

### Exclusivo do modo administrador

| Funcionalidade | Descrição |
|---|---|
| **Editar Dados** | Tipo e Status (listas suspensas), CID 2026, Reserva 1 e 2, CPF das reservas e Alerta para a Mesa |

Somente os campos efetivamente alterados são gravados, e o botão Salvar fica
desabilitado quando nada mudou.

---

## 🧺 Cestas Extras (exclusivo do administrador)

| Funcionalidade | Descrição |
|---|---|
| **Dar Cesta Extra** | Número sequencial pré-preenchido e editável; Número, Nome e Motivo obrigatórios |
| **Lista de Cestas Extras** | Tela própria com todos os registros, contagem de entregues e pendentes |
| **Marcar como entregue** | Com popup de confirmação; grava na coluna ENTREGUE |
| **Propagação instantânea** | Uma cesta lançada num dispositivo aparece nos outros em segundos, sem releitura do Drive |
| **Indicador de sincronização** | Mostra se há gravações em andamento e o horário da última sincronização |
| **Recarregar do Drive** | Puxa alterações feitas diretamente na planilha |
| **Recuperação de falhas** | Se a gravação falhar, o dado permanece visível e há botão para tentar novamente |

As cestas extras vivem na aba `Sheet2` e são independentes do cadastro principal.

---

## 🔐 Modo administrador

Um link discreto no rodapé abre a solicitação de senha, lida dos *Secrets* do app.
Enquanto a senha não estiver configurada, o acesso administrativo fica indisponível
em vez de liberado. Ao sair do modo administrador, telas restritas são fechadas
automaticamente.

Vale registrar o alcance dessa proteção: ela organiza quem faz o quê, mas não é uma
barreira de segurança — a senha vive no processo do app e não há limite de
tentativas. O que realmente protege os dados são as credenciais da conta de serviço
do Drive, que ficam apenas no servidor.

---

## 💾 Gravação e integridade

Toda alteração vai para o `.xlsx` no Google Drive através de uma conta de serviço.
Como cada gravação reenvia o arquivo inteiro, o app adota algumas salvaguardas:

- **Escrita serializada** — um bloqueio único no processo garante que dois
  salvamentos simultâneos não sobrescrevam um ao outro.
- **Endereçamento por conteúdo** — a linha é localizada pelo NÚMERO e a coluna pelo
  nome do cabeçalho, sempre no arquivo recém-baixado. Posições fixas não são usadas,
  porque uma linha ou coluna inserida na planilha deslocaria todas as seguintes.
- **Recusa em caso de ambiguidade** — se o número não for encontrado, ou houver mais
  de um registro com o mesmo número, nada é alterado e o motivo é exibido.
- **Concorrência otimista nas crianças** — como a edição substitui a lista inteira,
  o salvamento é recusado se outra pessoa alterou a lista nesse intervalo.
- **Gravação antes da tela** — a interface só reflete a alteração depois que o Drive
  confirma, para nunca mostrar como salvo algo que falhou.
- **Cestas extras assíncronas** — o clique responde na hora e a gravação segue em
  segundo plano, agrupando alterações pendentes numa única ida ao Drive.
- **Duplo carimbo de tempo** — cada alteração registra quando foi salva no app e
  quando chegou à planilha; a diferença revela espera de rede.

---

## 🗂️ Estrutura do arquivo `.xlsx`

### Aba 1 — cadastro

Duas linhas de cabeçalho:

| Linha | Conteúdo |
|---|---|
| Linha 1 | Tags: `SEARCH`, `Tier 1` ou `Tier 2` |
| Linha 2 | Nomes das colunas |
| Linha 3+ | Dados dos registros |

**Campos `SEARCH`** (buscáveis): NUMERO, CPF, NOME, RESERVA 1, CPF RESERVA 1,
RESERVA 2, CPF RESERVA 2

**Campos `Tier 1`** (visíveis no card): TIPO, CID 2026, STATUS,
ALERTA PARA A MESA e as colunas mensais de entrega

**Campos `Tier 2`** (não exibidos): endereço, telefone, observações internas e as
três colunas de uso interno do app

As colunas mensais devem ser **datas reais** na linha 2 — o app as exibe como
`Mmm/AA` (ex.: `Mar/26`) e usa o mês corrente para liberar o botão de entrega.

### Colunas de uso interno

| Cabeçalho | Conteúdo |
|---|---|
| `CRIANÇAS` | JSON com a lista de crianças do registro |
| `Salvo no App` | Momento em que o usuário confirmou o salvamento |
| `Gravado na Planilha` | Momento em que a planilha foi efetivamente gravada |

O app localiza essas três pelo nome do cabeçalho e também aceita os nomes
anteriores (`Última Alteração`, `Alterado Por:`, `Motivo da Alteração`), para que
renomear a planilha e publicar o código não precisem acontecer ao mesmo tempo.

### Formato do JSON de crianças

```json
{"criancas":[{"NÚMERO":"15","NOME":"Ana","SEXO":"Feminino","IDADE":8,"OBS":""}]}
```

Campos adicionais presentes no JSON são preservados nas edições feitas pelo app,
mesmo não aparecendo no formulário. Uma célula com conteúdo inválido é tratada como
"sem crianças" na exibição, e o salvamento é recusado para não sobrescrevê-la.

### Aba 2 — cestas extras

Cabeçalho na primeira linha: `NÚMERO`, `NOME`, `CPF`, `TELEFONE`, `MOTIVO`,
`ENTREGUE`.
