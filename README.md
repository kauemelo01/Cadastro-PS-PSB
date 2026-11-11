# 📋 Cadastro de Visitas — Pastoral Social

Aplicativo web (Streamlit), pensado para uso no celular, para consultar e editar os cadastros da planilha **Visita_2026.xlsx**, que fica no Google Drive.

A pessoa digita o número do registro, confere o resumo do cadastro e edita todos os campos num popup organizado por temas. Cada gravação registra **quem** alterou e **quando**.

---

## Funcionalidades

- **Busca por número do registro** (exata; ignora zeros à esquerda).
- **Link direto de edição:** `https://cadastro-ps-psb.streamlit.app/?id=15` abre automaticamente o popup do cadastro de número 15.
- **Mapas na tela inicial:** links para o Google Maps e para o **OpenStreetMap (offline)**, usando o arquivo HTML do próprio repositório.
- **Card de resumo** com nome, número, CPF, tipo, CID 2026, telefone, endereço, avaliação de cestas e tipo de beneficiário.
- **Botão "✏️ Editar Cadastro"**, sempre disponível em cada card. Não há login nem perfil de administrador.
- **Popup de edição** com todos os campos da planilha, em seções que abrem e fecham com um toque (boa navegação no celular):
  - 🪪 Identificação
  - 👥 Reservas
  - 📍 Endereço e contato
  - 💼 Trabalho e saúde
  - 👨‍👩‍👧 Moradores da casa
  - 🏠 Condições da casa
  - 💰 Contas da casa
  - ⭐ **Avaliação de cestas básicas** (em destaque: PRECISA MUITO / PRECISA / NÃO PRECISA)
  - 📝 Observações e avaliação do entrevistador
- **Segundo popup de confirmação** com a lista "antes → depois" das alterações e **nome obrigatório** de quem está salvando.
- **Auditoria:** o nome e a data/hora (horário de Brasília) são gravados nas colunas **51** e **52** da planilha.
- **Só o que mudou é gravado**, preservando os tipos da planilha (números, datas, textos).
- **Cache persistente:** se o Google Drive estiver indisponível, as alterações ficam pendentes e continuam visíveis no aplicativo.
- **Envio em lote:** quando houver pendências, a tela principal exibe **"☁️ Enviar alterações pendentes"**. Um clique envia todas de uma vez, preservando o autor e a data/hora original de cada confirmação.

### Regras dos campos

| Campo | Comportamento |
|---|---|
| Número do registro | **Bloqueado** — é o único campo que não pode ser editado |
| CPF (titular e reservas) | Se digitado só com 11 dígitos, é formatado como `000.000.000-00` |
| Data de nascimento | Seletor de data (`DD/MM/AAAA`) |
| Eletrodomésticos e facilidades | Múltipla escolha ("Pode marcar mais de um") |
| Campos de quantidade, estado civil, tipo etc. | Listas de opções; valores antigos fora da lista são preservados |
| Sim/Não | Botões `SIM` / `NÃO` |
| Nome | Não pode ficar em branco |

### Abrir um cadastro por link

Acrescente `?id=NUMERO` à URL do aplicativo, usando o **número do cadastro**, e não a posição da linha na planilha:

- [Editar cadastro 15](https://cadastro-ps-psb.streamlit.app/?id=15)
- [Editar cadastro 220](https://cadastro-ps-psb.streamlit.app/?id=220)
- [Editar cadastro 220-A](https://cadastro-ps-psb.streamlit.app/?id=220-A), caso esse número com sufixo exista.

O link preenche a busca e abre o mesmo popup usado pelo botão **"✏️ Editar Cadastro"**, incluindo as alterações pendentes em cache. Números sem sufixo ignoram zeros à esquerda, como na busca normal (`?id=0015` procura o cadastro 15). O popup abre uma única vez por link na sessão; fechar, cancelar, salvar ou fazer outra busca não o reabre automaticamente. Um link com outro `id` pode abrir outro cadastro. Ao abrir o link em uma nova sessão, o popup abre novamente.

Se o cadastro não existir, se o parâmetro estiver vazio ou repetido, ou se houver mais de uma linha com o mesmo número, o app mostra um aviso e não abre automaticamente um cadastro ambíguo. Sem `id` na URL, a aplicação funciona como antes. O salvamento, a confirmação, a auditoria, as regras de conflito, o cache e os estilos permanecem iguais. Os links passam a funcionar após publicar esta versão de `main.py` no aplicativo.

---

## Estrutura esperada da planilha

- **Linha 1:** numeração das colunas.
- **Linha 2:** cabeçalhos.
- **Linhas 3 em diante:** um cadastro por linha.
- **Colunas 1–50:** campos do cadastro (o app localiza as colunas pelo **texto do cabeçalho**, não pela posição).
- **Coluna 51 — `ÚLTIMA EDIÇÃO POR`** e **coluna 52 — `ÚLTIMA EDIÇÃO EM`:** criadas pelo app na primeira gravação (cabeçalhos incluídos), e desocultadas caso estejam ocultas.

> O arquivo no Drive precisa ser um **`.xlsx` de verdade**, e não uma planilha nativa do Google Sheets.

---

## Como a gravação funciona

A cada confirmação, o app registra a alteração no cache local antes de tentar gravar no Drive. Sem pendências anteriores, ele baixa o `.xlsx`, altera o cadastro e envia o arquivo de volta. Após a confirmação do Drive, remove a alteração do cache.

1. A linha é localizada pelo **NUMERO** (nunca pela posição na tela).
2. Os campos alterados são gravados; o nome e o timestamp vão para as colunas 51 e 52.
3. **Conflito de edição:** se, no momento de salvar, outra pessoa tiver alterado o mesmo campo desde que a tela foi carregada (ou o cadastro não for encontrado com segurança), o app **não recusa a gravação**. Ele acrescenta uma **nova linha ao final da planilha** com a versão atual do cadastro mais as alterações do usuário, e o número ganha um sufixo de letra (`220` → `220-A`, `220-B`…). O nome e o timestamp também são gravados nessa linha, e a linha original fica intacta.
   - Alterações de outra pessoa em campos que o usuário **não** editou não geram conflito.
   - Os cadastros duplicados gerados por conflito precisam ser mesclados ou apagados manualmente na planilha.

### Cache e envio das pendências

1. Se houver falha de conexão, autenticação, permissão ou outra falha na gravação, a alteração permanece no cache. O app avisa que ela está aguardando envio.
2. Enquanto houver pendências, novas confirmações também entram na fila, mantendo a ordem das alterações. Elas não são enviadas isoladamente nem antecipadas às anteriores.
3. Quando a conexão e o acesso ao Drive estiverem disponíveis, clique em **"☁️ Enviar alterações pendentes"** na tela principal. O app baixa a planilha uma vez, aplica toda a fila na ordem em que as alterações foram confirmadas e faz **um único upload**.
4. Cada alteração usa o nome e o timestamp registrados na confirmação original, no horário de Brasília. O horário do envio não substitui o horário da edição. Em edições sucessivas do mesmo cadastro, as colunas 51/52 mostram a última edição aplicada, como antes.
5. Se qualquer item não puder ser aplicado ou o envio falhar, a fila inteira permanece no cache. Ela só é removida após a confirmação de sucesso. Um novo clique tenta enviar as pendências novamente.

---

Sem as credenciais do Drive, o app mostra o aviso e mantém as alterações no **cache pendente**. Depois de configurar o acesso, use o botão da tela principal para enviá-las à planilha.

Crie o arquivo `.streamlit/secrets.toml` (local) ou preencha **Settings → Secrets** no Streamlit Community Cloud:

> ⚠️ **Nunca** faça commit do `secrets.toml` nem do JSON da service account. Adicione `.streamlit/secrets.toml` ao `.gitignore`.

---

### Habilitando a gravação no Google Drive

1. No [Google Cloud Console](https://console.cloud.google.com/), crie um projeto e ative a **Google Drive API**.
2. Crie uma **Service Account** e gere uma chave **JSON**.
3. No Drive, **compartilhe o arquivo `.xlsx` com o e-mail da service account** (`...@...iam.gserviceaccount.com`) com permissão de **Editor**.
4. Copie o **ID do arquivo** (trecho da URL entre `/d/` e `/view`) para `GDRIVE_FILE_ID`.
5. Cole o conteúdo do JSON em `GDRIVE_SA_CREDENTIALS`.

### Leitura dos dados

Os dados são lidos de `GITHUB_RAW_URL` e ficam em cache por **5 minutos**. Por isso, uma edição feita por outra pessoa pode demorar até esse tempo para aparecer. Depois de uma gravação, o cache é limpo e as edições da própria sessão continuam visíveis até o arquivo atualizar.

---

## Tecnologias

[Streamlit](https://streamlit.io/) · [pandas](https://pandas.pydata.org/) · [openpyxl](https://openpyxl.readthedocs.io/) · [Google Drive API](https://developers.google.com/drive)