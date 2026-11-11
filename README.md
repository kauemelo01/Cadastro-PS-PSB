# 📋 Cadastro de Visitas — Pastoral Social

Aplicativo web (Streamlit), pensado para uso no celular, para consultar e editar os cadastros da planilha **Visita_2026.xlsx**, que fica no Google Drive.

A pessoa digita o número do registro, confere o resumo do cadastro e edita todos os campos num popup organizado por temas. Cada gravação registra **quem** alterou e **quando**.

---

## Funcionalidades

- **Busca por número do registro** (exata; ignora zeros à esquerda).
- **Card de resumo** com nome, número, CPF, tipo, CID 2026, telefone, endereço, avaliação de cestas e tipo de beneficiário.
- **Botão único: "✏️ Editar Cadastro"**, sempre disponível. Não há login nem perfil de administrador.
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

A cada confirmação, o app baixa o `.xlsx` do Drive, altera o cadastro e envia o arquivo de volta.

1. A linha é localizada pelo **NUMERO** (nunca pela posição na tela).
2. Os campos alterados são gravados; o nome e o timestamp vão para as colunas 51 e 52.
3. **Conflito de edição:** se, no momento de salvar, outra pessoa tiver alterado o mesmo campo desde que a tela foi carregada (ou o cadastro não for encontrado com segurança), o app **não recusa a gravação**. Ele acrescenta uma **nova linha ao final da planilha** com a versão atual do cadastro mais as alterações do usuário, e o número ganha um sufixo de letra (`220` → `220-A`, `220-B`…). O nome e o timestamp também são gravados nessa linha, e a linha original fica intacta.
   - Alterações de outra pessoa em campos que o usuário **não** editou não geram conflito.
   - Os cadastros duplicados gerados por conflito precisam ser mesclados ou apagados manualmente na planilha.

---

Sem as credenciais do Drive, o app funciona em **modo somente sessão**: mostra um aviso e as edições valem apenas enquanto a página estiver aberta (nada é gravado na planilha).

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