# 🎁 Apadrinhamento — Pastoral Social Jardim Colombo

App web (Streamlit), pensado para celular, para cadastrar padrinhos das crianças atendidas pela Pastoral Social. Os dados são lidos e gravados diretamente na planilha `cadastro.xlsx` no Google Drive.

## Funcionalidades

- **Busca por Nº do registro / carteirinha**: mostra todas as crianças vinculadas àquela família.
- **Cartão por criança** com:
  - nome da criança em destaque, sexo e idade;
  - nome do beneficiário (em menor destaque);
  - observações da criança, quando houver;
  - ícone conforme sexo e faixa etária:

    | Faixa         | Menino   | Menina   |
    |---------------|----------|----------|
    | 0 a 3 anos    | 👶 🧢    | 👶 🎀    |
    | 4 a 8 anos    | 👦       | 👧       |
    | 9 anos ou mais| 🙋‍♂️      | 🙋‍♀️      |

- **Cadastrar / Editar Padrinho**: botão em cada criança abre um popup com os campos **Padrinho** e **Telefone** (ambos obrigatórios; telefone com DDD, mín. 10 dígitos). Se a criança já tem padrinho, ele é exibido no cartão e o botão vira "Editar Padrinho".
- **Estatísticas** no rodapé: crianças apadrinhadas, % concluído, total e quantas faltam — no geral e por faixa etária (0 a 3, 4 a 8, 9 ou mais).

## Estrutura da planilha

O app usa a **primeira aba** de `cadastro.xlsx`, com cabeçalho na linha 1:

| Coluna | Cabeçalho                  | Observação                                   |
|--------|----------------------------|----------------------------------------------|
| A      | `Nº Família / Carteirinha` | obrigatório                                  |
| B      | `Beneficiário (Cesta)`     |                                              |
| C      | `Nome da criança`          | obrigatório                                  |
| D      | `Sexo`                     | `Masculino` / `Feminino`                     |
| E      | `Idade`                    | ex.: `6 anos`, `3 meses`                     |
| F      | `Observações`              |                                              |
| —      | `Padrinho`                 | criada automaticamente no 1º salvamento      |
| —      | `Telefone Padrinho`        | criada automaticamente no 1º salvamento      |

Notas:
- A ordem das colunas não importa; o app localiza cada uma pelo nome do cabeçalho (ignorando acentos e maiúsculas).
- Linhas sem número ou sem nome da criança são ignoradas.
- Ao salvar, a linha é localizada pelo **Nº + nome da criança**, então reordenar a planilha não faz o padrinho ser gravado na criança errada.

## Configuração

### 1. Dependências

`requirements.txt`:

```
streamlit
pandas
openpyxl
requests
google-auth
google-api-python-client
```

### 2. Google Drive

1. Crie uma *service account* no Google Cloud com a **Google Drive API** ativada e gere uma chave JSON.
2. Compartilhe o arquivo `cadastro.xlsx` no Drive com o e-mail da service account, com permissão de **Editor**.
3. Gere um link de download direto do arquivo (formato `https://drive.google.com/uc?export=download&id=<ID_DO_ARQUIVO>`).

### 3. Secrets

No Streamlit Community Cloud (*App → Settings → Secrets*) ou em `.streamlit/secrets.toml` para rodar localmente:

```toml
GITHUB_RAW_URL = "https://drive.google.com/uc?export=download&id=<ID_DO_ARQUIVO>"
GDRIVE_FILE_ID = "<ID_DO_ARQUIVO>"
GDRIVE_SA_CREDENTIALS = '''
{ ...conteúdo do JSON da service account... }
'''
```

| Secret                  | Uso                                              |
|-------------------------|--------------------------------------------------|
| `GITHUB_RAW_URL`        | link de leitura da planilha                      |
| `GDRIVE_FILE_ID`        | ID do arquivo no Drive, usado para gravar        |
| `GDRIVE_SA_CREDENTIALS` | credenciais da service account (JSON)            |

Sem `GDRIVE_FILE_ID` e `GDRIVE_SA_CREDENTIALS`, o app funciona só para consulta e exibe um aviso de que o salvamento está desativado.

> ⚠️ Nunca faça commit do `secrets.toml` nem do JSON da service account. Adicione `.streamlit/secrets.toml` ao `.gitignore`.

## Rodando localmente

```bash
pip install -r requirements.txt
streamlit run main.py
```

## Deploy

1. Suba `main.py` e `requirements.txt` para este repositório.
2. Em [share.streamlit.io](https://share.streamlit.io), crie um app apontando para o repositório, branch `main` e arquivo `main.py`.
3. Configure os Secrets conforme acima.

## Observações técnicas

- Os dados ficam em cache por 5 minutos; após cada salvamento o cache é limpo e a alteração aparece na hora.
- Cada gravação baixa a versão mais recente do arquivo no Drive, altera apenas as células de padrinho/telefone da criança e reenvia o arquivo.
