# Dashboard de Elegibilidades

Lê uma planilha do Google Sheets (uma aba por dia do mês, sempre no mesmo
layout) e mostra um painel interativo em Streamlit com 3 visões: Visão
Geral, Por Dia e Por Local.

## 1. Criar a Service Account no Google Cloud

1. Acesse https://console.cloud.google.com/ e crie (ou use) um projeto.
2. Ative as APIs **Google Sheets API** e **Google Drive API**.
3. Vá em *IAM & Admin -> Service Accounts -> Create Service Account*.
4. Após criar, entre nela, aba *Keys -> Add Key -> JSON* — isso baixa um
   arquivo `.json` com as credenciais.
5. **Compartilhe sua planilha do Google Sheets** com o e-mail dessa conta
   de serviço (algo como `nome@projeto.iam.gserviceaccount.com`), como se
   fosse compartilhar com uma pessoa — permissão de leitor já basta.

## 2. Configurar os secrets

Copie `.streamlit/secrets.toml.example` para `.streamlit/secrets.toml` e
preencha com os dados do JSON baixado no passo anterior, mais o ID (ou a
URL) da sua planilha. O ID é o trecho da URL entre `/d/` e `/edit`:

```
https://docs.google.com/spreadsheets/d/ESTE_TRECHO_AQUI/edit
```

**Nunca** suba esse arquivo para o GitHub — o `.gitignore` já está
configurado para ignorá-lo.

## 3. Rodar localmente

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## 4. Subir no GitHub

```bash
git init
git add .
git commit -m "Dashboard de elegibilidades"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git
git push -u origin main
```

(o `secrets.toml` real não vai junto, por causa do `.gitignore` — isso é
o esperado e o correto)

## 5. Deploy no Streamlit Community Cloud

1. Acesse https://share.streamlit.io/ e conecte sua conta do GitHub.
2. Clique em *New app*, escolha o repositório e o arquivo `app.py`.
3. Em *Advanced settings -> Secrets*, cole o conteúdo do seu
   `secrets.toml` (o mesmo conteúdo do passo 2).
4. Deploy. Pronto — o app fica público em uma URL `*.streamlit.app`.

## 6. Ajustando ao layout real da sua planilha

O painel lê os números de células fixas dentro de cada aba (definidas em
`CELL_MAP`, no topo do `app.py`), porque cada aba representa um dia e
segue sempre o mesmo modelo. Se algum valor aparecer errado ou vazio no
dashboard, é sinal de que a célula real na sua planilha é diferente da
que está mapeada — basta abrir `app.py` e corrigir o endereço da célula
correspondente.

Por padrão, o app tenta reconhecer o nome de cada aba como um dia do mês
selecionado na barra lateral (aceita formatos como `1`, `01`, `dia 1`,
`01/08`, `01/08/2026`). Se suas abas tiverem outro padrão de nome, ajuste
a função `parse_day_from_title` em `app.py`.
