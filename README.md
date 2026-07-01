# octano-anp-data

Pipeline automático (e gratuito) que mantém a base de fiscalização da ANP usada pelo app **octano** sempre atualizada.

## Como funciona

Toda segunda-feira às 06:00 (horário de Brasília), uma GitHub Action:

1. Baixa o arquivo oficial `dados-fisc-a-partir-2019.xlsx` do [Painel Dinâmico da Fiscalização do Abastecimento da ANP](https://www.gov.br/anp/pt-br/centrais-de-conteudo/paineis-dinamicos-da-anp/painel-dinamico-da-fiscalizacao-do-abastecimento) — download direto, sem captcha;
2. Valida o formato (header e quantidade mínima de linhas) e converte para CSV no formato que o app consome;
3. Publica `base_anp.csv` e `metadata.json` na release [`latest`](../../releases/tag/latest), substituindo os arquivos anteriores.

Se qualquer passo falhar, o workflow falha e o GitHub envia e-mail automaticamente para o dono do repositório. Opcionalmente, configurando os secrets `MAIL_USERNAME` e `MAIL_APP_PASSWORD` (senha de app do Gmail), um e-mail extra é enviado para luiaffontinelles@gmail.com.

## URLs que o app consome

```
https://github.com/LuisFontinelles/octano-anp-data/releases/download/latest/metadata.json
https://github.com/LuisFontinelles/octano-anp-data/releases/download/latest/base_anp.csv
```

O `metadata.json` tem a data do download (`updatedAt`), o número de linhas e o SHA-256 do CSV — o app só baixa o CSV de 48 MB quando o hash muda.

## Rodar manualmente

Na aba **Actions** → workflow "Atualizar base ANP" → **Run workflow**. Ou localmente:

```bash
pip install -r requirements.txt
python scripts/update_anp.py
```

## Observações

- A ANP atualiza esse arquivo **mensalmente**, com ~2 meses de defasagem; o cron semanal só garante que a novidade é pega logo que sai.
- Se a ANP mudar o formato do arquivo (colunas diferentes), o script falha de propósito em vez de publicar dados corrompidos — o app continua com a última base válida.
