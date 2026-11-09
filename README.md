# Cálculo em lote — PU / Taxa / Duration

Script pontual (`calculo_lote.py`), independente do add-in e da calculadora local. Lê a aba,
calcula tudo em paralelo pelas APIs e grava de uma vez.

| A | B | C | D | E |
|---|---|---|---|---|
| Ticker | Data | Taxa | PU | Duration |

- **PU preenchido** → calcula Taxa (C) e Duration (E).
- **Só taxa preenchida** → calcula PU (D) e Duration (E).
- **Nenhum dos dois** → nada (E fica em branco). **Os dois** → vale o PU (taxa é recalculada).
- Falhou em todas as fontes → célula em branco; a lista sai no resumo e no `%TEMP%\calculo_lote.log`.
- Taxa no formato do Excel (6,5% = 0,065), na entrada e na saída. Duration de Macaulay em anos.
- Fontes: **BTG → FI (/deb, /cr) → B3**. Título público (NTN-B/F, código CETIP) vai direto à B3.
- Cabeçalho opcional: linha sem data válida em B é ignorada. Fórmula em C/D de entrada é preservada.
- Se um endpoint do BTG estourar a cota (429), ele é desligado naquela rodada e o resumo avisa.

## Rodar

Preencha a 1ª aba do **`calculo_lote.xlsx`** (ao lado do `.py`) e rode:

```
python D:\ItauBBA\calculo-lote\calculo_lote.py
```

A planilha pode estar **fechada**: um Excel invisível abre, calcula, salva e fecha. Se estiver aberta,
o script usa a aberta e salva. Sem o arquivo, a 1ª execução cria o modelo com o cabeçalho.

Velocidade medida: ~10 linhas/s (12 chamadas em paralelo) → 2.000 linhas em ~3–4 min.

Também dá num botão (xlwings), na pasta que chamou: `RunPython "import calculo_lote; calculo_lote.main()"`.

## Credenciais

Variáveis de ambiente (as mesmas do add-in): `token_btg_otc`, `token_fianalytics_debentures`,
`token_fianalytics_cri_cra`, `token_calc_b3`, `proxy_http`/`proxy_https`.
Fallback de dev: `.env` apontado por `CALCULO_LOTE_ENV` (aceita `B3_CALC_TOKEN`).

Ajustes no topo do `.py`: `COL_*`, `MAX_WORKERS` (12), `ARQUIVO`, `TIMEOUT_SEG` (20).
