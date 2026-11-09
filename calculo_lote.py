# =============================================================================
# calculo_lote.py — PU / Taxa / Duration em lote, direto na planilha
# -----------------------------------------------------------------------------
# Demanda pontual, independente do add-in e da calculadora local.
#
# Layout da aba (cabeçalho opcional na linha 1):
#     A = Ticker   B = Data   C = Taxa   D = PU   E = Duration
#
# Regra por linha:
#   - PU preenchido (D)          -> calcula a TAXA (C) e a DURATION (E)
#   - senão, Taxa preenchida (C) -> calcula o PU (D)   e a DURATION (E)
#   - nenhum dos dois            -> não calcula nada (E fica em branco)
#   Se as duas estiverem preenchidas, vale o PU (a taxa é recalculada).
#   Falhou em todas as fontes -> célula de saída em branco.
#
# Fontes, nesta ordem: BTG OTC -> FI Analytics (/deb, depois /cr) -> B3 Calculator.
# Título público (NTN-B/NTN-F/código CETIP) pula BTG e FI e vai direto para a B3.
#
# Taxa no formato do Excel: 6,4618% na célula (= 0,064618). A saída também sai assim.
# Duration de Macaulay em ANOS (como as APIs devolvem).
#
# Como rodar:
#   1) Terminal:  python calculo_lote.py
#      Processa a 1ª aba do calculo_lote.xlsx AO LADO do .py, com a planilha fechada
#      (Excel invisível abre, calcula, salva e fecha). Se estiver aberta, usa a aberta.
#      Se o arquivo não existir, cria o modelo vazio com o cabeçalho.
#   2) Botão no Excel (xlwings), na pasta/aba que chamou:
#          Sub CalcularLote()
#              RunPython "import calculo_lote; calculo_lote.main()"
#          End Sub
#
# Credenciais (variáveis de ambiente, iguais às do banco/add-in):
#   token_btg_otc, token_fianalytics_debentures, token_fianalytics_cri_cra,
#   token_calc_b3, proxy_http / proxy_https
# Fallback de dev: um .env (CHAVE=valor) apontado por CALCULO_LOTE_ENV ou ao lado
# do script; lá vale também B3_CALC_TOKEN no lugar de token_calc_b3.
# =============================================================================

import os
import re
import sys
import json
import time
import threading
import urllib.request
import urllib.error
from datetime import date, datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

import xlwings as xw

# ─── configuração ────────────────────────────────────────────────────────────

COL_TICKER   = 1   # A
COL_DATA     = 2   # B
COL_TAXA     = 3   # C
COL_PU       = 4   # D
COL_DURATION = 5   # E

TIMEOUT_SEG = 20   # por chamada
MAX_WORKERS = 12   # linhas calculadas em paralelo (1 = sequencial, para depurar).
                   # Medido: 8 -> ~9 linhas/s, 16 -> ~11 linhas/s; mais que isso só arrisca 429.

ENV_KEY_BTG    = "token_btg_otc"
ENV_KEY_FI_DEB = "token_fianalytics_debentures"
ENV_KEY_FI_CR  = "token_fianalytics_cri_cra"
ENV_TOKEN_B3   = "token_calc_b3"

BTG_BASE    = "https://otcmarkets-api.btgpactualsolutions.com/api/v1/btg-otc-mkts"
BTG_CLIENT  = "calculo-lote"
BTG_TOKEN_SEG = 25 * 60   # o JWT do BTG vale 30 min
FI_BASE     = "https://endpoint.fi-analytics.com.br"
FI_DEB_PATH = "/deb/debenturecalculator"
FI_CR_PATH  = "/cr/cricracalculator"
B3_BASE     = "https://api.calculadorarendafixa.com.br"

PASTA = os.path.dirname(os.path.abspath(__file__))
ARQUIVO = os.path.join(PASTA, "calculo_lote.xlsx")   # planilha de entrada/saída, ao lado do .py
LOG = os.path.join(os.environ.get("TEMP", PASTA), "calculo_lote.log")
EXCEL_EPOCH = date(1899, 12, 30)


# ─── credenciais e proxy ─────────────────────────────────────────────────────

def CarregarEnv():
    """.env de desenvolvimento (CHAVE=valor). No banco valem as variáveis de ambiente."""
    creds = {}
    for caminho in (os.environ.get("CALCULO_LOTE_ENV"), os.path.join(PASTA, ".env")):
        if not caminho:
            continue
        try:
            with open(caminho, encoding="utf-8") as arquivo:
                for linha in arquivo:
                    linha = linha.strip()
                    if "=" in linha and not linha.startswith("#"):
                        chave, valor = linha.split("=", 1)
                        creds.setdefault(chave.strip(), valor.strip())
        except OSError:
            continue
    return creds


CREDENCIAIS = CarregarEnv()


def Cred(nomeEnv, chaveArquivo=None):
    return os.getenv(nomeEnv) or CREDENCIAIS.get(nomeEnv) or (
        CREDENCIAIS.get(chaveArquivo) if chaveArquivo else None)


def ConstruirOpener():
    proxies = {}
    if os.getenv("proxy_http"):
        proxies["http"] = os.getenv("proxy_http")
    if os.getenv("proxy_https"):
        proxies["https"] = os.getenv("proxy_https")
    return urllib.request.build_opener(urllib.request.ProxyHandler(proxies))


_opener = ConstruirOpener()


def Abrir(req):
    return _opener.open(req, timeout=TIMEOUT_SEG)


def Num(valor):
    """Tira ruído de float (6.461800000000001 -> 6.4618)."""
    return float(f"{float(valor):.10g}")


def Positivo(valor):
    try:
        v = float(valor)
    except (TypeError, ValueError):
        return None
    return v if v > 0 else None


# ─── títulos públicos ────────────────────────────────────────────────────────

_RE_NTNB = re.compile(r"^NTN-?B\s*(\d{2})$")
_RE_NTNF = re.compile(r"^NTN-?F\s*(\d{2})$")
_RE_CETIP_TP = re.compile(r"^(760199|950199|100000|210100)\d{8}$")   # NTN-B, NTN-F, LTN, LFT


def NormalizarTicker(tk):
    """'NTNB35' / 'NTN-B 35' -> código CETIP da B3. Demais tickers passam inalterados."""
    m = _RE_NTNB.match(tk)
    if m:
        ano = 2000 + int(m.group(1))
        return f"760199{ano}{'0815' if ano % 2 == 0 else '0515'}"
    m = _RE_NTNF.match(tk)
    if m:
        return f"950199{2000 + int(m.group(1))}0101"
    return tk


def EhTituloPublico(tk):
    return NormalizarTicker(tk) != tk or bool(_RE_CETIP_TP.match(tk))


# ─── BTG OTC ─────────────────────────────────────────────────────────────────

_travaBtg = threading.Lock()
_tokenBtg = None   # (jwt, obtido_em)


def TokenBtg(renovar=False):
    global _tokenBtg
    with _travaBtg:
        if renovar or _tokenBtg is None or time.time() - _tokenBtg[1] > BTG_TOKEN_SEG:
            _tokenBtg = None
            chave = Cred(ENV_KEY_BTG)
            if not chave:
                return None
            try:
                req = urllib.request.Request(
                    f"{BTG_BASE}/authenticate",
                    data=json.dumps({"api_key": chave, "client_id": BTG_CLIENT}).encode(),
                    headers={"Content-Type": "application/json"}, method="POST")
                with Abrir(req) as resp:
                    dados = json.loads(resp.read().decode("utf-8", "replace"))
                token = dados.get("AccessToken") if isinstance(dados, dict) else None
                if token:
                    _tokenBtg = (token, time.time())
            except Exception:
                pass
        return _tokenBtg[0] if _tokenBtg else None


ESPERAS_429 = (1, 2, 4, 8, 15)   # segundos entre novas tentativas quando a API manda 429
# O BTG tem cota POR ENDPOINT e, estourada, segue em 429 sem Retry-After. Depois de duas
# esperas curtas o endpoint é desligado na rodada (disjuntor) e a cascata vai direto p/ a FI.
ESPERAS_429_BTG = (1, 3)
btgBloqueado = set()   # endpoints do BTG desligados nesta rodada por 429


def GetBtg(caminho, params):
    if caminho in btgBloqueado:
        return None
    url = f"{BTG_BASE}/{caminho}?" + "&".join(f"{k}={v}" for k, v in params.items())
    renovou, esperas = False, list(ESPERAS_429_BTG)
    while True:
        if caminho in btgBloqueado:
            return None
        token = TokenBtg()
        if not token:
            return None
        try:
            req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
            with Abrir(req) as resp:
                dados = json.loads(resp.read().decode("utf-8", "replace"))
            return dados[0] if isinstance(dados, list) and dados else dados
        except urllib.error.HTTPError as erro:
            if erro.code == 401 and not renovou:
                renovou = True
                TokenBtg(renovar=True)
                continue
            if erro.code == 429:
                if esperas:
                    time.sleep(esperas.pop(0))
                    continue
                btgBloqueado.add(caminho)
            return None
        except Exception:
            return None


def PrecoBtg(tk, dataIso, taxa):
    if EhTituloPublico(tk):
        return None
    d = GetBtg("pricing/pu", {"ticker": tk, "rate": taxa, "date": dataIso})
    if isinstance(d, dict) and Positivo(d.get("pu")):
        return {"pu": float(d["pu"]), "duration": Positivo(d.get("duration"))}
    return None


def TaxaBtg(tk, dataIso, pu):
    if EhTituloPublico(tk):
        return None
    d = GetBtg("pricing/rate", {"ticker": tk, "pu": pu, "date": dataIso})
    if isinstance(d, dict) and Positivo(d.get("rate")):
        return Positivo(d["rate"]), Positivo(d.get("duration"))   # já vem com a duration
    return None


# ─── FI Analytics ────────────────────────────────────────────────────────────

_CHAVE_FI = {FI_DEB_PATH: ENV_KEY_FI_DEB, FI_CR_PATH: ENV_KEY_FI_CR}
_pathFi = {}   # ticker -> path que respondeu (evita bater no /deb toda vez para CRI/CRA)


def PostFi(path, corpo):
    chave = Cred(_CHAVE_FI[path])
    if not chave:
        return None
    esperas = list(ESPERAS_429)
    while True:
        try:
            req = urllib.request.Request(
                f"{FI_BASE}{path}", data=json.dumps(corpo).encode(),
                headers={"Content-Type": "application/json; charset=utf-8", "x-api-key": chave},
                method="POST")
            with Abrir(req) as resp:
                externo = json.loads(resp.read().decode())
            return json.loads(externo) if isinstance(externo, str) else externo   # double-encoded
        except urllib.error.HTTPError as erro:
            if erro.code == 429 and esperas:
                time.sleep(esperas.pop(0))
                continue
            return None
        except Exception:
            return None


def PostFiAuto(tk, corpo, campo):
    """/deb e depois /cr (ou o que já respondeu antes para este ticker)."""
    if EhTituloPublico(tk):
        return None
    paths = [FI_DEB_PATH, FI_CR_PATH]
    if _pathFi.get(tk) == FI_CR_PATH:
        paths.reverse()
    for path in paths:
        d = PostFi(path, corpo)
        if isinstance(d, dict) and d.get(campo) is not None:
            _pathFi[tk] = path
            return d
    return None


def PrecoFi(tk, dataIso, taxa):
    d = PostFiAuto(tk, {"ticker": tk, "date": dataIso, "rate": taxa}, "m2m")
    if d and Positivo(d.get("m2m")):
        return {"pu": float(d["m2m"]), "duration": Positivo(d.get("maculayDuration"))}
    return None


def TaxaFi(tk, dataIso, pu):
    d = PostFiAuto(tk, {"ticker": tk, "date": dataIso, "pu": pu}, "m2mRate")
    taxa = Positivo(d.get("m2mRate")) if d else None
    return (taxa * 100, Positivo(d.get("maculayDuration"))) if taxa else None


# ─── B3 Calculator ───────────────────────────────────────────────────────────

_travaB3 = threading.Lock()
_tokenB3 = None


def TokenB3(renovar=False):
    global _tokenB3
    with _travaB3:
        if renovar or _tokenB3 is None:
            _tokenB3 = None
            token = Cred(ENV_TOKEN_B3, "B3_CALC_TOKEN")
            if not token:
                return None
            try:
                req = urllib.request.Request(
                    f"{B3_BASE}/login", data=json.dumps({"token": token}).encode(),
                    headers={"Content-Type": "application/json"}, method="POST")
                with Abrir(req) as resp:
                    _tokenB3 = json.loads(resp.read().decode()).get("Authorization")
            except Exception:
                pass
        return _tokenB3


def GetB3(url):
    renovou, esperas = False, list(ESPERAS_429)
    while True:
        token = TokenB3()
        if not token:
            return None
        try:
            req = urllib.request.Request(url, headers={"Authorization": token})
            with Abrir(req) as resp:
                return json.loads(resp.read().decode())
        except urllib.error.HTTPError as erro:
            if erro.code == 401 and not renovou:
                renovou = True
                TokenB3(renovar=True)
                continue
            if erro.code == 429 and esperas:
                time.sleep(esperas.pop(0))
                continue
            return None
        except Exception:
            return None


def PrecoB3(tk, dataIso, taxa):
    d = GetB3(f"{B3_BASE}/calcPU/{NormalizarTicker(tk)}/{dataIso}/{taxa}")
    if isinstance(d, dict) and Positivo(d.get("PU")):
        return {"pu": float(d["PU"]), "duration": Positivo(d.get("duration"))}
    return None


def TaxaB3(tk, dataIso, pu):
    d = GetB3(f"{B3_BASE}/calcYield/{NormalizarTicker(tk)}/{dataIso}/{pu}")
    if isinstance(d, dict) and Positivo(d.get("yield")):
        return Positivo(d["yield"]), Positivo(d.get("duration"))
    return None


# ─── cascata BTG -> FI -> B3 ─────────────────────────────────────────────────

FONTES_PRECO = (("btg", PrecoBtg), ("fi", PrecoFi), ("b3", PrecoB3))
FONTES_TAXA  = (("btg", TaxaBtg), ("fi", TaxaFi), ("b3", TaxaB3))


def Preco(tk, dataIso, taxaPct):
    """({'pu', 'duration'}, fonte) ou (None, None). taxa em % a.a. (6.4618)."""
    taxaPct = Num(taxaPct)
    for nome, funcao in FONTES_PRECO:
        r = funcao(tk, dataIso, taxaPct)
        if r:
            return r, nome
    return None, None


def Taxa(tk, dataIso, pu):
    """(taxa % a.a., duration|None, fonte) ou (None, None, None)."""
    pu = Num(pu)
    for nome, funcao in FONTES_TAXA:
        r = funcao(tk, dataIso, pu)
        if r:
            return r[0], r[1], nome
    return None, None, None


def CalcularLinha(tk, dataIso, taxaPct, pu):
    """Devolve (taxaPct, pu, duration, fonte). Valores None = não calculado."""
    if pu is not None:
        taxaPct, duration, fonte = Taxa(tk, dataIso, pu)
        if taxaPct is None:
            return None, pu, None, None
        if duration is None:   # a fonte da taxa não mandou duration: precifica na taxa achada
            r, _ = Preco(tk, dataIso, taxaPct)
            duration = (r or {}).get("duration")
        return taxaPct, pu, duration, fonte
    r, fonte = Preco(tk, dataIso, taxaPct)
    if not r:
        return taxaPct, None, None, None
    return taxaPct, r["pu"], r["duration"], fonte


# ─── planilha ────────────────────────────────────────────────────────────────

def ParseData(valor):
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return EXCEL_EPOCH + timedelta(days=int(valor))
    if isinstance(valor, str):
        texto = valor.strip()
        for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d/%m/%y"):
            try:
                return datetime.strptime(texto, formato).date()
            except ValueError:
                continue
    return None


def ParseNum(valor):
    """Número da célula ou None (vazio/texto)."""
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace(" ", "")
    pct = texto.endswith("%")
    texto = texto.rstrip("%")
    if texto.count(",") == 1:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        v = float(texto)
    except ValueError:
        return None
    return v / 100 if pct else v


def LocalizarAba(nomeAba=None):
    if isinstance(nomeAba, xw.Sheet):
        return nomeAba
    if nomeAba:
        return xw.books.active.sheets[nomeAba]
    try:
        return xw.Book.caller().sheets.active
    except Exception:
        return xw.books.active.sheets.active


def CriarModelo():
    """Cria o calculo_lote.xlsx vazio, só com o cabeçalho, para preencher."""
    app = xw.App(visible=False, add_book=False)
    try:
        livro = app.books.add()
        aba = livro.sheets[0]
        aba.name = "Lote"
        aba.range("A1").value = ["Ticker", "Data", "Taxa", "PU", "Duration"]
        aba.range("B:B").number_format = "dd/mm/yyyy"   # código em inglês; o Excel PT exibe igual
        aba.range("C:C").number_format = "0.0000%"
        livro.save(ARQUIVO)
        livro.close()
    finally:
        app.quit()


def RodarArquivo():
    """Processa o calculo_lote.xlsx da pasta do script, na 1ª aba, com a planilha FECHADA
    (abre num Excel invisível, calcula, salva e fecha). Se ela estiver aberta, usa a aberta
    e salva."""
    if not os.path.exists(ARQUIVO):
        CriarModelo()
        return f"Criei {ARQUIVO}. Preencha A=Ticker, B=Data, C=Taxa ou D=PU e rode de novo."
    alvo = os.path.normcase(os.path.abspath(ARQUIVO))
    for app in xw.apps:   # já aberta em algum Excel?
        for livro in app.books:
            if os.path.normcase(livro.fullname) == alvo:
                resumo = main(livro.sheets[0])
                livro.save()
                return resumo
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    try:
        livro = app.books.open(ARQUIVO)
        resumo = main(livro.sheets[0])
        livro.save()
        livro.close()
        return resumo
    finally:
        app.quit()


def LerLinhas(aba):
    fundo = aba.range((aba.cells.last_cell.row, COL_TICKER)).end("up").row
    bloco = aba.range((1, COL_TICKER), (fundo, COL_DURATION))
    valores = bloco.value
    formulas = bloco.formula
    if fundo == 1:
        valores, formulas = [valores], [formulas]
    linhas = []
    for i, (v, f) in enumerate(zip(valores, formulas)):
        ticker, dataCel, taxaCel, puCel = v[0], v[1], v[2], v[3]
        if ticker is None or not str(ticker).strip():
            continue
        dataRef = ParseData(dataCel)
        if dataRef is None:
            continue   # cabeçalho ou linha sem data válida
        linhas.append({
            "linha": i + 1,
            "ticker": str(ticker).strip().upper(),
            "dataIso": dataRef.isoformat(),
            "taxa": ParseNum(taxaCel),   # decimal do Excel (0,064618)
            "pu": ParseNum(puCel),
            # C:D como estavam: fórmula de entrada é regravada como fórmula, valor como valor
            "formulas": [fx if isinstance(fx, str) and fx.startswith("=") else vx
                         for fx, vx in zip(f[2:4], v[2:4])] + [None],
        })
    return linhas


def Registrar(mensagem):
    try:
        with open(LOG, "a", encoding="utf-8") as arquivo:
            arquivo.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {mensagem}\n")
    except OSError:
        pass


def main(nomeAba=None):
    inicio = time.time()
    btgBloqueado.clear()   # o xlwings pode reaproveitar o processo entre cliques
    aba = LocalizarAba(nomeAba)
    app = aba.book.app
    linhas = LerLinhas(aba)
    if not linhas:
        return f"'{aba.name}': nenhuma linha com ticker (A) e data (B)."

    # tarefas únicas: (ticker, data, modo, valor) — linha repetida vira uma chamada só
    def Chave(item):
        if item["pu"] is not None:
            return (item["ticker"], item["dataIso"], "pu", Num(item["pu"]))
        if item["taxa"] is not None:
            return (item["ticker"], item["dataIso"], "taxa", Num(item["taxa"] * 100))
        return None

    tarefas = sorted({c for c in map(Chave, linhas) if c})
    resultados = {}

    def Rodar(chave):
        tk, dataIso, modo, valor = chave
        if modo == "pu":
            return CalcularLinha(tk, dataIso, None, valor)
        return CalcularLinha(tk, dataIso, valor, None)

    try:
        app.status_bar = f"Calculando 0/{len(tarefas)}..."
    except Exception:
        pass
    with ThreadPoolExecutor(max_workers=max(1, MAX_WORKERS)) as pool:
        futuros = {pool.submit(Rodar, c): c for c in tarefas}
        for n, futuro in enumerate(as_completed(futuros), 1):
            try:
                resultados[futuros[futuro]] = futuro.result()
            except Exception:
                resultados[futuros[futuro]] = (None, None, None, None)
            if n % 10 == 0 or n == len(tarefas):
                try:
                    app.status_bar = f"Calculando {n}/{len(tarefas)}..."
                except Exception:
                    pass
                print(f"\r{n}/{len(tarefas)}", end="", flush=True)
    print()

    # escrita: um bloco C:E por linha calculada; entrada com fórmula é preservada
    contagem = {"btg": 0, "fi": 0, "b3": 0, "falhou": 0, "vazio": 0}
    falhas = []
    escrita = {}
    for item in linhas:
        chave = Chave(item)
        fTaxa, fPu, _ = item["formulas"]
        if chave is None:
            contagem["vazio"] += 1
            escrita[item["linha"]] = [fTaxa, fPu, None]
            continue
        taxaPct, pu, duration, fonte = resultados[chave]
        if chave[2] == "pu":
            saida = [Num(taxaPct / 100) if taxaPct is not None else None, fPu, duration]
        else:
            saida = [fTaxa, pu, duration]
        escrita[item["linha"]] = saida
        if fonte:
            contagem[fonte] += 1
        else:
            contagem["falhou"] += 1
            falhas.append(f"{item['ticker']} {item['dataIso']} (linha {item['linha']})")

    # agrupa linhas contíguas para escrever em poucos blocos
    numeros = sorted(escrita)
    blocos, atual = [], [numeros[0]]
    for n in numeros[1:]:
        if n == atual[-1] + 1:
            atual.append(n)
        else:
            blocos.append(atual)
            atual = [n]
    blocos.append(atual)
    app.screen_updating = False
    try:
        for bloco in blocos:
            aba.range((bloco[0], COL_TAXA), (bloco[-1], COL_DURATION)).value = [escrita[n] for n in bloco]
    finally:
        app.screen_updating = True

    resumo = (f"{len(linhas)} linhas, {len(tarefas)} cálculos em {time.time() - inicio:.0f}s | "
              f"BTG {contagem['btg']} · FI {contagem['fi']} · B3 {contagem['b3']} · "
              f"falhou {contagem['falhou']} · sem PU/taxa {contagem['vazio']}")
    if btgBloqueado:
        resumo += f" | BTG {', '.join(sorted(btgBloqueado))} com limite estourado (429): foi p/ FI/B3"
    try:
        app.status_bar = resumo
    except Exception:
        pass
    Registrar(f"{aba.book.name}!{aba.name} — {resumo}")
    for falha in falhas:
        Registrar(f"   sem resposta: {falha}")
    return resumo + ("\nSem resposta: " + "; ".join(falhas[:30]) if falhas else "")


if __name__ == "__main__":
    print(RodarArquivo())
