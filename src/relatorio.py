# -*- coding: utf-8 -*-
"""
Relatorio de Vendas do GO e gestao do token (sem interface).

Mesma logica do go_relatorio.py do vault (Automacoes/Agente-GO), adaptada para
ser chamada pela janela: em vez de SystemExit, levanta excecoes com mensagem
pronta para mostrar ao usuario.

O relatorio inteiro sai de UM GET em /report (descoberto por HAR em 25/08/2026):
    GET /report?q=report&report_id=44&access_token=<token>
        &filter_report[0][field]=dataNegociacao&filter_report[0][type]=between
        &filter_report[0][from]=AAAA-MM-DD 00:00:00&filter_report[0][to]=AAAA-MM-DD 23:59:59
        &filter_report[1][field]=unidade&filter_report[1][type]=eq&filter_report[1][value]=<cod>

Token: o GO nao implementa renovacao (refresh_token responde
unsupported_grant_type) e o login exige reCAPTCHA. Por isso o token e sempre
colado pelo usuario, a partir do GO aberto no navegador, e vale 12 h.
"""
import os
import time
from datetime import date, datetime, timedelta
from urllib.parse import quote

import requests

try:
    import keyring
except Exception:  # sem cofre: o token vale so enquanto a janela estiver aberta
    keyring = None

BASE = "https://solutivd.gestao.plus"
SERVICE = "agente-go"
REPORT_ID_VENDAS = "44"  # "CSV - Vendas por unidade"
CAMPO_DATA = "dataNegociacao"
VALIDADE_SEG = 43200  # 12 h (expires_in do /oauth)
TIMEOUT_RELATORIO = 300  # periodos longos demoram a gerar no servidor
DOWNLOAD_DIR = os.path.join(os.environ.get("USERPROFILE", "."), "Downloads")


class TokenRecusado(Exception):
    """O GO nao aceitou o access_token (vencido, revogado ou copiado errado)."""


class ErroRelatorio(Exception):
    """Falha que nao e de token (resposta inesperada, rede, disco)."""


# ---------------------------------------------------------------------------
# Token no cofre do Windows
# ---------------------------------------------------------------------------
def token_salvo():
    """(token, expira_em_epoch) do cofre; (None, None) se nao houver."""
    if not keyring:
        return None, None
    try:
        tk = keyring.get_password(SERVICE, "access_token")
        exp = keyring.get_password(SERVICE, "token_expira_em")
        return tk, (float(exp) if exp else None)
    except Exception:
        return None, None


def salvar_token(token):
    """Grava o token e a hora prevista de vencimento. Devolve o epoch de vencimento."""
    expira = time.time() + VALIDADE_SEG
    if keyring:
        keyring.set_password(SERVICE, "access_token", token)
        keyring.set_password(SERVICE, "token_expira_em", str(expira))
    return expira


def testar_token(token, unidade_teste="52"):
    """Faz uma chamada leve para saber se o GO aceita o token.

    Devolve True/False. Levanta ErroRelatorio se nem deu para perguntar (rede).
    """
    try:
        r = requests.get(f"{BASE}/unidade/{unidade_teste}",
                         params={"access_token": token},
                         headers={"Accept": "application/json"}, timeout=20)
    except requests.RequestException as e:
        raise ErroRelatorio(f"Sem conexao com o GO: {e}")
    if r.status_code in (401, 403):
        return False
    return r.ok and "json" in (r.headers.get("Content-Type") or "").lower()


# ---------------------------------------------------------------------------
# Datas e unidades
# ---------------------------------------------------------------------------
def data_br_para_iso(ddmmaaaa):
    """DD/MM/AAAA -> AAAA-MM-DD. Levanta ValueError com mensagem clara."""
    try:
        return datetime.strptime(ddmmaaaa.strip(), "%d/%m/%Y").strftime("%Y-%m-%d")
    except ValueError:
        raise ValueError(f"Data invalida: {ddmmaaaa!r}. Use DD/MM/AAAA.")


def periodo_atalho(nome, hoje=None):
    """Datas (DD/MM/AAAA, DD/MM/AAAA) dos atalhos da tela."""
    hoje = hoje or date.today()
    if nome == "hoje":
        ini = fim = hoje
    elif nome == "ontem":
        ini = fim = hoje - timedelta(days=1)
    elif nome == "este_mes":
        ini, fim = hoje.replace(day=1), hoje
    elif nome == "mes_passado":
        fim = hoje.replace(day=1) - timedelta(days=1)
        ini = fim.replace(day=1)
    elif nome == "ultimos_7":
        ini, fim = hoje - timedelta(days=6), hoje
    else:
        raise ValueError(nome)
    return ini.strftime("%d/%m/%Y"), fim.strftime("%d/%m/%Y")


def nome_unidade(token, codigo):
    """Nome da unidade pelo codigo (ex.: 52 -> 'E-commerce - Soluti').

    None se nao achar. Levanta TokenRecusado se o GO recusar o token.
    """
    try:
        r = requests.get(f"{BASE}/unidade/{codigo}", params={"access_token": token},
                         headers={"Accept": "application/json"}, timeout=20)
    except requests.RequestException:
        return None
    if r.status_code in (401, 403):
        raise TokenRecusado()
    if not r.ok or "json" not in (r.headers.get("Content-Type") or "").lower():
        return None
    d = r.json() or {}
    return d.get("descricaoCompleta") or d.get("descricao")


def listar_unidades(token):
    """Lista [(codigo, nome)] das unidades visiveis para o usuario.

    Usa o mesmo padrao de listagem das outras colecoes do GO (_embedded.<nome>).
    E best-effort: se o formato for outro, devolve [] e a tela segue aceitando
    o codigo digitado a mao.
    """
    try:
        r = requests.get(f"{BASE}/unidade",
                         params={"pageSize": 500, "access_token": token},
                         headers={"Accept": "application/json"}, timeout=30)
    except requests.RequestException:
        return []
    if r.status_code in (401, 403):
        raise TokenRecusado()
    if not r.ok or "json" not in (r.headers.get("Content-Type") or "").lower():
        return []
    itens = (r.json().get("_embedded") or {}).get("unidade") or []
    out = []
    for u in itens:
        cod = u.get("id") or u.get("codigo")
        nome = u.get("descricaoCompleta") or u.get("descricao")
        if cod not in (None, "") and nome:
            out.append((str(cod), str(nome)))
    out.sort(key=lambda x: x[1].lower())
    return out


# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------
def monta_url(token, data_ini_iso, data_fim_iso, unidade="", report_id=REPORT_ID_VENDAS):
    """Query montada na mao com quote(safe=''): o portal manda %20 no espaco,
    e o params= do requests mandaria '+'."""
    params = [("q", "report"), ("report_id", str(report_id)), ("access_token", token),
              ("filter_report[0][field]", CAMPO_DATA),
              ("filter_report[0][type]", "between"),
              ("filter_report[0][from]", f"{data_ini_iso} 00:00:00"),
              ("filter_report[0][to]", f"{data_fim_iso} 23:59:59")]
    if unidade:
        params += [("filter_report[1][field]", "unidade"),
                   ("filter_report[1][type]", "eq"),
                   ("filter_report[1][value]", str(unidade))]
    qs = "&".join(f"{quote(k, safe='')}={quote(str(v), safe='')}" for k, v in params)
    return f"{BASE}/report?{qs}"


def nome_arquivo_padrao(data_ini_iso, data_fim_iso, unidade=""):
    return (f"vendas_{data_ini_iso}_a_{data_fim_iso}"
            + (f"_un{unidade}" if unidade else "_todas") + ".xlsx")


def baixar_relatorio(token, data_ini_br, data_fim_br, unidade="", destino=None,
                     report_id=REPORT_ID_VENDAS, progresso=None):
    """Baixa o XLSX e devolve (caminho, bytes, segundos).

    So grava se a resposta for mesmo planilha: com token vencido o GO pode
    responder 200 com JSON/HTML, e gravar isso como .xlsx geraria arquivo
    corrompido. Grava primeiro num .parcial e so renomeia no fim, para nao
    deixar meio arquivo com o nome final se a conexao cair.
    `progresso(bytes_baixados)` e chamado durante o download.
    """
    ini = data_br_para_iso(data_ini_br)
    fim = data_br_para_iso(data_fim_br)
    if ini > fim:
        raise ValueError("A data inicial e posterior a data final.")
    unidade = "".join(ch for ch in str(unidade or "") if ch.isdigit())
    if not destino:
        destino = os.path.join(DOWNLOAD_DIR, nome_arquivo_padrao(ini, fim, unidade))

    t0 = time.time()
    url = monta_url(token, ini, fim, unidade, report_id)
    try:
        with requests.get(url, stream=True, timeout=TIMEOUT_RELATORIO) as r:
            ctype = (r.headers.get("Content-Type") or "").lower()
            if "spreadsheetml" not in ctype and "octet-stream" not in ctype:
                trecho = r.text[:300].replace("\n", " ")
                if r.status_code in (401, 403) or "token" in trecho.lower():
                    raise TokenRecusado(trecho)
                raise ErroRelatorio(
                    f"O GO nao devolveu uma planilha (HTTP {r.status_code}, {ctype or 'sem tipo'}).\n"
                    f"Resposta: {trecho}")
            r.raise_for_status()
            parcial = destino + ".parcial"
            total = 0
            with open(parcial, "wb") as f:
                for chunk in r.iter_content(chunk_size=64 * 1024):
                    if chunk:
                        f.write(chunk)
                        total += len(chunk)
                        if progresso:
                            progresso(total)
    except requests.RequestException as e:
        raise ErroRelatorio(f"Falha de conexao com o GO: {e}")
    os.replace(parcial, destino)
    return destino, total, time.time() - t0
