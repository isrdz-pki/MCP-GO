# -*- coding: utf-8 -*-
"""
Consulta de Pedidos do GO - App (interface neutra com customtkinter)
====================================================================

Consulta pedidos do ERP GO pela API (sem navegador). Tres usos:

  1. Campo do TOKEN: cole o access_token do GO (vale ~12h). O app guarda o ultimo
     no cofre do Windows so para pre-preencher. Tem uma "colinha" com o comando de
     pegar o token no console do Edge, e um botao que copia esse comando.
  2. Consultar UM pedido: mostra TODOS os campos que a API devolve (tabela
     horizontal: cabecalhos em cima, valores embaixo).
  3. Consultar em LOTE: escolha uma planilha (.xlsx) com uma coluna de pedidos;
     gera uma planilha nova com todos os campos. Pedido que a API nao achar vira
     "pedido incorreto ou nao encontrado".

Rodar:   py go_consulta_gui.py
Gerar .exe:
    pip install pyinstaller
    pyinstaller --onefile --windowed --collect-all customtkinter go_consulta_gui.py
"""
import json
import os
import queue
import re
import threading
from datetime import datetime

import requests

try:
    import keyring
except Exception:
    keyring = None

import openpyxl

import customtkinter as ctk
from tkinter import filedialog, messagebox

# ---------------------------------------------------------------------------
# Constantes da API / dados
# ---------------------------------------------------------------------------
BASE = "https://solutivd.gestao.plus"
SERVICE = "agente-go"
TIMEOUT = 30
DOWNLOAD_DIR = os.path.join(os.environ.get("USERPROFILE", "."), "Downloads")
STATUS = {1: "Liberado"}  # observado no grid; os demais ficam como codigo
CMD_TOKEN = "copy(localStorage.getItem('access_token'))"
MSG_NAO_ENCONTRADO = "pedido incorreto ou nao encontrado"
MSG_INVALIDO = "valor nao parece um numero de pedido (10 digitos iniciando com 7)"
MSG_ENCONTRADO = "encontrado"

# paleta NEUTRA (grafite/cinza) - sem azul, sem branco
COR_BG = "#1F1F1F"      # fundo
COR_CARD = "#2A2A2A"    # cartoes
COR_CARD2 = "#383838"   # realces / cabecalho da tabela
COR_CAMPO = "#171717"   # fundo de entradas
COR_TXT = "#DDDDDD"     # texto (cinza claro, nao branco)
COR_SUB = "#9E9E9E"     # texto secundario
COR_BTN = "#4D4D4D"     # botao principal
COR_BTN_H = "#5E5E5E"
COR_BTN2 = "#3A3A3A"    # botao secundario
COR_BTN2_H = "#4A4A4A"
COR_ERRO = "#D98A8A"    # vermelho suave so para erro

RAIO_CARD = 8
RAIO_PEQ = 5
RAIO_CELULA = 3

LABELS = {
    "codigo": "Numero do pedido", "id": "Id interno", "status": "Status",
    "valorTotal": "Valor total", "valorDesconto": "Valor desconto",
    "valorFrete": "Valor frete", "valorRetido": "Valor retido",
    "percentualDesconto": "% desconto", "dataNegociacao": "Data da compra",
    "dataFaturamento": "Data de faturamento", "dataEntrega": "Data de entrega",
    "situacaoNf": "Situacao NF", "situacaoFinanceiro": "Situacao financeiro",
    "situacaoEntrega": "Situacao entrega", "origem": "Origem",
    "observacao": "Observacao", "voucher": "Voucher",
    "codigoExterno": "Codigo externo", "movimentacaoItensDescricao": "Itens",
    "recorrenciaPeriodicidade": "Recorrencia - periodicidade",
    "recorrenciaStatus": "Recorrencia - status",
    "codigoDeRastreio": "Codigo de rastreio",
    "parceiro.nome": "Cliente", "parceiro.razaoSocial": "Cliente (razao social)",
    "parceiro.cpfCnpj": "CPF/CNPJ", "unidade.descricao": "Unidade",
    "unidade.codigo": "Unidade (codigo)",
    "unidadeDeEntrega.descricao": "Unidade de entrega",
    "vendedor.nome": "Vendedor", "tipoDeNegociacao.descricao": "Forma de pagamento",
    "tipoMovimentacao.descricao": "Tipo de movimentacao",
    "empresa.razaoSocial": "Empresa", "centroDeResultado.descricao": "Centro de resultado",
}
ORDEM = [
    "Numero do pedido", "Cliente", "CPF/CNPJ", "Itens", "Valor total",
    "Valor desconto", "% desconto", "Forma de pagamento", "Data da compra",
    "Data de faturamento", "Unidade", "Origem", "Vendedor", "Status",
    "Situacao NF", "Situacao financeiro", "Situacao entrega", "Observacao",
    "Tipo de movimentacao", "Id interno",
    "Qtd anexos", "Anexo 1", "Anexo 2", "Anexo 3", "Anexo 4", "Anexo 5",
]


class TokenInvalido(Exception):
    """O GO recusou o token (vencido/invalido)."""


# ---------------------------------------------------------------------------
# Nucleo da consulta (sem GUI) - testado
# ---------------------------------------------------------------------------
def fmt_data(v):
    if isinstance(v, dict):
        v = v.get("date")
    if not v:
        return ""
    try:
        return datetime.strptime(str(v)[:19], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return str(v)


def so_digitos(v):
    return re.sub(r"\D", "", str(v or ""))


def fmt_doc(v):
    """Formata CPF (xxx.xxx.xxx-xx) ou CNPJ (xx.xxx.xxx/xxxx-xx)."""
    d = so_digitos(v)
    if len(d) == 11:
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    return str(v or "")


def codigo_para_id(codigo):
    c = so_digitos(codigo)
    if len(c) < 5:
        return None
    return c[1:-3]


def achatar(d):
    out = {}

    def put(chave, val):
        out[LABELS.get(chave, chave)] = val

    for k, v in d.items():
        if k in ("_links", "_embedded") or v in (None, ""):
            continue
        if isinstance(v, dict):
            put(k, fmt_data(v) if "date" in v else json.dumps(v, ensure_ascii=False))
        elif isinstance(v, list):
            if v:
                put(k, json.dumps(v, ensure_ascii=False))
        elif k == "status":
            put(k, STATUS.get(v, f"codigo {v}"))
        else:
            put(k, v)

    for ek, ev in (d.get("_embedded") or {}).items():
        if isinstance(ev, dict):
            for fk in ("descricao", "descricaoCompleta", "nome", "razaoSocial",
                       "nomeFantasia", "cpfCnpj", "identificador", "codigo"):
                val = ev.get(fk)
                if val not in (None, ""):
                    put(f"{ek}.{fk}", val)
    return out


def achatar_parceiro(p):
    """Dados cadastrais do cliente como lista [(rotulo, valor)], na ordem de
    leitura, pulando vazios. NAO expoe o campo 'token' (hash/segredo)."""
    out = []

    def add(rot, val):
        if val not in (None, "", []):
            out.append((rot, val))

    tipos = [rot for chave, rot in (
        ("tipoCliente", "Cliente"), ("tipoFornecedor", "Fornecedor"),
        ("tipoIndicador", "Indicador"), ("tipoTransportadora", "Transportadora"),
        ("tipoFuncionario", "Funcionario")) if p.get(chave)]

    add("Nome", p.get("nome"))
    rs = p.get("razaoSocial")
    if rs and rs != p.get("nome"):
        add("Razao social", rs)
    add("CPF/CNPJ", fmt_doc(p.get("cpfCnpj")))
    add("Tipo de pessoa", {"PF": "Pessoa fisica", "PJ": "Pessoa juridica"}.get(
        p.get("tipoPessoa"), p.get("tipoPessoa")))
    add("Tipo de cadastro", ", ".join(tipos))
    add("Data de nascimento", fmt_data(p.get("dataDeNascimento")))
    add("RG", p.get("rg"))
    add("Inscricao estadual", p.get("inscricaoEstadual"))
    add("Inscricao municipal", p.get("inscricaoMunicipal"))
    add("Codigo", p.get("codigo"))
    add("E-mail", p.get("email"))
    add("Telefone", p.get("telefone"))
    add("Celular", p.get("celular"))
    add("Telefone comercial", p.get("telefoneComercial"))

    endereco = (p.get("endereco") or "").strip()
    num = p.get("numero")
    linha = ", ".join(x for x in (endereco, str(num) if num not in (None, "") else "") if x)
    comp = (p.get("complemento") or "").strip()
    if comp:
        linha = f"{linha} - {comp}" if linha else comp
    add("Endereco", linha)
    add("Bairro", p.get("bairro"))
    add("Cidade", p.get("cidade"))
    add("Estado", p.get("estado"))
    add("CEP", p.get("cep"))

    ind = (p.get("_embedded") or {}).get("parceiroIndicador")
    if isinstance(ind, dict):
        add("Indicado por", ind.get("nome"))
    add("Data de cadastro", fmt_data(p.get("dataDeCadastro")))
    add("Data de alteracao", fmt_data(p.get("dataDeAlteracao")))
    return out


def _listar_anexos(token, mov_id):
    """GET /movimentacao-financeiro-anexo filtrado pela movimentacao.

    Descoberto por HAR (2026-10-07): a aba 'Anexos' da venda chama esse endpoint;
    cada item tem 'descricao' e, em _embedded.file, o campo 'uuid' que JA E o 'q'
    assinado (HMAC) do download. Levanta TokenInvalido em 401/403; devolve a lista
    crua de anexos (pode ser [])."""
    url = (f"{BASE}/movimentacao-financeiro-anexo"
           f"?filter[0][field]=movimentacao&filter[0][type]=eq&filter[0][value]={mov_id}"
           f"&pageSize=50&access_token={token}")
    r = requests.get(url, headers={"Accept": "application/json"}, timeout=TIMEOUT)
    if r.status_code in (401, 403):
        raise TokenInvalido()
    if r.status_code != 200 or "json" not in (r.headers.get("Content-Type") or "").lower():
        return []
    return (r.json().get("_embedded") or {}).get("movimentacao_financeiro_anexo") or []


def _anexo_para_link(item):
    """De um item de anexo, devolve {descricao, link} ou None.
    Link auto-suficiente (nao precisa de token): /file?forceDownload=true&q=<uuid>."""
    f = (item.get("_embedded") or {}).get("file") or {}
    uuid = f.get("uuid")
    if not uuid:
        return None
    return {"descricao": item.get("descricao") or f.get("name") or "anexo",
            "link": f"{BASE}/file?forceDownload=true&q={uuid}"}


def anexos_do_pedido(token, mov_id):
    """Lista anexos (descricao + link). [] em qualquer erro (nao derruba consulta)."""
    try:
        itens = _listar_anexos(token, mov_id)
    except Exception:
        return []
    return [a for a in (_anexo_para_link(i) for i in itens) if a]


def baixar_anexos(token, codigo, pasta):
    """Baixa os arquivos dos anexos (NF/DANFE, XML) do pedido para `pasta`.
    Levanta TokenInvalido se o token for recusado. Devolve [(descricao, caminho)]."""
    ident = codigo_para_id(codigo)
    if not ident:
        return []
    itens = _listar_anexos(token, ident)  # pode levantar TokenInvalido
    salvos = []
    for it in itens:
        a = _anexo_para_link(it)
        if not a:
            continue
        try:
            d = requests.get(a["link"], timeout=120)
        except requests.RequestException:
            continue
        if not (d.ok and d.content):
            continue
        nome = re.sub(r'[\\/:*?"<>|]+', "_", a["descricao"]).strip()[:120]
        if not re.search(r"\.\w{2,4}$", nome):
            ct = d.headers.get("Content-Type", "")
            nome += ".pdf" if "pdf" in ct else (".xml" if "xml" in ct else ".bin")
        dest = os.path.join(pasta, f"{so_digitos(codigo)}_{nome}")
        try:
            with open(dest, "wb") as fh:
                fh.write(d.content)
            salvos.append((a["descricao"], dest))
        except Exception:
            pass
    return salvos


def buscar_parceiro(token, documento):
    """Acha o cliente (parceiro) pelo CPF/CNPJ (campo 'cpfCnpj'). Levanta
    TokenInvalido; None se nao achar."""
    doc = so_digitos(documento)
    url = (f"{BASE}/parceiro?filter[0][field]=cpfCnpj&filter[0][type]=eq"
           f"&filter[0][value]={doc}&pageSize=5&access_token={token}")
    r = requests.get(url, headers={"Accept": "application/json"}, timeout=TIMEOUT)
    if r.status_code in (401, 403):
        raise TokenInvalido()
    if r.status_code != 200:
        return None
    lst = (r.json().get("_embedded") or {}).get("parceiro") or []
    return lst[0] if lst else None


def pedidos_do_parceiro(token, parceiro_id, limite=500):
    """Lista as movimentacoes (pedidos) de um parceiro, mais recentes primeiro."""
    url = (f"{BASE}/movimentacao?filter[0][field]=parceiro&filter[0][type]=eq"
           f"&filter[0][value]={parceiro_id}"
           f"&order-by[0][type]=field&order-by[0][field]=dataNegociacao"
           f"&order-by[0][direction]=desc&pageSize={limite}&access_token={token}")
    r = requests.get(url, headers={"Accept": "application/json"}, timeout=90)
    if r.status_code in (401, 403):
        raise TokenInvalido()
    if r.status_code != 200:
        return []
    return (r.json().get("_embedded") or {}).get("movimentacao") or []


def salvar_pedidos_xlsx(pedidos, destino):
    """Grava uma lista de movimentacoes (achatadas) num .xlsx."""
    linhas, cols = [], []
    for ped in pedidos:
        campos = achatar(ped)
        linhas.append(campos)
        for c in campos:
            if c not in cols:
                cols.append(c)
    cab = ordenar_colunas(cols)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Compras"
    ws.append(cab)
    for b in ws[1]:
        b.font = openpyxl.styles.Font(bold=True)
    for lin in linhas:
        ws.append([lin.get(c, "") for c in cab])
    for i, c in enumerate(cab, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = min(max(len(c) + 2, 12), 42)
    wb.save(destino)


def consultar(token, pedido, com_anexos=False):
    ident = codigo_para_id(pedido)
    if not ident:
        return (False, MSG_NAO_ENCONTRADO)
    ped = so_digitos(pedido)
    url = f"{BASE}/movimentacao/{ident}?access_token={token}"
    try:
        r = requests.get(url, headers={"Accept": "application/json"}, timeout=TIMEOUT)
    except requests.RequestException as e:
        return (False, f"erro de conexao: {e}")
    if r.status_code in (401, 403):
        raise TokenInvalido()
    if r.status_code == 404:
        return (False, MSG_NAO_ENCONTRADO)
    if "json" not in (r.headers.get("Content-Type") or "").lower():
        return (False, MSG_NAO_ENCONTRADO)
    dados = r.json()
    if str(dados.get("codigo") or "") != ped:
        return (False, MSG_NAO_ENCONTRADO)
    campos = achatar(dados)
    if com_anexos:
        anexos = anexos_do_pedido(token, ident)
        campos["Qtd anexos"] = len(anexos)
        for i, a in enumerate(anexos, start=1):
            campos[f"Anexo {i}"] = f"{a['descricao']} -> {a['link']}"
    return (True, campos)


def _extrair_pedido(v):
    """Devolve o codigo de 10 digitos (string) se a celula parecer um pedido do GO
    (10 digitos comecando com 7), senao ''. Trata numero guardado como float
    (ex.: 7884122078.0, que viraria 11 digitos no so_digitos)."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    d = so_digitos(v)
    return d if (len(d) == 10 and d.startswith("7")) else ""


def ler_pedidos_planilha(path):
    """Le os numeros de pedido de uma planilha, em QUALQUER coluna. Varre todas as
    colunas de todas as abas, conta quantos pedidos validos cada coluna tem e usa a
    coluna com mais pedidos (acha a coluna certa sozinha, com ou sem cabecalho)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        melhor = []
        for ws in wb.worksheets:
            linhas = list(ws.iter_rows(values_only=True))
            if not linhas:
                continue
            largura = max((len(r) for r in linhas), default=0)
            for i in range(largura):
                achados = [p for p in (_extrair_pedido(r[i]) for r in linhas
                                       if i < len(r)) if p]
                if len(achados) > len(melhor):
                    melhor = achados
    finally:
        wb.close()
    vistos, out = set(), []
    for p in melhor:
        if p not in vistos:
            vistos.add(p)
            out.append(p)
    return out


def _valor_celula_str(v):
    """Texto limpo da celula para exibir (numero inteiro sem o '.0')."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def ler_itens_planilha(path):
    """Le, EM ORDEM, os valores da coluna de pedidos (a coluna que tem mais pedidos
    validos), exatamente como o usuario mandou - inclusive celulas que NAO sao pedido
    (para serem anotadas na saida). Pula celulas totalmente vazias e, se houver, a
    linha de cabecalho. Devolve lista de strings na ordem original (sem deduplicar)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        melhor_col, melhor_qtd, melhor_linhas = None, 0, None
        for ws in wb.worksheets:
            linhas = list(ws.iter_rows(values_only=True))
            if not linhas:
                continue
            largura = max((len(r) for r in linhas), default=0)
            for i in range(largura):
                qtd = sum(1 for r in linhas if i < len(r) and _extrair_pedido(r[i]))
                if qtd > melhor_qtd:
                    melhor_qtd, melhor_col, melhor_linhas = qtd, i, linhas
        if melhor_col is None:
            return []
        col, linhas = melhor_col, melhor_linhas
        # linha 0 e cabecalho se NAO tiver nenhum pedido valido em nenhuma coluna
        inicio = 0 if any(_extrair_pedido(c) for c in linhas[0]) else 1
        itens = []
        for r in linhas[inicio:]:
            val = _valor_celula_str(r[col]) if col < len(r) else ""
            if val:  # pula celula vazia
                itens.append(val)
        return itens
    finally:
        wb.close()


def ordenar_colunas(cols):
    return [c for c in ORDEM if c in cols] + sorted(c for c in cols if c not in ORDEM)


def salvar_lote(linhas, cols, destino):
    cabecalho = ["Numero do Pedido", "Observacao"] + ordenar_colunas(cols)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Consulta GO"
    ws.append(cabecalho)
    for b in ws[1]:
        b.font = openpyxl.styles.Font(bold=True)
    for lin in linhas:
        ws.append([lin.get(c, "") for c in cabecalho])
    for i, c in enumerate(cabecalho, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = min(max(len(c) + 2, 12), 42)
    wb.save(destino)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.fila = queue.Queue()
        self._caminho_planilha = None
        self._cpf_pedidos = []
        self._parceiro = None

        self.title("Consulta de Pedidos - GO")
        self.geometry("880x800")
        self.minsize(760, 660)
        self.configure(fg_color=COR_BG)

        self.f_tit = ctk.CTkFont(family="Segoe UI Semibold", size=22, weight="bold")
        self.f_sec = ctk.CTkFont(family="Segoe UI Semibold", size=14, weight="bold")
        self.f_txt = ctk.CTkFont(family="Segoe UI", size=13)
        self.f_sub = ctk.CTkFont(family="Segoe UI", size=11)
        self.f_mono = ctk.CTkFont(family="Consolas", size=13)
        self.f_cab = ctk.CTkFont(family="Segoe UI Semibold", size=11, weight="bold")
        self.f_cel = ctk.CTkFont(family="Segoe UI", size=12)

        cont = ctk.CTkScrollableFrame(self, fg_color="transparent")
        cont.pack(fill="both", expand=True, padx=18, pady=(14, 6))
        cont.grid_columnconfigure(0, weight=1)

        self._cabecalho(cont)
        self._card_token(cont)
        self._card_individual(cont)
        self._card_cpf(cont)
        self._card_lote(cont)

        self.lbl_status = ctk.CTkLabel(self, text="Pronto.", font=self.f_sub,
                                       text_color=COR_SUB, anchor="w")
        self.lbl_status.pack(fill="x", padx=20, pady=(0, 10))

        self._carregar_token_salvo()
        self.after(100, self._processar_fila)

    # ---------- blocos visuais ----------
    def _cabecalho(self, parent):
        box = ctk.CTkFrame(parent, fg_color="transparent")
        box.grid(row=0, column=0, sticky="ew", pady=(2, 10))
        ctk.CTkLabel(box, text="Consulta de Pedidos", font=self.f_tit,
                     text_color=COR_TXT).pack(anchor="w")
        ctk.CTkLabel(box, text="ERP Gestao Online (GO) - consulta via API",
                     font=self.f_sub, text_color=COR_SUB).pack(anchor="w")

    def _card(self, parent, row, titulo):
        card = ctk.CTkFrame(parent, fg_color=COR_CARD, corner_radius=RAIO_CARD)
        card.grid(row=row, column=0, sticky="ew", pady=8)
        card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(card, text=titulo, font=self.f_sec, text_color=COR_TXT).grid(
            row=0, column=0, sticky="w", padx=16, pady=(12, 2))
        return card

    def _card_token(self, parent):
        card = self._card(parent, 1, "1 - Token de autenticacao")
        linha = ctk.CTkFrame(card, fg_color="transparent")
        linha.grid(row=1, column=0, sticky="ew", padx=16, pady=6)
        linha.grid_columnconfigure(0, weight=1)
        self.var_token = ctk.StringVar()
        self.ent_token = ctk.CTkEntry(linha, textvariable=self.var_token, show="*",
                                      height=38, font=self.f_txt, corner_radius=RAIO_PEQ,
                                      fg_color=COR_CAMPO, border_width=1,
                                      placeholder_text="cole aqui o access_token do GO")
        self.ent_token.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.var_mostrar = ctk.BooleanVar(value=False)
        ctk.CTkSwitch(linha, text="Mostrar", variable=self.var_mostrar,
                      command=self._toggle_token, font=self.f_sub,
                      progress_color=COR_BTN).grid(row=0, column=1, padx=6)
        ctk.CTkButton(linha, text="Salvar", width=90, height=38, corner_radius=RAIO_PEQ,
                      command=self.salvar_token, fg_color=COR_BTN,
                      hover_color=COR_BTN_H).grid(row=0, column=2)

        cola = ctk.CTkFrame(card, fg_color=COR_CARD2, corner_radius=RAIO_PEQ)
        cola.grid(row=2, column=0, sticky="ew", padx=16, pady=(6, 14))
        cola.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            cola, font=self.f_sub, text_color=COR_SUB, justify="left",
            text="Colinha - como pegar o token:  no GO (Edge) ja logado, tecle F12, "
                 "abra a aba Console, cole o comando abaixo e tecle Enter.\n"
                 "O token (40 caracteres) vai para a area de transferencia; cole no campo acima.",
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))
        cmd = ctk.CTkEntry(cola, font=self.f_mono, height=36, text_color=COR_TXT,
                           fg_color=COR_CAMPO, border_width=0, corner_radius=RAIO_PEQ)
        cmd.insert(0, CMD_TOKEN)
        cmd.configure(state="readonly")
        cmd.grid(row=1, column=0, sticky="ew", padx=(12, 8), pady=(0, 12))
        ctk.CTkButton(cola, text="Copiar comando", width=130, height=36, corner_radius=RAIO_PEQ,
                      command=self.copiar_comando, fg_color=COR_BTN,
                      hover_color=COR_BTN_H).grid(row=1, column=1, padx=(0, 12), pady=(0, 12))

    def _card_individual(self, parent):
        card = self._card(parent, 2, "Consultar um pedido (numero de 10 digitos, ex.: 7884122078)")
        linha = ctk.CTkFrame(card, fg_color="transparent")
        linha.grid(row=1, column=0, sticky="ew", padx=16, pady=6)
        linha.grid_columnconfigure(0, weight=1)
        self.var_pedido = ctk.StringVar()
        ent = ctk.CTkEntry(linha, textvariable=self.var_pedido, height=38, font=self.f_txt,
                           corner_radius=RAIO_PEQ, fg_color=COR_CAMPO, border_width=1,
                           placeholder_text="numero do pedido (ex.: 7884122078)")
        ent.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ent.bind("<Return>", lambda e: self.consultar_um())
        ctk.CTkButton(linha, text="Consultar", width=110, height=38, corner_radius=RAIO_PEQ,
                      command=self.consultar_um, fg_color=COR_BTN,
                      hover_color=COR_BTN_H).grid(row=0, column=1)
        ctk.CTkButton(linha, text="Baixar NF", width=110, height=38, corner_radius=RAIO_PEQ,
                      command=self.baixar_nf, fg_color=COR_BTN2,
                      hover_color=COR_BTN2_H).grid(row=0, column=2, padx=(8, 0))

        # tabela HORIZONTAL compacta: cabecalhos em cima, valores embaixo
        self.res = ctk.CTkScrollableFrame(card, fg_color=COR_BG, corner_radius=RAIO_PEQ,
                                          height=130, orientation="horizontal")
        self.res.grid(row=2, column=0, sticky="ew", padx=16, pady=(6, 14))
        self._placeholder_resultado()

    def _card_cpf(self, parent):
        card = self._card(parent, 3, "Buscar compras por CPF (11 digitos) ou CNPJ (14 digitos)")
        linha = ctk.CTkFrame(card, fg_color="transparent")
        linha.grid(row=1, column=0, sticky="ew", padx=16, pady=6)
        linha.grid_columnconfigure(0, weight=1)
        self.var_doc = ctk.StringVar()
        ent = ctk.CTkEntry(linha, textvariable=self.var_doc, height=38, font=self.f_txt,
                           corner_radius=RAIO_PEQ, fg_color=COR_CAMPO, border_width=1,
                           placeholder_text="CPF (11 digitos) ou CNPJ (14 digitos)")
        ent.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ent.bind("<Return>", lambda e: self.buscar_cpf())
        ctk.CTkButton(linha, text="Buscar", width=100, height=38, corner_radius=RAIO_PEQ,
                      command=self.buscar_cpf, fg_color=COR_BTN,
                      hover_color=COR_BTN_H).grid(row=0, column=1)
        self.btn_salvar_cpf = ctk.CTkButton(linha, text="Salvar .xlsx", width=110, height=38,
                                            corner_radius=RAIO_PEQ, command=self.salvar_cpf,
                                            fg_color=COR_BTN2, hover_color=COR_BTN2_H,
                                            state="disabled")
        self.btn_salvar_cpf.grid(row=0, column=2, padx=(8, 0))

        self.lbl_cpf = ctk.CTkLabel(card, text="", font=self.f_sub, text_color=COR_SUB, anchor="w")
        self.lbl_cpf.grid(row=2, column=0, sticky="w", padx=16)

        # abas: Compras (lista de pedidos + baixar NF por linha) e Dados (cadastro)
        self.tabs_cpf = ctk.CTkTabview(
            card, fg_color=COR_BG, corner_radius=RAIO_PEQ,
            segmented_button_fg_color=COR_CARD2,
            segmented_button_selected_color=COR_BTN,
            segmented_button_selected_hover_color=COR_BTN_H,
            segmented_button_unselected_color=COR_CARD2,
            segmented_button_unselected_hover_color=COR_BTN2_H,
            text_color=COR_TXT)
        self.tabs_cpf.grid(row=3, column=0, sticky="ew", padx=16, pady=(6, 14))
        self.tabs_cpf.add("Compras")
        self.tabs_cpf.add("Dados")
        for nome in ("Compras", "Dados"):
            self.tabs_cpf.tab(nome).grid_columnconfigure(0, weight=1)

        self.res_compras = ctk.CTkScrollableFrame(self.tabs_cpf.tab("Compras"),
                                                  fg_color=COR_BG, corner_radius=0, height=210)
        self.res_compras.grid(row=0, column=0, sticky="nsew")
        self.res_dados = ctk.CTkScrollableFrame(self.tabs_cpf.tab("Dados"),
                                                fg_color=COR_BG, corner_radius=0, height=210)
        self.res_dados.grid(row=0, column=0, sticky="nsew")

        ctk.CTkLabel(self.res_compras, text="As compras do CPF/CNPJ aparecem aqui.",
                     font=self.f_sub, text_color=COR_SUB).grid(row=0, column=0, padx=12, pady=12)
        ctk.CTkLabel(self.res_dados, text="Os dados cadastrais do cliente aparecem aqui.",
                     font=self.f_sub, text_color=COR_SUB).grid(row=0, column=0, padx=12, pady=12)

    def _card_lote(self, parent):
        card = self._card(parent, 4, "Consultar em lote (planilha .xlsx com uma coluna de pedidos)")
        linha = ctk.CTkFrame(card, fg_color="transparent")
        linha.grid(row=1, column=0, sticky="ew", padx=16, pady=6)
        linha.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(linha, text="Selecionar planilha...", width=170, height=38,
                      corner_radius=RAIO_PEQ, command=self.escolher_planilha,
                      fg_color=COR_BTN2, hover_color=COR_BTN2_H).grid(row=0, column=0)
        self.var_arquivo = ctk.StringVar(value="nenhuma planilha selecionada")
        ctk.CTkLabel(linha, textvariable=self.var_arquivo, font=self.f_sub,
                     text_color=COR_SUB, anchor="w").grid(row=0, column=1, sticky="w", padx=10)

        self.var_anexos = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(card, text="Trazer tambem os links dos anexos (NF/XML)",
                        variable=self.var_anexos, font=self.f_sub,
                        fg_color=COR_BTN, hover_color=COR_BTN_H,
                        checkbox_width=18, checkbox_height=18, corner_radius=RAIO_CELULA).grid(
            row=2, column=0, sticky="w", padx=16, pady=(2, 4))

        self.btn_lote = ctk.CTkButton(card, text="Consultar lote e salvar resultado",
                                      height=40, corner_radius=RAIO_PEQ,
                                      command=self.consultar_lote, fg_color=COR_BTN,
                                      hover_color=COR_BTN_H, state="disabled")
        self.btn_lote.grid(row=3, column=0, sticky="ew", padx=16, pady=(6, 6))
        self.prog = ctk.CTkProgressBar(card, progress_color=COR_BTN, corner_radius=RAIO_PEQ)
        self.prog.set(0)
        self.prog.grid(row=4, column=0, sticky="ew", padx=16, pady=(2, 14))

    # ---------- celula copiavel ----------
    def _celula(self, parent, texto, width=150, fg=COR_CARD):
        """Campo somente-leitura: da para selecionar e copiar (Ctrl+C), nao editar."""
        e = ctk.CTkEntry(parent, width=width, height=30, font=self.f_cel,
                         fg_color=fg, border_width=0, corner_radius=RAIO_CELULA,
                         text_color=COR_TXT)
        e.insert(0, "" if texto is None else str(texto))
        e.configure(state="readonly")
        return e

    # ---------- resultado individual ----------
    def _limpar_resultado(self):
        for w in self.res.winfo_children():
            w.destroy()

    def _placeholder_resultado(self):
        ctk.CTkLabel(self.res, text="O resultado do pedido aparece aqui.",
                     font=self.f_sub, text_color=COR_SUB).grid(
            row=0, column=0, sticky="w", padx=12, pady=12)

    def _mostrar_resultado(self, campos):
        self._limpar_resultado()
        for col, (campo, valor) in enumerate(campos.items()):
            self.res.grid_columnconfigure(col, minsize=165)
            ctk.CTkLabel(self.res, text=campo, font=self.f_cab, text_color=COR_TXT,
                         fg_color=COR_CARD2, corner_radius=RAIO_CELULA, width=159,
                         wraplength=155, justify="left", anchor="w").grid(
                row=0, column=col, sticky="nsew", padx=2, pady=(2, 1))
            self._celula(self.res, valor, width=159).grid(
                row=1, column=col, sticky="nsew", padx=2, pady=(0, 2))

    def _mostrar_mensagem(self, texto, cor=COR_ERRO):
        self._limpar_resultado()
        ctk.CTkLabel(self.res, text=texto, font=self.f_txt, text_color=cor).grid(
            row=0, column=0, sticky="w", padx=12, pady=12)

    # ---------- token ----------
    def _toggle_token(self):
        self.ent_token.configure(show="" if self.var_mostrar.get() else "*")

    def copiar_comando(self):
        self.clipboard_clear()
        self.clipboard_append(CMD_TOKEN)
        self.update()
        self.status("Comando copiado - cole no Console do Edge (F12).")

    def _carregar_token_salvo(self):
        if keyring:
            try:
                tk_ = keyring.get_password(SERVICE, "access_token")
                if tk_:
                    self.var_token.set(tk_)
                    self.status("Token carregado do cofre (confira se ainda e valido).")
            except Exception:
                pass

    def salvar_token(self):
        tok = self.var_token.get().strip()
        if not tok:
            messagebox.showwarning("Token", "Cole o token antes de salvar.")
            return
        if keyring:
            try:
                keyring.set_password(SERVICE, "access_token", tok)
                self.status("Token salvo no cofre do Windows.")
            except Exception as e:
                messagebox.showerror("Token", f"Nao consegui salvar: {e}")
        else:
            self.status("keyring indisponivel - token usado so nesta sessao.")

    def _token_ou_aviso(self):
        tok = self.var_token.get().strip()
        if not tok:
            messagebox.showwarning("Token", "Cole o token de autenticacao do GO primeiro.")
            return None
        return tok

    # ---------- consulta individual ----------
    def consultar_um(self):
        tok = self._token_ou_aviso()
        if not tok:
            return
        ped = self.var_pedido.get().strip()
        if not ped:
            messagebox.showwarning("Pedido", "Digite o numero do pedido.")
            return
        self._mostrar_mensagem("Consultando...", cor=COR_SUB)
        self.status(f"Consultando pedido {ped}...")
        threading.Thread(target=self._worker_um, args=(tok, ped), daemon=True).start()

    def _worker_um(self, tok, ped):
        try:
            ok, res = consultar(tok, ped, com_anexos=True)
        except TokenInvalido:
            self.fila.put(("token_invalido",))
            return
        self.fila.put(("um", ok, res))

    # ---------- baixar NF / anexos ----------
    def baixar_nf(self):
        tok = self._token_ou_aviso()
        if not tok:
            return
        ped = self.var_pedido.get().strip()
        if not ped:
            messagebox.showwarning("Pedido", "Digite o numero do pedido.")
            return
        self.status(f"Baixando anexos do pedido {ped}...")
        threading.Thread(target=self._worker_baixar, args=(tok, ped), daemon=True).start()

    def _worker_baixar(self, tok, ped):
        try:
            salvos = baixar_anexos(tok, ped, DOWNLOAD_DIR)
        except TokenInvalido:
            self.fila.put(("token_invalido",))
            return
        except Exception as e:
            self.fila.put(("erro", f"Erro ao baixar: {e}"))
            return
        self.fila.put(("baixou", ped, salvos))

    # ---------- busca por CPF/CNPJ ----------
    def buscar_cpf(self):
        tok = self._token_ou_aviso()
        if not tok:
            return
        doc = so_digitos(self.var_doc.get())
        if len(doc) not in (11, 14):
            messagebox.showwarning("CPF/CNPJ", "Digite um CPF (11 digitos) ou CNPJ (14 digitos).")
            return
        self._limpar_cpf()
        self.lbl_cpf.configure(text="Buscando...")
        self.btn_salvar_cpf.configure(state="disabled")
        self.status(f"Buscando compras do documento {doc}...")
        threading.Thread(target=self._worker_cpf, args=(tok, doc), daemon=True).start()

    def _worker_cpf(self, tok, doc):
        try:
            parceiro = buscar_parceiro(tok, doc)
            if not parceiro:
                self.fila.put(("cpf_vazio",))
                return
            peds = pedidos_do_parceiro(tok, parceiro.get("id"))
        except TokenInvalido:
            self.fila.put(("token_invalido",))
            return
        except Exception as e:
            self.fila.put(("erro", f"Erro na busca: {e}"))
            return
        self.fila.put(("cpf_ok", parceiro, peds))

    def _limpar_cpf(self):
        for frame in (self.res_compras, self.res_dados):
            for w in frame.winfo_children():
                w.destroy()

    def _mostrar_cpf(self, parceiro, peds):
        self._limpar_cpf()
        self._cpf_pedidos = peds
        self._parceiro = parceiro
        nome = parceiro.get("nome") or parceiro.get("razaoSocial") or "-"
        doc = fmt_doc(parceiro.get("cpfCnpj"))
        self.lbl_cpf.configure(text=f"Cliente: {nome}   |   {doc}   |   {len(peds)} compra(s)")
        self._mostrar_compras(peds)
        self._mostrar_dados(parceiro)
        self.btn_salvar_cpf.configure(state="normal" if peds else "disabled")

    def _mostrar_compras(self, peds):
        cab = ["Pedido", "Data", "Valor", "Item", "Situacao NF", "NF"]
        larg = [115, 90, 95, 250, 100, 100]
        for c, (t, w) in enumerate(zip(cab, larg)):
            self.res_compras.grid_columnconfigure(c, minsize=w)
            ctk.CTkLabel(self.res_compras, text=t, font=self.f_cab, text_color=COR_TXT,
                         fg_color=COR_CARD2, corner_radius=RAIO_CELULA, anchor="w").grid(
                row=0, column=c, sticky="nsew", padx=1, pady=(2, 1))
        if not peds:
            ctk.CTkLabel(self.res_compras, text="Nenhuma compra encontrada para esse cliente.",
                         font=self.f_sub, text_color=COR_SUB).grid(
                row=1, column=0, columnspan=len(cab), padx=12, pady=10, sticky="w")
            return
        for i, m in enumerate(peds, start=1):
            codigo = str(m.get("codigo") or "")
            vals = [
                codigo,
                fmt_data(m.get("dataNegociacao"))[:10],
                f"R$ {m.get('valorTotal')}",
                m.get("movimentacaoItensDescricao") or "",
                m.get("situacaoNf") or "-",
            ]
            bg = COR_CARD if i % 2 else COR_BG
            for c, v in enumerate(vals):
                self._celula(self.res_compras, v, width=larg[c] - 6, fg=bg).grid(
                    row=i, column=c, sticky="ew", padx=1, pady=1)
            ctk.CTkButton(self.res_compras, text="Baixar NF", width=92, height=26,
                          corner_radius=RAIO_CELULA, font=self.f_sub,
                          fg_color=COR_BTN2, hover_color=COR_BTN2_H,
                          command=lambda cod=codigo: self._baixar_nf_linha(cod)).grid(
                row=i, column=5, padx=3, pady=2)

    def _mostrar_dados(self, parceiro):
        dados = achatar_parceiro(parceiro)
        if not dados:
            ctk.CTkLabel(self.res_dados, text="Sem dados cadastrais.",
                         font=self.f_sub, text_color=COR_SUB).grid(row=0, column=0, padx=12, pady=10)
            return
        self.res_dados.grid_columnconfigure(0, minsize=160)
        self.res_dados.grid_columnconfigure(1, weight=1)
        for i, (rot, val) in enumerate(dados):
            bg = COR_CARD if i % 2 else COR_BG
            ctk.CTkLabel(self.res_dados, text=rot, font=self.f_cab, text_color=COR_SUB,
                         fg_color=bg, anchor="w").grid(row=i, column=0, sticky="nsew", padx=1, pady=1)
            self._celula(self.res_dados, val, width=380, fg=bg).grid(
                row=i, column=1, sticky="ew", padx=1, pady=1)

    def _baixar_nf_linha(self, codigo):
        tok = self._token_ou_aviso()
        if not tok:
            return
        self.status(f"Baixando anexos do pedido {codigo}...")
        threading.Thread(target=self._worker_baixar, args=(tok, codigo), daemon=True).start()

    def salvar_cpf(self):
        if not self._cpf_pedidos:
            return
        destino = filedialog.asksaveasfilename(
            title="Salvar compras como", defaultextension=".xlsx",
            initialfile="compras_cliente.xlsx", filetypes=[("Planilha Excel", "*.xlsx")])
        if not destino:
            return
        try:
            salvar_pedidos_xlsx(self._cpf_pedidos, destino)
        except Exception as e:
            messagebox.showerror("Salvar", f"Erro ao salvar: {e}")
            return
        self.status(f"Salvo: {destino}")
        messagebox.showinfo("Salvo", f"{len(self._cpf_pedidos)} compra(s) salva(s) em:\n{destino}")

    # ---------- lote ----------
    def escolher_planilha(self):
        caminho = filedialog.askopenfilename(
            title="Escolha a planilha com os pedidos",
            filetypes=[("Planilhas Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if not caminho:
            return
        self._caminho_planilha = caminho
        self.var_arquivo.set(os.path.basename(caminho))
        self.btn_lote.configure(state="normal")
        self.status("Planilha selecionada. Clique em 'Consultar lote'.")

    def consultar_lote(self):
        tok = self._token_ou_aviso()
        if not tok or not self._caminho_planilha:
            return
        try:
            itens = ler_itens_planilha(self._caminho_planilha)
        except Exception as e:
            messagebox.showerror("Planilha", f"Nao consegui ler a planilha:\n{e}")
            return
        if not itens:
            messagebox.showwarning(
                "Planilha",
                "Nao encontrei uma coluna com pedidos (10 digitos comecando com 7).\n"
                "Confira se os pedidos estao numa coluna da planilha.")
            return
        destino = filedialog.asksaveasfilename(
            title="Salvar resultado como", defaultextension=".xlsx",
            initialfile="resultado_consulta_pedidos.xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")])
        if not destino:
            return
        self.btn_lote.configure(state="disabled")
        self.prog.set(0)
        self.status(f"Consultando {len(itens)} linha(s)...")
        threading.Thread(target=self._worker_lote, args=(tok, itens, destino),
                         daemon=True).start()

    def _worker_lote(self, tok, itens, destino):
        # Mantem a ORDEM original e gera UMA linha de saida por valor enviado.
        # Valor que nao e pedido -> observacao; pedido nao achado -> observacao.
        linhas, cols = [], []
        com_anexos = bool(self.var_anexos.get())
        total = len(itens)
        for i, item in enumerate(itens, start=1):
            pedido = _extrair_pedido(item)
            if not pedido:
                lin = {"Numero do Pedido": item, "Observacao": MSG_INVALIDO}
            else:
                try:
                    ok, res = consultar(tok, pedido, com_anexos=com_anexos)
                except TokenInvalido:
                    self.fila.put(("token_invalido",))
                    return
                if ok:
                    lin = {"Numero do Pedido": item, "Observacao": MSG_ENCONTRADO}
                    lin.update(res)
                    for c in res:
                        if c not in cols:
                            cols.append(c)
                else:
                    lin = {"Numero do Pedido": item, "Observacao": res}
            linhas.append(lin)
            self.fila.put(("progresso", i, total))
        try:
            salvar_lote(linhas, cols, destino)
        except Exception as e:
            self.fila.put(("erro", f"Erro ao salvar: {e}"))
            return
        achou = sum(1 for l in linhas if l.get("Observacao") == MSG_ENCONTRADO)
        self.fila.put(("lote_ok", destino, achou, len(linhas) - achou))

    # ---------- fila / UI ----------
    def _processar_fila(self):
        try:
            while True:
                msg = self.fila.get_nowait()
                tipo = msg[0]
                if tipo == "um":
                    _, ok, res = msg
                    if ok:
                        self._mostrar_resultado(res)
                        self.status(f"Pedido consultado: {len(res)} campos.")
                    else:
                        self._mostrar_mensagem(res)
                        self.status(res)
                elif tipo == "baixou":
                    _, ped, salvos = msg
                    if salvos:
                        nomes = "\n".join(os.path.basename(p) for _, p in salvos)
                        self.status(f"{len(salvos)} anexo(s) baixado(s) em Downloads.")
                        messagebox.showinfo(
                            "Anexos baixados",
                            f"Pedido {ped}: {len(salvos)} arquivo(s) salvo(s) em:\n"
                            f"{DOWNLOAD_DIR}\n\n{nomes}")
                    else:
                        self.status("Nenhum anexo encontrado para esse pedido.")
                        messagebox.showinfo("Anexos",
                                            f"Pedido {ped}: nenhum anexo encontrado.")
                elif tipo == "cpf_ok":
                    _, parceiro, peds = msg
                    self._mostrar_cpf(parceiro, peds)
                    self.status(f"{len(peds)} compra(s) encontrada(s).")
                elif tipo == "cpf_vazio":
                    self._limpar_cpf()
                    self.lbl_cpf.configure(text="Nenhum cliente com esse CPF/CNPJ.")
                    self.btn_salvar_cpf.configure(state="disabled")
                    self.status("Nenhum cliente encontrado.")
                elif tipo == "progresso":
                    _, i, n = msg
                    self.prog.set(i / n)
                    self.status(f"Consultando... {i}/{n}")
                elif tipo == "lote_ok":
                    _, destino, achou, falhou = msg
                    self.btn_lote.configure(state="normal")
                    self.prog.set(1)
                    self.status(f"Pronto: {achou} encontrado(s), {falhou} com observacao.")
                    messagebox.showinfo("Lote concluido",
                                        f"{achou} pedido(s) encontrado(s)\n"
                                        f"{falhou} com observacao (nao encontrado / invalido)\n\n"
                                        f"A planilha mantem a ordem que voce enviou.\n\n"
                                        f"Salvo em:\n{destino}")
                elif tipo == "token_invalido":
                    self.btn_lote.configure(state="normal")
                    self.status("Token recusado pelo GO (vencido ou invalido).")
                    self._mostrar_mensagem("Token recusado pelo GO (vencido ou invalido).")
                    messagebox.showerror("Token invalido",
                                         "O GO recusou o token (vencido ou invalido).\n\n"
                                         "Use a colinha para pegar um novo, cole no campo e "
                                         "tente de novo.")
                elif tipo == "erro":
                    self.btn_lote.configure(state="normal")
                    self.status(msg[1])
                    messagebox.showerror("Erro", msg[1])
        except queue.Empty:
            pass
        self.after(100, self._processar_fila)

    def status(self, txt):
        self.lbl_status.configure(text=txt)


def main():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    App().mainloop()


if __name__ == "__main__":
    main()
