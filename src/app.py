# -*- coding: utf-8 -*-
"""
Agente GO - aplicativo unico (janela) para o ERP Gestao Online.

Junta todas as funcionalidades do Agente-GO num executavel:
  - Token: colar, salvar no cofre do Windows, testar, ver validade (12 h)
  - Relatorio de vendas (report 44) por periodo e unidade  -> relatorio.py
  - Consultar um pedido + baixar NF/anexos                  -> consulta_gui.py
  - Buscar compras por CPF/CNPJ (com dados cadastrais)      -> consulta_gui.py
  - Consultar pedidos em lote a partir de uma planilha      -> consulta_gui.py

Modo sem janela (para agendar), usa o token salvo no cofre:
  "Agente GO.exe" --relatorio 01/10/2026 31/10/2026 --unidade 52 [--saida arq.xlsx]
  O resultado vai para %LOCALAPPDATA%\\Agente GO\\agente_go.log (codigo de saida 0 = ok).
"""
import argparse
import os
import queue
import sys
import threading
import time

import customtkinter as ctk
from tkinter import filedialog, messagebox

import consulta_gui as base
import relatorio as rel
from consulta_gui import (COR_BG, COR_BTN, COR_BTN2, COR_BTN2_H, COR_BTN_H, COR_CAMPO,
                          COR_CARD, COR_CARD2, COR_ERRO, COR_SUB, COR_TXT, RAIO_CARD,
                          RAIO_PEQ, CMD_TOKEN)

VERSAO = "1.0"
COR_OK = "#9CC79C"
PASTA_APP = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "Agente GO")


def recurso(nome):
    """Caminho de arquivo empacotado (funciona no .exe e rodando o .py)."""
    raiz = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(raiz, nome)


def fmt_bytes(n):
    return f"{n / 1024 / 1024:.1f} MB" if n >= 1024 * 1024 else f"{n / 1024:.0f} KB"


class AgenteGO(base.App):
    """Reaproveita os blocos da janela de pedidos e acrescenta o relatorio.

    O __init__ e reescrito (nao chama o do base.App) para montar o layout em
    abas; os metodos de consulta/lote/CPF do base.App sao usados como estao.
    """

    def __init__(self):
        ctk.CTk.__init__(self)
        self.fila = queue.Queue()        # mensagens das consultas (base.App)
        self.fila_rel = queue.Queue()    # mensagens do relatorio / token
        self._caminho_planilha = None
        self._cpf_pedidos = []
        self._parceiro = None
        self._pasta_rel = rel.DOWNLOAD_DIR
        self._baixando = False

        self.title(f"Agente GO {VERSAO}")
        altura = max(600, min(860, self.winfo_screenheight() - 90))
        self.geometry(f"920x{altura}+40+10")
        self.minsize(780, 560)
        self.configure(fg_color=COR_BG)
        try:
            self.iconbitmap(recurso("icone.ico"))
        except Exception:
            pass

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

        box = ctk.CTkFrame(cont, fg_color="transparent")
        box.grid(row=0, column=0, sticky="ew", pady=(2, 8))
        ctk.CTkLabel(box, text="Agente GO", font=self.f_tit, text_color=COR_TXT).pack(anchor="w")
        ctk.CTkLabel(box, text="Gestão Online (GO) - relatório de vendas e consulta de pedidos, via API",
                     font=self.f_sub, text_color=COR_SUB).pack(anchor="w")

        self._card_token_v2(cont)

        abas = ctk.CTkTabview(
            cont, fg_color="transparent", corner_radius=RAIO_CARD,
            segmented_button_fg_color=COR_CARD2, segmented_button_selected_color=COR_BTN,
            segmented_button_selected_hover_color=COR_BTN_H,
            segmented_button_unselected_color=COR_CARD2,
            segmented_button_unselected_hover_color=COR_BTN2_H, text_color=COR_TXT)
        abas.grid(row=2, column=0, sticky="ew", pady=(4, 0))
        nomes = ["Relatório de vendas", "Pedido", "CPF / CNPJ", "Lote (planilha)"]
        for n in nomes:
            abas.add(n)
            abas.tab(n).grid_columnconfigure(0, weight=1)

        self._card_relatorio(abas.tab(nomes[0]))
        self._card_individual(abas.tab(nomes[1]))
        self._card_cpf(abas.tab(nomes[2]))
        self._card_lote(abas.tab(nomes[3]))

        self.lbl_status = ctk.CTkLabel(self, text="Pronto.", font=self.f_sub,
                                       text_color=COR_SUB, anchor="w")
        self.lbl_status.pack(fill="x", padx=20, pady=(0, 10))

        self._carregar_token_salvo()
        self.after(100, self._processar_fila)
        self.after(120, self._processar_fila_rel)
        self.after(60_000, self._tique_validade)

    # ------------------------------------------------------------------ token
    def _card_token_v2(self, parent):
        card = self._card(parent, 1, "Token de acesso do GO")
        linha = ctk.CTkFrame(card, fg_color="transparent")
        linha.grid(row=1, column=0, sticky="ew", padx=16, pady=6)
        linha.grid_columnconfigure(0, weight=1)
        self.var_token = ctk.StringVar()
        self.ent_token = ctk.CTkEntry(linha, textvariable=self.var_token, show="*", height=38,
                                      font=self.f_txt, corner_radius=RAIO_PEQ, fg_color=COR_CAMPO,
                                      border_width=1, placeholder_text="cole aqui o access_token do GO")
        self.ent_token.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.var_mostrar = ctk.BooleanVar(value=False)
        ctk.CTkSwitch(linha, text="Mostrar", variable=self.var_mostrar, command=self._toggle_token,
                      font=self.f_sub, progress_color=COR_BTN).grid(row=0, column=1, padx=6)
        ctk.CTkButton(linha, text="Salvar", width=80, height=38, corner_radius=RAIO_PEQ,
                      command=self.salvar_token, fg_color=COR_BTN,
                      hover_color=COR_BTN_H).grid(row=0, column=2)
        ctk.CTkButton(linha, text="Testar", width=80, height=38, corner_radius=RAIO_PEQ,
                      command=self.testar_token, fg_color=COR_BTN2,
                      hover_color=COR_BTN2_H).grid(row=0, column=3, padx=(8, 0))

        self.lbl_validade = ctk.CTkLabel(card, text="", font=self.f_sub, text_color=COR_SUB, anchor="w")
        self.lbl_validade.grid(row=2, column=0, sticky="w", padx=16)

        cola = ctk.CTkFrame(card, fg_color=COR_CARD2, corner_radius=RAIO_PEQ)
        cola.grid(row=3, column=0, sticky="ew", padx=16, pady=(6, 14))
        cola.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            cola, font=self.f_sub, text_color=COR_SUB, justify="left",
            text="Como pegar o token: abra o GO no navegador já logado, tecle F12, vá na aba Console, "
                 "cole o comando abaixo e tecle Enter.\nO token (40 caracteres) vai para a área de "
                 "transferência. Cole no campo acima e clique em Salvar. Ele vale cerca de 12 horas.",
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(10, 4))
        cmd = ctk.CTkEntry(cola, font=self.f_mono, height=34, text_color=COR_TXT,
                           fg_color=COR_CAMPO, border_width=0, corner_radius=RAIO_PEQ)
        cmd.insert(0, CMD_TOKEN)
        cmd.configure(state="readonly")
        cmd.grid(row=1, column=0, sticky="ew", padx=(12, 8), pady=(0, 12))
        ctk.CTkButton(cola, text="Copiar comando", width=130, height=34, corner_radius=RAIO_PEQ,
                      command=self.copiar_comando, fg_color=COR_BTN,
                      hover_color=COR_BTN_H).grid(row=1, column=1, padx=(0, 12), pady=(0, 12))

    def _mostrar_validade(self, expira):
        if not expira:
            self.lbl_validade.configure(text="Validade: desconhecida (salve o token para acompanhar).",
                                        text_color=COR_SUB)
            return
        resta = expira - time.time()
        quando = time.strftime("%d/%m %H:%M", time.localtime(expira))
        if resta > 0:
            self.lbl_validade.configure(
                text=f"Validade prevista: até {quando} (faltam {resta / 3600:.1f} h).",
                text_color=COR_SUB if resta > 1800 else COR_ERRO)
        else:
            self.lbl_validade.configure(
                text=f"Token provavelmente VENCIDO desde {quando}. Pegue um novo no GO.",
                text_color=COR_ERRO)

    def _tique_validade(self):
        _, exp = rel.token_salvo()
        if self.var_token.get().strip():
            self._mostrar_validade(exp)
        self.after(60_000, self._tique_validade)

    def _carregar_token_salvo(self):
        tk_, exp = rel.token_salvo()
        if tk_:
            self.var_token.set(tk_)
            self._mostrar_validade(exp)
            self.status("Token carregado do cofre do Windows.")
        else:
            self.lbl_validade.configure(text="Nenhum token salvo ainda.", text_color=COR_SUB)

    def salvar_token(self):
        tok = self.var_token.get().strip()
        if not tok:
            messagebox.showwarning("Token", "Cole o token antes de salvar.")
            return
        if len(tok) != 40 and not messagebox.askyesno(
                "Token", f"O token tem {len(tok)} caracteres (o normal são 40).\nSalvar mesmo assim?"):
            return
        try:
            exp = rel.salvar_token(tok)
        except Exception as e:
            messagebox.showerror("Token", f"Não consegui salvar no cofre do Windows:\n{e}")
            return
        self._mostrar_validade(exp)
        self.status("Token salvo no cofre do Windows.")

    def testar_token(self):
        tok = self._token_ou_aviso()
        if not tok:
            return
        self.status("Testando o token no GO...")
        threading.Thread(target=self._worker_teste, args=(tok,), daemon=True).start()

    def _worker_teste(self, tok):
        try:
            self.fila_rel.put(("teste", rel.testar_token(tok)))
        except Exception as e:
            self.fila_rel.put(("rel_erro", str(e)))

    # -------------------------------------------------------------- relatorio
    def _card_relatorio(self, parent):
        card = self._card(parent, 0, "Relatório de vendas (CSV - Vendas por unidade)")

        per = ctk.CTkFrame(card, fg_color="transparent")
        per.grid(row=1, column=0, sticky="ew", padx=16, pady=(6, 2))
        ctk.CTkLabel(per, text="Período (data da negociação)", font=self.f_sub,
                     text_color=COR_SUB).grid(row=0, column=0, columnspan=4, sticky="w")
        ini, fim = rel.periodo_atalho("mes_passado")
        self.var_ini = ctk.StringVar(value=ini)
        self.var_fim = ctk.StringVar(value=fim)
        for col, (rot, var) in enumerate((("De", self.var_ini), ("Até", self.var_fim))):
            ctk.CTkLabel(per, text=rot, font=self.f_txt, text_color=COR_TXT).grid(
                row=1, column=col * 2, padx=(0 if col == 0 else 14, 6), pady=4)
            ctk.CTkEntry(per, textvariable=var, width=120, height=36, font=self.f_txt,
                         corner_radius=RAIO_PEQ, fg_color=COR_CAMPO, border_width=1,
                         placeholder_text="DD/MM/AAAA").grid(row=1, column=col * 2 + 1, pady=4)

        atal = ctk.CTkFrame(card, fg_color="transparent")
        atal.grid(row=2, column=0, sticky="w", padx=16, pady=(2, 8))
        for i, (rot, chave) in enumerate((("Hoje", "hoje"), ("Ontem", "ontem"),
                                          ("Últimos 7 dias", "ultimos_7"),
                                          ("Este mês", "este_mes"), ("Mês passado", "mes_passado"))):
            ctk.CTkButton(atal, text=rot, height=28, width=10, corner_radius=RAIO_PEQ,
                          font=self.f_sub, fg_color=COR_BTN2, hover_color=COR_BTN2_H,
                          command=lambda c=chave: self._aplicar_atalho(c)).grid(
                row=0, column=i, padx=(0, 6))

        uni = ctk.CTkFrame(card, fg_color="transparent")
        uni.grid(row=3, column=0, sticky="ew", padx=16, pady=(4, 2))
        uni.grid_columnconfigure(3, weight=1)
        ctk.CTkLabel(uni, text="Unidade", font=self.f_txt, text_color=COR_TXT).grid(
            row=0, column=0, padx=(0, 6))
        self.var_unidade = ctk.StringVar(value="52")
        ctk.CTkEntry(uni, textvariable=self.var_unidade, width=90, height=36, font=self.f_txt,
                     corner_radius=RAIO_PEQ, fg_color=COR_CAMPO, border_width=1,
                     placeholder_text="código").grid(row=0, column=1)
        ctk.CTkButton(uni, text="Conferir", width=90, height=36, corner_radius=RAIO_PEQ,
                      fg_color=COR_BTN2, hover_color=COR_BTN2_H,
                      command=self.conferir_unidade).grid(row=0, column=2, padx=8)
        self.cmb_unidades = ctk.CTkComboBox(
            uni, values=["Clique em 'Listar' para ver as unidades"], height=36,
            font=self.f_sub, fg_color=COR_CAMPO, border_width=1, button_color=COR_BTN2,
            button_hover_color=COR_BTN2_H, dropdown_fg_color=COR_CARD2,
            command=self._escolher_unidade)
        self.cmb_unidades.set("Clique em 'Listar' para escolher pela lista")
        self.cmb_unidades.configure(state="readonly")
        self.cmb_unidades.grid(row=0, column=3, sticky="ew")
        ctk.CTkButton(uni, text="Listar", width=70, height=36, corner_radius=RAIO_PEQ,
                      fg_color=COR_BTN2, hover_color=COR_BTN2_H,
                      command=self.listar_unidades).grid(row=0, column=4, padx=(8, 0))
        self._mapa_unidades = {}
        self.lbl_unidade = ctk.CTkLabel(card, text="Código 52 = E-commerce - Soluti. Deixe vazio para "
                                        "trazer todas as unidades.", font=self.f_sub,
                                        text_color=COR_SUB, anchor="w")
        self.lbl_unidade.grid(row=4, column=0, sticky="w", padx=16)

        dst = ctk.CTkFrame(card, fg_color="transparent")
        dst.grid(row=5, column=0, sticky="ew", padx=16, pady=(10, 2))
        dst.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(dst, text="Salvar em", font=self.f_txt, text_color=COR_TXT).grid(
            row=0, column=0, padx=(0, 8))
        self.var_pasta = ctk.StringVar(value=self._pasta_rel)
        ctk.CTkLabel(dst, textvariable=self.var_pasta, font=self.f_sub, text_color=COR_SUB,
                     anchor="w").grid(row=0, column=1, sticky="ew")
        ctk.CTkButton(dst, text="Alterar pasta...", width=120, height=32, corner_radius=RAIO_PEQ,
                      fg_color=COR_BTN2, hover_color=COR_BTN2_H,
                      command=self.escolher_pasta).grid(row=0, column=2)

        self.var_abrir = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(card, text="Abrir a planilha quando terminar", variable=self.var_abrir,
                        font=self.f_sub, fg_color=COR_BTN, hover_color=COR_BTN_H,
                        checkbox_width=18, checkbox_height=18).grid(
            row=6, column=0, sticky="w", padx=16, pady=(8, 4))

        self.btn_rel = ctk.CTkButton(card, text="Baixar relatório", height=42, corner_radius=RAIO_PEQ,
                                     font=self.f_sec, command=self.baixar_relatorio,
                                     fg_color=COR_BTN, hover_color=COR_BTN_H)
        self.btn_rel.grid(row=7, column=0, sticky="ew", padx=16, pady=(6, 6))
        self.prog_rel = ctk.CTkProgressBar(card, progress_color=COR_BTN, corner_radius=RAIO_PEQ,
                                           mode="indeterminate")
        self.prog_rel.grid(row=8, column=0, sticky="ew", padx=16, pady=(2, 4))
        self.prog_rel.set(0)
        self.lbl_rel = ctk.CTkLabel(card, text="Períodos longos podem levar alguns minutos para o GO gerar.",
                                    font=self.f_sub, text_color=COR_SUB, anchor="w", justify="left")
        self.lbl_rel.grid(row=9, column=0, sticky="w", padx=16, pady=(0, 14))

    def _aplicar_atalho(self, chave):
        ini, fim = rel.periodo_atalho(chave)
        self.var_ini.set(ini)
        self.var_fim.set(fim)

    def _escolher_unidade(self, texto):
        cod = self._mapa_unidades.get(texto)
        if cod:
            self.var_unidade.set(cod)
            self.lbl_unidade.configure(text=f"Unidade {cod}: {texto.split(' - ', 1)[-1]}",
                                       text_color=COR_SUB)

    def escolher_pasta(self):
        p = filedialog.askdirectory(title="Pasta onde salvar o relatório", initialdir=self._pasta_rel)
        if p:
            self._pasta_rel = p
            self.var_pasta.set(p)

    def conferir_unidade(self):
        tok = self._token_ou_aviso()
        cod = "".join(ch for ch in self.var_unidade.get() if ch.isdigit())
        if not tok:
            return
        if not cod:
            self.lbl_unidade.configure(text="Sem código: o relatório traz todas as unidades.",
                                       text_color=COR_SUB)
            return
        self.lbl_unidade.configure(text="Conferindo...", text_color=COR_SUB)
        threading.Thread(target=self._worker_unidade, args=(tok, cod), daemon=True).start()

    def _worker_unidade(self, tok, cod):
        try:
            self.fila_rel.put(("unidade", cod, rel.nome_unidade(tok, cod)))
        except rel.TokenRecusado:
            self.fila_rel.put(("rel_token",))

    def listar_unidades(self):
        tok = self._token_ou_aviso()
        if not tok:
            return
        self.status("Buscando a lista de unidades...")
        threading.Thread(target=self._worker_listar, args=(tok,), daemon=True).start()

    def _worker_listar(self, tok):
        try:
            self.fila_rel.put(("lista_unidades", rel.listar_unidades(tok)))
        except rel.TokenRecusado:
            self.fila_rel.put(("rel_token",))

    def baixar_relatorio(self):
        if self._baixando:
            return
        tok = self._token_ou_aviso()
        if not tok:
            return
        ini, fim = self.var_ini.get().strip(), self.var_fim.get().strip()
        cod = "".join(ch for ch in self.var_unidade.get() if ch.isdigit())
        try:
            iso_i, iso_f = rel.data_br_para_iso(ini), rel.data_br_para_iso(fim)
        except ValueError as e:
            messagebox.showwarning("Período", str(e))
            return
        if iso_i > iso_f:
            messagebox.showwarning("Período", "A data inicial é posterior à data final.")
            return
        destino = os.path.join(self._pasta_rel, rel.nome_arquivo_padrao(iso_i, iso_f, cod))
        if os.path.exists(destino) and not messagebox.askyesno(
                "Arquivo existente", f"Já existe:\n{destino}\n\nSubstituir?"):
            return
        self._baixando = True
        self.btn_rel.configure(state="disabled", text="Gerando no GO...")
        self.prog_rel.start()
        self.lbl_rel.configure(text="Aguardando o GO gerar a planilha...", text_color=COR_SUB)
        self.status(f"Baixando relatório de {ini} a {fim}"
                    + (f", unidade {cod}" if cod else ", todas as unidades") + "...")
        threading.Thread(target=self._worker_relatorio, args=(tok, ini, fim, cod, destino),
                         daemon=True).start()

    def _worker_relatorio(self, tok, ini, fim, cod, destino):
        def prog(n):
            self.fila_rel.put(("rel_prog", n))
        try:
            caminho, tam, seg = rel.baixar_relatorio(tok, ini, fim, cod, destino, progresso=prog)
        except rel.TokenRecusado:
            self.fila_rel.put(("rel_token",))
            return
        except PermissionError:
            self.fila_rel.put(("rel_erro", f"Não consegui gravar o arquivo. Ele está aberto no Excel?\n{destino}"))
            return
        except Exception as e:
            self.fila_rel.put(("rel_erro", str(e)))
            return
        self.fila_rel.put(("rel_ok", caminho, tam, seg))

    def _fim_download(self):
        self._baixando = False
        self.prog_rel.stop()
        self.prog_rel.set(0)
        self.btn_rel.configure(state="normal", text="Baixar relatório")

    def _processar_fila_rel(self):
        try:
            while True:
                msg = self.fila_rel.get_nowait()
                tipo = msg[0]
                if tipo == "rel_prog":
                    self.lbl_rel.configure(text=f"Recebendo... {fmt_bytes(msg[1])}")
                elif tipo == "rel_ok":
                    _, caminho, tam, seg = msg
                    self._fim_download()
                    self.lbl_rel.configure(text=f"Pronto em {seg:.0f}s ({fmt_bytes(tam)}).\n{caminho}",
                                           text_color=COR_OK)
                    self.status(f"Relatório salvo: {caminho}")
                    if self.var_abrir.get():
                        try:
                            os.startfile(caminho)
                        except Exception:
                            pass
                elif tipo == "rel_token":
                    self._fim_download()
                    self.lbl_rel.configure(text="O GO recusou o token.", text_color=COR_ERRO)
                    self._mostrar_validade(time.time() - 1)
                    self.status("Token recusado pelo GO (vencido ou inválido).")
                    messagebox.showerror("Token recusado",
                                         "O GO recusou o token (vencido ou copiado errado).\n\n"
                                         "Pegue um novo no GO (veja o passo a passo no topo da janela), "
                                         "cole, clique em Salvar e tente de novo.")
                elif tipo == "rel_erro":
                    self._fim_download()
                    self.lbl_rel.configure(text=msg[1][:300], text_color=COR_ERRO)
                    self.status("Erro - veja a mensagem.")
                    messagebox.showerror("Erro", msg[1])
                elif tipo == "teste":
                    if msg[1]:
                        self.status("Token OK: o GO aceitou.")
                        messagebox.showinfo("Token", "O GO aceitou o token.")
                    else:
                        self._mostrar_validade(time.time() - 1)
                        self.status("Token recusado pelo GO.")
                        messagebox.showerror("Token", "O GO recusou o token. Pegue um novo e salve.")
                elif tipo == "unidade":
                    _, cod, nome = msg
                    if nome:
                        self.lbl_unidade.configure(text=f"Unidade {cod}: {nome}", text_color=COR_OK)
                    else:
                        self.lbl_unidade.configure(text=f"Unidade {cod} não encontrada. Confira o código.",
                                                   text_color=COR_ERRO)
                elif tipo == "lista_unidades":
                    lista = msg[1]
                    if lista:
                        self._mapa_unidades = {f"{c} - {n}": c for c, n in lista}
                        self.cmb_unidades.configure(values=list(self._mapa_unidades))
                        self.cmb_unidades.set("Escolha uma unidade")
                        self.status(f"{len(lista)} unidade(s) carregada(s).")
                    else:
                        self.status("Não consegui listar as unidades. Digite o código e use 'Conferir'.")
        except queue.Empty:
            pass
        self.after(120, self._processar_fila_rel)


# ---------------------------------------------------------------------------
# Modo sem janela
# ---------------------------------------------------------------------------
def modo_linha_de_comando(argv):
    ap = argparse.ArgumentParser(prog="Agente GO")
    ap.add_argument("--relatorio", nargs=2, metavar=("DATA_INI", "DATA_FIM"), required=True)
    ap.add_argument("--unidade", default="")
    ap.add_argument("--saida", default="")
    args = ap.parse_args(argv)

    os.makedirs(PASTA_APP, exist_ok=True)
    log_path = os.path.join(PASTA_APP, "agente_go.log")

    def log(msg):
        linha = f"{time.strftime('%d/%m/%Y %H:%M:%S')} {msg}"
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
        if sys.stdout:
            print(linha, flush=True)

    tok, _ = rel.token_salvo()
    if not tok:
        log("ERRO: nenhum token salvo. Abra o Agente GO, cole o token e clique em Salvar.")
        return 2
    try:
        caminho, tam, seg = rel.baixar_relatorio(tok, args.relatorio[0], args.relatorio[1],
                                                 args.unidade, args.saida or None)
    except rel.TokenRecusado:
        log("ERRO: o GO recusou o token salvo (vencido). Salve um novo pela janela.")
        return 3
    except Exception as e:
        log(f"ERRO: {e}")
        return 1
    log(f"OK: {caminho} ({fmt_bytes(tam)}, {seg:.0f}s)")
    return 0


def main():
    if len(sys.argv) > 1 and sys.argv[1].startswith("--"):
        sys.exit(modo_linha_de_comando(sys.argv[1:]))
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    AgenteGO().mainloop()


if __name__ == "__main__":
    main()
