# -*- coding: utf-8 -*-
"""Gera o .zip para enviar aos colegas: pasta "Agente GO" (exe + _internal) + LEIA-ME.

Rodar depois do gerar_executavel.bat:  py empacotar.py
"""
import os
import re
import zipfile

AQUI = os.path.dirname(os.path.abspath(__file__))
DIST = os.path.join(AQUI, "dist", "Agente GO")
LEIA = os.path.join(AQUI, "LEIA-ME.txt")

with open(os.path.join(AQUI, "src", "app.py"), encoding="utf-8") as f:
    versao = re.search(r'VERSAO = "([^"]+)"', f.read()).group(1)

destino = os.path.join(os.path.dirname(AQUI), f"Agente GO v{versao}.zip")
if not os.path.isfile(os.path.join(DIST, "Agente GO.exe")):
    raise SystemExit("Rode o gerar_executavel.bat antes.")

n = 0
with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
    for raiz, _, arquivos in os.walk(DIST):
        for a in arquivos:
            cam = os.path.join(raiz, a)
            z.write(cam, os.path.join("Agente GO", os.path.relpath(cam, DIST)))
            n += 1
    z.write(LEIA, os.path.join("Agente GO", "LEIA-ME.txt"))

print(f"OK: {destino}  ({n + 1} arquivos, {os.path.getsize(destino) / 1024 / 1024:.1f} MB)")
