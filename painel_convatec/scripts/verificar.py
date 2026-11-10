#!/usr/bin/env python3
"""Confere os dois HTMLs antes da entrega: total igual à base, cliente sem transportadora,
sem marcador, sem segredo, sem e-mail e sem dependência externa.

Uso, na raiz do projeto:  python -I scripts/verificar.py
"""
import json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gerar_painel import RAIZ, ler_base, mapear

SEGREDO = re.compile(r"gh[pousr]_[A-Za-z0-9]{20,}|github_pat_|\bsk-[A-Za-z0-9_-]{24,}|\bAKIA[0-9A-Z]{16}\b|\beyJ[A-Za-z0-9_-]{20,}\.eyJ")
EMAIL = re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+")
erros = []
cfg = json.load(open(os.path.join(RAIZ, "config.json"), encoding="utf-8"))
cab, linhas = ler_base(os.path.join(RAIZ, cfg["entrada"]))
cols, regs, _ = mapear(cfg, cab, linhas)
idx_nf = cols.index("nf")
for modo, destino in cfg["saidas"].items():
    caminho = os.path.join(RAIZ, destino)
    if not os.path.isfile(caminho):
        erros.append(f"{modo}: HTML não gerado")
        continue
    html = open(caminho, encoding="utf-8").read()
    if re.search(r"__[A-Z_]+__", html):
        erros.append(f"{modo}: sobrou marcador __X__")
    m = re.search(r'<script id="dados" type="application/json">(.*?)</script>', html, re.S)
    d = json.loads(m.group(1).replace("<\\/", "</"))
    if len(d["rows"]) != len(regs):
        erros.append(f"{modo}: HTML tem {len(d['rows'])} notas, base tem {len(regs)}")
    if sorted(r[d["cols"].index("nf")] for r in d["rows"]) != sorted(r[idx_nf] for r in regs):
        erros.append(f"{modo}: lista de NFs diferente da base")
    if modo == "cliente":
        vaz = [c for c in cfg["so_interno"] if c in d["cols"]]
        if vaz:
            erros.append(f"cliente: colunas internas no HTML: {vaz}")
        tsps = {r[cols.index("tsp")] for r in regs if "tsp" in cols} - {""}
        achou = [t for t in tsps if t in html]
        if achou:
            erros.append(f"cliente: nome de transportadora no HTML: {achou[:3]}")
    if SEGREDO.search(html):
        erros.append(f"{modo}: possível segredo no HTML")
    if EMAIL.search(re.sub(r"data:image/[^\"]+", "", html)):
        erros.append(f"{modo}: e-mail no HTML: {EMAIL.search(html).group(0)}")
    if re.search(r'(src|href)="https?://', html):
        erros.append(f"{modo}: HTML depende de arquivo externo")
    print(f"{modo}: {len(d['rows'])} notas, {len(d['cols'])} colunas, {os.path.getsize(caminho) // 1024} KB")
if erros:
    print("FALHOU:")
    for e in erros:
        print(" -", e)
    sys.exit(1)
print(f"OK: {len(regs)} notas conferidas com a base nos dois painéis, cliente sem transportadora, sem e-mail, sem segredo, sem dependência externa.")
