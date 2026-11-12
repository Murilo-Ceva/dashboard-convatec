#!/usr/bin/env python3
"""Le a base ConvaTec (CSV ou CSV.gz, ; ou ,) e gera dois HTMLs unicos: interno e cliente.

Uso, na raiz do projeto:  python -I scripts/gerar_painel.py [caminho_da_base]
O cliente nao recebe as colunas de "so_interno" (transportadora): elas nem entram no arquivo.
Sem bibliotecas externas: so Python 3.
"""
import base64, csv, datetime as dt, gzip, io, json, os, re, site, sys, unicodedata

sys.path.append(site.getusersitepackages())  # -I esconde o Pillow do usuário (só usado para reduzir os logos)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def norm(s):
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return " ".join(s.replace("�", "").split()).upper()


def ler_base(caminho):
    bruto = open(caminho, "rb").read()
    if caminho.lower().endswith(".gz"):
        bruto = gzip.decompress(bruto)
    try:
        texto = bruto.decode("utf-8-sig")
    except UnicodeDecodeError:
        texto = bruto.decode("cp1252")
    amostra = texto[:4096]
    sep = ";" if amostra.count(";") >= amostra.count(",") else ","
    linhas = list(csv.reader(io.StringIO(texto), delimiter=sep))
    return linhas[0], linhas[1:]


DATAS = {"emissao", "prev", "nova_prev", "entrega", "dt_coleta", "coleta", "prev_coleta", "data_agend", "dt_agto"}


def serial_data(v):
    """CSV salvo sem formatação traz a data como número do Excel (46335,66385). Converte para dd/mm/aaaa."""
    m = re.fullmatch(r"(\d{5})(?:[.,]\d+)?", v)
    if not m:
        return v
    d = dt.date(1899, 12, 30) + dt.timedelta(days=int(m[1]))
    return d.strftime("%d/%m/%Y")


def mapear(cfg, cab, linhas):
    """Devolve (colunas canonicas, linhas so com essas colunas, avisos). Para com erro se faltar obrigatoria."""
    cn = [norm(c) for c in cab]
    idx, avisos = {}, []
    for chave, nomes in cfg["colunas"].items():
        for nome in [chave] + nomes:
            if norm(nome) in cn:
                idx[chave] = cn.index(norm(nome))
                break
    faltam = [c for c in cfg["obrigatorias"] if c not in idx]
    if faltam:
        raise SystemExit(f"ERRO: colunas ausentes: {faltam}. Encontradas: {cab}")
    for c in cfg["colunas"]:
        if c not in idx:
            avisos.append(f"coluna '{c}' não encontrada: o painel mostra a lacuna")
    cols = list(idx)
    out = []
    for ln in linhas:
        reg = [(ln[idx[c]] if idx[c] < len(ln) else "").strip() for c in cols]
        reg = ["" if v.startswith("#") else v for v in reg]  # #N/D, #REF! do Excel
        reg = [serial_data(v) if c in DATAS else v for c, v in zip(cols, reg)]
        d = dict(zip(cols, reg))
        if d.get("nf") or d.get("cte"):
            out.append(reg)
    return cols, out, avisos


def data(v):
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", v or "")
    return dt.date(int(m[3]), int(m[2]), int(m[1])) if m else None


def situacao(d, ref):
    """Mesma regra do HTML. Previsão original (coluna PREVISAO DE ENTREGA) é a base do prazo,
    salvo cliente que agendou para depois dela: aí o prazo é a data agendada."""
    prev, ent = data(d.get("prev")), data(d.get("entrega"))
    sim = lambda k: (d.get(k) or "").strip().upper() == "SIM"
    dag = data(d.get("data_agend")) or (data(d.get("nova_prev")) if sim("agend_ok") else None)
    if prev and dag and sim("agendado") and sim("agend_ok") and dag > prev:
        prev = dag
    if not prev:
        return "sp"
    if ent:
        return "ok" if ent <= prev else "fora"
    return "venc" if prev < ref else "tr"


def imagem(caminho, largura):
    try:
        from PIL import Image
        im = Image.open(caminho).convert("RGBA")
        if im.width > largura:
            im = im.resize((largura, round(im.height * largura / im.width)), Image.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, "PNG", optimize=True)
        b = buf.getvalue()
    except ImportError:
        b = open(caminho, "rb").read()
    return "data:image/png;base64," + base64.b64encode(b).decode()


def foto(caminho, largura=1600, qualidade=72):
    """Fundo do topo (caminhão CEVA no túnel azul), reduzido para não pesar o HTML."""
    try:
        from PIL import Image
        im = Image.open(caminho).convert("RGB")
        if im.width > largura:
            im = im.resize((largura, round(im.height * largura / im.width)), Image.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=qualidade, optimize=True, progressive=True)
        b = buf.getvalue()
    except ImportError:
        b = open(caminho, "rb").read()
    return "data:image/jpeg;base64," + base64.b64encode(b).decode()


def main():
    cfg = json.load(open(os.path.join(RAIZ, "config.json"), encoding="utf-8"))
    entrada = sys.argv[1] if len(sys.argv) > 1 else os.path.join(RAIZ, cfg["entrada"])
    cab, linhas = ler_base(entrada)
    cols, regs, avisos = mapear(cfg, cab, linhas)
    if not regs:
        raise SystemExit("ERRO: nenhuma linha com NF ou CT-e.")
    ref = dt.date.fromtimestamp(os.path.getmtime(entrada)) if cfg["data_referencia"] == "arquivo" else dt.date.fromisoformat(cfg["data_referencia"])
    tpl = open(os.path.join(RAIZ, "painel", "template.html"), encoding="utf-8").read()
    js = lambda o: json.dumps(o, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    logos = {"ceva": imagem(os.path.join(RAIZ, "painel", "logo_ceva.png"), 220),
             "ceva_branco": imagem(os.path.join(RAIZ, "painel", "logo_ceva_branco.png"), 220),
             "cliente_branco": imagem(os.path.join(RAIZ, "painel", "logo_convatec_branco.png"), 220),
             "cliente": imagem(os.path.join(RAIZ, "painel", "logo_convatec.png"), 220),
             "fundo": foto(os.path.join(RAIZ, "painel", "fundo_caminhao.jpg"))}
    saidas = list(cfg["saidas"].items())
    if cfg.get("publico"):  # versao para site publico: mesmo painel, SEM nenhuma nota; a pessoa carrega a base no navegador
        saidas.append(("publico", cfg["publico"]))
    for modo, destino in saidas:
        vazio = modo == "publico"
        if vazio:
            modo = "interno"
        c = [x for x in cols if modo == "interno" or x not in cfg["so_interno"]]
        keep = [cols.index(x) for x in c]
        dados = {"cols": [] if vazio else c, "rows": [] if vazio else [[r[i] for i in keep] for r in regs], "ref": "" if vazio else ref.isoformat(),
                 "fonte": "" if vazio else os.path.basename(entrada), "gerado": dt.datetime.now().strftime("%d/%m/%Y %H:%M")}
        pub = {"modo": modo, "meta": cfg["meta_otd"], "colunas": cfg["colunas"], "so_interno": cfg["so_interno"]}
        html = (tpl.replace("__TITULO__", cfg["titulo"] + (" · visão cliente" if modo == "cliente" else ""))
                .replace("__SUBTITULO__", cfg["subtitulo"])
                .replace("__LOGO_CEVA_BRANCO__", logos["ceva_branco"]).replace("__LOGO_CEVA__", logos["ceva"]).replace("__LOGO_CLIENTE_BRANCO__", logos["cliente_branco"]).replace("__LOGO_CLIENTE__", logos["cliente"]).replace("__FUNDO__", logos["fundo"])
                .replace("__CONFIG__", js(pub)).replace("__DADOS__", js(dados)))
        saida = os.path.join(RAIZ, destino)
        os.makedirs(os.path.dirname(saida), exist_ok=True)
        open(saida, "w", encoding="utf-8", newline="\n").write(html)
        print(f"OK {'publico (sem dados)' if vazio else modo}: {0 if vazio else len(regs)} NFs | {destino}")
        copia = cfg.get("onedrive", {}).get(modo) if not vazio else None
        if copia:  # pasta compartilhada no OneDrive da CEVA: quem tem acesso abre sempre a versao mais nova
            os.makedirs(os.path.dirname(copia), exist_ok=True)
            open(copia, "w", encoding="utf-8", newline="\n").write(html)
            print(f"   copiado para o OneDrive: {copia}")
    d = [dict(zip(cols, r)) for r in regs]
    cont = {k: sum(situacao(x, ref) == k for x in d) for k in ("ok", "fora", "tr", "venc", "sp")}
    print(f"Referência {ref:%d/%m/%Y} | CT-es {len({x['cte'] for x in d})} | situação {cont}")
    for a in avisos:
        print("AVISO:", a)


if __name__ == "__main__":
    main()
