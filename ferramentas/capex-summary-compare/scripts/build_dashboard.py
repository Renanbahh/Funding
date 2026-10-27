#!/usr/bin/env python3
"""Gera o dashboard HTML 'Odata CapEx Control' a partir da aba 'CapEx Summary' do arquivo PROD.

Uso:
  python3 -I build_dashboard.py --folder <pasta> --template <template.html> --out <saida.html>
  python3 -I build_dashboard.py --prod <arquivo PROD.xlsx> --template ... --out ...

Regras (batem com o Power BI de referencia):
- Projetos: linhas com Site e Phase/Project entre o cabecalho e a primeira linha so com texto na
  coluna A (hoje 'Expansions'). Linhas 'Total' ficam de fora.
- Excluidos: categoria 'Fitout Revenue' e 'Utility Revenue', e Type 'Reimbursement'.
- Type exibido: categoria OFE -> 'Spec - OFE'; categoria Fitout -> 'Fitout'; senao o Type da planilha.
- Spent = meses <= data de corte; Remaining = meses posteriores.
"""
import argparse, datetime as dt, glob, json, os, re, sys, warnings
warnings.filterwarnings("ignore")
import openpyxl
import zipfile
from zoneinfo import ZoneInfo

TZ = ZoneInfo("America/Sao_Paulo")


def last_saved(path):
    """Data/hora do ultimo salvamento gravada pelo Excel (docProps/core.xml, em UTC), no horario local.
    Usa isso em vez da data do sistema de arquivos, que muda ao copiar ou sincronizar o arquivo."""
    try:
        xml = zipfile.ZipFile(path).read("docProps/core.xml").decode("utf-8", "ignore")
        m = re.search(r"<dcterms:modified[^>]*>([^<]+)<", xml)
        if m:
            t = dt.datetime.fromisoformat(m.group(1).replace("Z", "+00:00"))
            if t.tzinfo is None:
                t = t.replace(tzinfo=dt.timezone.utc)
            return t.astimezone(TZ)
    except Exception:
        pass
    return dt.datetime.fromtimestamp(os.path.getmtime(path), TZ)

SHEET = "CapEx Summary"
# Projetos de ajuste (ex.: "Ajuste Brasil", "Ajuste Mexico") ficam sempre fora da analise, nas duas versoes.
EXCLUDE_PROJECT = re.compile(r"^\s*ajuste\b", re.I)
EXCLUDE_CATEGORIES = {"Fitout Revenue", "Utility Revenue"}
EXCLUDE_TYPES = {"Reimbursement"}
MONTHS_PT = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]


def txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def name_date(path):
    m = re.search(r"(\d{2})-(\d{2})-(\d{2,4})", os.path.basename(path))
    if not m:
        return None
    d, mth, y = map(int, m.groups())
    y = y + 2000 if y < 100 else y
    try:
        return dt.date(y, mth, d)
    except ValueError:
        return None


def pick(folder, pattern, label):
    files = [f for f in glob.glob(os.path.join(folder, "*.xls*")) if not os.path.basename(f).startswith("~$")]
    hit = [f for f in files if re.search(pattern, os.path.basename(f), re.I)]
    if not hit:
        sys.exit(f"ERRO: nenhum arquivo com '{label}' no nome em {folder}: {[os.path.basename(f) for f in files]}")
    return max(hit, key=lambda f: (name_date(f) or dt.date.min, os.path.getmtime(f)))


def pick_prod(folder):
    return pick(folder, r"\bprod\b", "PROD")


def pick_cons(folder):
    return pick(folder, r"\bcons", "Cons")


def raw_type(category, typ):
    """Type da pagina de comparacao: categoria Fitout -> 'Fitout'; 'Non-Spec'/'Non-spec' unificados."""
    if category == "Fitout":
        return "Fitout"
    if typ.lower() == "non-spec":
        return "Non-spec"
    return typ or "(vazio)"


def display_type(category, typ):
    if category == "OFE":
        return "Spec - OFE"
    if category == "Fitout":
        return "Fitout"
    return typ or "(vazio)"


DETAIL_SHEET = "CapEx Detail USD"


def read_contingency(wb):
    """Soma o 'Remaining' das linhas de categoria 'Contingency/Escalation' por (Site, Phase/Project)."""
    if DETAIL_SHEET not in wb.sheetnames:
        print(f"AVISO: aba '{DETAIL_SHEET}' nao encontrada; Contingency fica zerada", flush=True)
        return {}
    out, hdr = {}, None
    for r in wb[DETAIL_SHEET].iter_rows(values_only=True):
        r = list(r)
        if hdr is None:
            names = [txt(x) for x in r]
            if "Category" in names and "Remaining" in names and "Site" in names:
                hdr = {n: i for i, n in reversed(list(enumerate(names))) if n}
            continue
        site, ph, cat = (txt(r[hdr[k]]) if hdr[k] < len(r) else "" for k in ("Site", "Phase/Project", "Category"))
        if site and ph and cat.lower() == "contingency/escalation":
            out[(site, ph)] = out.get((site, ph), 0.0) + num(r[hdr["Remaining"]])
    if hdr is None:
        print(f"AVISO: cabecalho da aba '{DETAIL_SHEET}' nao encontrado; Contingency fica zerada", flush=True)
    return out


def read(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if SHEET not in wb.sheetnames:
        sys.exit(f"ERRO: aba '{SHEET}' nao encontrada")
    rows = [list(r) for r in wb[SHEET].iter_rows(values_only=True)]
    contingency = read_contingency(wb)
    wb.close()

    hi = next((i for i, r in enumerate(rows[:30]) if len(r) > 1 and txt(r[0]) == "Site" and txt(r[1]) == "Phase/Project"), None)
    if hi is None:
        sys.exit("ERRO: cabecalho 'Site | Phase/Project' nao encontrado")
    hdr = rows[hi]
    col = {}
    for j, h in enumerate(hdr):
        k = txt(h)
        if k and not isinstance(h, dt.datetime) and k not in col:
            col[k] = j
    need = ["Capex categorization", "Type", "Other", "Commit?", "Country", "Remaining Capex"]
    miss = [n for n in need if n not in col]
    if miss:
        sys.exit(f"ERRO: colunas ausentes: {miss}")

    months = []
    for j, h in enumerate(hdr):
        if isinstance(h, dt.datetime):
            if months and j != months[-1][0] + 1:
                break
            months.append((j, h))
    cutoff = None
    for j in range(min(len(rows[hi - 1]), col["Remaining Capex"] + 3)):
        if isinstance(rows[hi - 1][j], dt.datetime):
            cutoff = rows[hi - 1][j]
    if cutoff is None:
        sys.exit("ERRO: data de corte nao encontrada acima de 'Remaining Capex'")

    projects, excluded = [], {"rows": 0, "total": 0.0}
    keyrows, pending, seen = [], [], {}   # para as possiveis renomeacoes (site + bloco)
    for i in range(hi + 1, len(rows)):
        r = rows[i] + [None] * (len(hdr) - len(rows[i]))
        a, b = txt(r[0]), txt(r[1])
        if txt(r[col["Commit?"]]) == "Total":
            for kr in pending:
                kr["block"] = a          # o bloco e o nome da linha 'Total' logo abaixo
            pending = []
            continue
        if a and not b:
            break
        if not (a and b):
            continue
        if EXCLUDE_PROJECT.match(b):
            continue
        cat, typ = txt(r[col["Capex categorization"]]), txt(r[col["Type"]])
        if typ.lower() != "reimbursement":
            base = f"{a} | {b}"
            seen[base] = seen.get(base, 0) + 1
            kr = {"key": base if seen[base] == 1 else f"{base} #{seen[base]}", "site": a, "block": "",
                  "efc": num(r[col["EFC"]]) if "EFC" in col else 0.0}
            keyrows.append(kr); pending.append(kr)
        vals = [round(num(r[j]), 2) for j, _ in months]
        if cat in EXCLUDE_CATEGORIES or typ in EXCLUDE_TYPES:
            excluded["rows"] += 1
            excluded["total"] += sum(vals)
            continue
        projects.append({
            "p": f"{a} {b}", "site": a, "row": i + 1,
            "country": txt(r[col["Country"]]) or "(vazio)",
            "other": txt(r[col["Other"]]) or "(vazio)",
            "type": display_type(cat, typ),
            "type2": raw_type(cat, typ),
            "cat": cat or "(vazio)",
            "commit": txt(r[col["Commit?"]]) or "(vazio)",
            # pagina 4 (sem FX): Budget = 'Total Budget'; EFC = coluna 'C' (EFC sem FX Variation);
            # Contingency = Remaining das linhas 'Contingency/Escalation' da aba CapEx Detail USD
            "bud": round(num(r[col["Total Budget"]]), 2) if "Total Budget" in col else 0.0,
            "efc": round(num(r[col["C"]]), 2) if "C" in col else 0.0,
            "cont": round(contingency.get((a, b), 0.0), 2),
            "v": vals,
        })
    return {
        "file": os.path.basename(path),
        "last_update": last_saved(path).strftime("%d/%m/%y %H:%M"),
        "file_date": (name_date(path) or dt.date.fromtimestamp(os.path.getmtime(path))).strftime("%d/%m/%y"),
        "cutoff": cutoff.strftime("%Y-%m"),
        "cutoff_label": f"{MONTHS_PT[cutoff.month - 1].capitalize()}-{cutoff.strftime('%y')}",
        "months": [h.strftime("%Y-%m") for _, h in months],
        "projects": projects,
        "keyrows": keyrows,
        "excluded": excluded,
        "generated": dt.datetime.now().strftime("%d/%m/%y %H:%M"),
    }


def renames(old, new):
    """Possiveis renomeacoes: projeto removido e projeto novo no mesmo site e bloco (o de EFC mais proximo)."""
    O = {k["key"]: k for k in old["keyrows"]}
    N = {k["key"]: k for k in new["keyrows"]}
    added = [k for k in N if k not in O]
    out = []
    for r in [k for k in O if k not in N]:
        po = O[r]
        cands = [x for x in added if N[x]["site"] == po["site"] and N[x]["block"] == po["block"]]
        if cands:
            best = min(cands, key=lambda x: abs(N[x]["efc"] - po["efc"]))
            out.append({"old": r, "new": best, "site": po["site"]})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--folder")
    ap.add_argument("--prod")
    ap.add_argument("--cons")
    ap.add_argument("--prev-url", help="link do arquivo anterior no SharePoint (so para exibir na pagina inicial)")
    ap.add_argument("--cur-url", help="link do arquivo atual no SharePoint (so para exibir na pagina inicial)")
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    prod = a.prod or (pick_prod(a.folder) if a.folder else None)
    cons = a.cons or (pick_cons(a.folder) if a.folder else None)
    if not prod or not cons:
        sys.exit("Informe --folder ou --prod e --cons")
    print(f"PROD (atual): {os.path.basename(prod)}\nCons (anterior): {os.path.basename(cons)}", flush=True)
    data = read(prod)
    prev = read(cons)
    cmp = {"renames": renames(prev, data)}
    meta = {"source": "claude", "generated": dt.datetime.now().strftime("%d/%m/%Y %H:%M"),
            "prev_link": a.prev_url or "", "cur_link": a.cur_url or ""}
    for d in (data, prev):
        d.pop("keyrows", None)
    js = json.dumps({"cur": data, "prev": prev, "cmp": cmp, "meta": meta}, ensure_ascii=False, separators=(",", ":"), default=str)
    html = open(a.template, encoding="utf-8").read().replace("__DATA__", js.replace("</", "<\\/"))
    open(a.out, "w", encoding="utf-8").write(html)

    # resumo para conferencia
    yr = int(data["cutoff"][:4])
    idx = [i for i, m in enumerate(data["months"]) if m.startswith(str(yr))]
    sp = sum(p["v"][i] for p in data["projects"] for i in idx if data["months"][i] <= data["cutoff"])
    rm = sum(p["v"][i] for p in data["projects"] for i in idx if data["months"][i] > data["cutoff"])
    print(json.dumps({"projetos": len(data["projects"]), "corte": data["cutoff"],
                      f"spent_{yr}_M": round(sp / 1e6, 1), f"remaining_{yr}_M": round(rm / 1e6, 1),
                      f"total_{yr}_M": round((sp + rm) / 1e6, 1),
                      "linhas_excluidas": data["excluded"]["rows"],
                      "anterior_projetos": len(prev["projects"]), "anterior_corte": prev["cutoff"], "saida": a.out}, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
