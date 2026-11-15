# Riepilogo settimanale Registro Pranzi: legge i dati da Firebase e li manda per mail
# con un Excel completo (pranzi, versamenti, ski room) e la copia di sicurezza (JSON).
import os, sys, io, json, smtplib, ssl, datetime, urllib.request, urllib.error
from email.message import EmailMessage

PREZZO = 13
SKI_LABEL = {"completa": "Preparazione completa", "sciolina": "Solo sciolina"}
API_KEY = "AIzaSyCt9D9Uomu3rNhDq1qDICuWnUdZJjmLlow"
PROJECT = "gestionale-422df"


def http(url, data=None, headers=None):
    h = dict(headers or {})
    if os.environ.get("SITE_URL"):
        h["Referer"] = os.environ["SITE_URL"]
    req = urllib.request.Request(url, data=data, headers=h)
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        sys.exit("Errore %s da %s: %s" % (e.code, url.split("?")[0], e.read().decode()[:400]))


def decodifica(v):
    if "stringValue" in v: return v["stringValue"]
    if "integerValue" in v: return int(v["integerValue"])
    if "doubleValue" in v: return v["doubleValue"]
    if "booleanValue" in v: return v["booleanValue"]
    if "nullValue" in v: return None
    if "timestampValue" in v: return v["timestampValue"]
    if "arrayValue" in v: return [decodifica(x) for x in v["arrayValue"].get("values", [])]
    if "mapValue" in v: return {k: decodifica(x) for k, x in v["mapValue"].get("fields", {}).items()}
    return None


def scarica():
    body = json.dumps({"email": os.environ["PRANZI_EMAIL"], "password": os.environ["PRANZI_PASSWORD"],
                       "returnSecureToken": True}).encode()
    tok = http("https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=" + API_KEY,
               body, {"Content-Type": "application/json"})["idToken"]
    doc = http("https://firestore.googleapis.com/v1/projects/%s/databases/(default)/documents/pranzi/stato" % PROJECT,
               None, {"Authorization": "Bearer " + tok})
    stato = decodifica({"mapValue": {"fields": doc["fields"]}})["state"]
    if isinstance(stato, str):
        stato = json.loads(stato)
    return stato


def it(d):
    try: return datetime.date.fromisoformat(d).strftime("%d/%m/%Y")
    except Exception: return d or "(senza data)"


def calcola(p):
    pagato = float(p.get("pagato") or 0)
    giorni = sorted(d for d, v in (p.get("attendance") or {}).items() if v == 1)
    ski = sum(float(x.get("eur") or 0) for x in (p.get("skiLog") or []))
    pasti = len(giorni) * PREZZO
    saldo = pagato - pasti - ski
    return pagato, giorni, pasti, ski, saldo


def euro(x):
    return ("%.2f" % x).replace(".", ",") + " €" if x % 1 else "%d €" % x


def costruisci(stato):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    wb = Workbook(); wb.remove(wb.active)
    testa = PatternFill("solid", fgColor="14306B")

    def foglio(nome, intestazioni, righe):
        ws = wb.create_sheet(nome[:31])
        ws.append(intestazioni)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF"); c.fill = testa; c.alignment = Alignment(wrap_text=True, vertical="center")
        for r in righe: ws.append(r)
        for i, h in enumerate(intestazioni, 1):
            w = max([len(str(h))] + [len(str(r[i-1])) for r in righe if i-1 < len(r)])
            ws.column_dimensions[get_column_letter(i)].width = min(max(10, w + 2), 40)
        ws.freeze_panes = "A2"
        return ws

    testo = []
    for s in stato.get("seasons", [])[-1:]:   # solo la stagione corrente (la copia JSON le contiene tutte)
        persone = sorted(stato.get("data", {}).get(s["id"], []), key=lambda p: (p.get("cognome", ""), p.get("nome", "")))
        if not persone: continue
        lab = s.get("label", s["id"]).replace("/", "-")
        riep, pranzi, vers, ski = [], [], [], []
        tot = {"pagato": 0, "pasti": 0, "ski": 0, "rest": 0, "dapagare": 0}
        for p in persone:
            nome = "%s %s" % (p.get("cognome", ""), p.get("nome", ""))
            pagato, giorni, pasti, skis, saldo = calcola(p)
            riep.append([p.get("cognome", ""), p.get("nome", ""), p.get("gruppo") or "", p.get("anno") or "", pagato, len(giorni), pasti, skis, saldo,
                         max(saldo, 0), max(-saldo, 0)])
            tot["pagato"] += pagato; tot["pasti"] += pasti; tot["ski"] += skis
            tot["rest"] += max(saldo, 0); tot["dapagare"] += max(-saldo, 0)
            for d in giorni: pranzi.append([nome, it(d), d, PREZZO])
            vs = sorted(p.get("versamenti") or [], key=lambda v: v.get("d", ""))
            somma = 0
            for v in vs:
                vers.append([nome, it(v.get("d")), v.get("d"), float(v.get("eur") or 0)]); somma += float(v.get("eur") or 0)
            if pagato - somma > 0.001:
                vers.append([nome, "(versamenti precedenti, senza data)", "", pagato - somma])
            for x in sorted(p.get("skiLog") or [], key=lambda x: x.get("d", "")):
                ski.append([nome, it(x.get("d")), x.get("d"), SKI_LABEL.get(x.get("t"), x.get("t") or ""), x.get("u") or "", x.get("g") or "", float(x.get("eur") or 0)])
        riep.append(["TOTALE", "", "", "", tot["pagato"], "", tot["pasti"], tot["ski"], "", tot["rest"], tot["dapagare"]])
        ws = foglio("Riepilogo " + lab, ["Cognome", "Nome", "Gruppo", "Anno", "Versato €", "Pasti mangiati", "Costo pasti €", "Ski room €",
                                        "Saldo €", "Da restituire €", "Da pagare €"], riep)
        for c in ws[ws.max_row]: c.font = Font(bold=True)
        foglio("Pranzi " + lab, ["Atleta", "Data", "Data (ISO)", "Costo €"], pranzi)
        foglio("Versamenti " + lab, ["Atleta", "Data", "Data (ISO)", "Importo €"], vers)
        foglio("Ski room " + lab, ["Atleta", "Data", "Data (ISO)", "Trattamento", "Uso", "Disciplina", "Importo €"], ski)
        testo.append((s.get("label", s["id"]), persone, tot, len(pranzi), len(ski)))
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue(), testo


def corpo(testo, oggi):
    righe = ["Riepilogo Registro Pranzi del %s" % oggi.strftime("%d/%m/%Y"), ""]
    for lab, persone, tot, npr, nsk in testo:
        righe += ["STAGIONE %s" % lab,
                  "Atleti: %d · pasti registrati: %d · trattamenti ski room: %d" % (len(persone), npr, nsk),
                  "Totale versato: %s · pasti: %s · ski room: %s" % (euro(tot["pagato"]), euro(tot["pasti"]), euro(tot["ski"])),
                  "Da restituire in tutto: %s · da incassare in tutto: %s" % (euro(tot["rest"]), euro(tot["dapagare"])), "",
                  "%-28s %9s %7s %9s %12s" % ("Atleta", "Versato", "Pasti", "Ski room", "Saldo")]
        for p in persone:
            pagato, giorni, pasti, skis, saldo = calcola(p)
            nota = "restituire" if saldo > 0 else ("DA PAGARE" if saldo < 0 else "")
            righe.append("%-28s %9s %7d %9s %12s" % (("%s %s" % (p.get("cognome", ""), p.get("nome", "")))[:28], euro(pagato), len(giorni), euro(skis), euro(abs(saldo)) + " " + nota))
        righe.append("")
    righe += ["In allegato: l'Excel con tutti i giorni di pranzo, i versamenti e il registro dello ski room,",
              "e la copia di sicurezza (JSON) che si può ripristinare dall'app."]
    return "\n".join(righe)


def nel_periodo():
    # le mail automatiche partono solo dal 1 dicembre al 25 aprile; l'invio a mano (tasto Mail) è sempre permesso
    if os.environ.get("GITHUB_EVENT_NAME") != "schedule":
        return True
    oggi = datetime.date.today()
    return (oggi.month, oggi.day) >= (12, 1) or (oggi.month, oggi.day) <= (4, 25)

def main():
    if not nel_periodo():
        print("Fuori dal periodo 1 dicembre - 25 aprile: nessuna mail."); return
    stato = scarica()
    oggi = datetime.date.today()
    xlsx, testo = costruisci(stato)
    if os.environ.get("SOLO_FILE"):
        open("registro-%s.xlsx" % oggi, "wb").write(xlsx); open("registro-%s.txt" % oggi, "w").write(corpo(testo, oggi)); return
    m = EmailMessage()
    m["Subject"] = "Registro Pranzi e Ski room – riepilogo del %s" % oggi.strftime("%d/%m/%Y")
    m["From"] = os.environ["GMAIL_USER"]; m["To"] = os.environ["MAIL_A"]
    m.set_content(corpo(testo, oggi))
    m.add_attachment(xlsx, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="registro-%s.xlsx" % oggi)
    m.add_attachment(json.dumps({"app": "registro-pranzi", "stato": stato}, ensure_ascii=False, indent=1).encode(),
                     maintype="application", subtype="json", filename="copia-sicurezza-%s.json" % oggi)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context()) as s:
        s.login(os.environ["GMAIL_USER"], os.environ["GMAIL_APP_PASSWORD"]); s.send_message(m)
    print("Mail inviata a", m["To"])


if __name__ == "__main__":
    main()
