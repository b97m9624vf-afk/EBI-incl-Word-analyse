import html as html_lib
import json
import os
import re
import ssl
import subprocess
import urllib.parse
import urllib.request
from html.parser import HTMLParser

KVK_PROFILE_ID = "5C10A89D-635E-49CC-94B8-042DD533B64A"


def is_bv_entity(naam: str) -> bool:
    """Return True if the borrower name represents a B.V. or corporate entity."""
    return bool(re.search(r"\b(B\.?V\.?|N\.?V\.?|Besloten\s+Vennootschap)\b", naam or "", re.I))


def lookup_kvk(name: str, city: str = "") -> list[dict]:
    """Query the official KVK Handelsregister public search API by company name and optional city."""
    clean_name = (name or "").strip()
    if not clean_name or clean_name.lower() in ("n.v.t.", "nvt"):
        return []
    query = f"{clean_name} {city}".strip() if city else clean_name
    url = f"https://web-api.kvk.nl/zoeken/v3/search?q={urllib.parse.quote(query)}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
        "profileId": KVK_PROFILE_ID,
        "apikey": KVK_PROFILE_ID,
        "Accept": "application/json",
    }
    raw_json = ""
    try:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, context=ctx, timeout=8) as resp:
            raw_json = resp.read().decode("utf-8", errors="replace")
    except Exception:
        try:
            res = subprocess.run(
                [
                    "curl",
                    "-sL",
                    "-H",
                    f"profileId: {KVK_PROFILE_ID}",
                    "-H",
                    f"apikey: {KVK_PROFILE_ID}",
                    url,
                ],
                capture_output=True,
                text=True,
                timeout=8,
                check=True,
            )
            raw_json = res.stdout
        except Exception:
            return []

    try:
        payload = json.loads(raw_json)
        items = payload.get("data", {}).get("items", [])
        results = []
        seen = set()
        for item in items:
            if not item.get("actief", True):
                continue
            kvk_nr = item.get("kvkNummer", "")
            naam = item.get("naam", "")
            plaats = (item.get("bezoeklocatie") or {}).get("plaats") or ""
            rechtsvorm = item.get("rechtsvormOmschrijving") or item.get("rechtsvormCode") or ""
            key = (kvk_nr, naam)
            if kvk_nr and key not in seen:
                seen.add(key)
                results.append(
                    {
                        "naam": naam,
                        "kvk": kvk_nr,
                        "plaats": plaats,
                        "rechtsvorm": rechtsvorm,
                    }
                )
        return results
    except Exception:
        return []



def fetch_html(url: str) -> str:
    """Fetch HTML from a Collin Crowdfund pitch URL with fallback to curl if SSL certs fail."""
    url = url.strip()
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )
    }
    try:
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, context=ctx, timeout=20) as resp:
            return resp.read().decode("utf-8", errors="replace")
    except Exception:
        res = subprocess.run(
            ["curl", "-sL", "-A", headers["User-Agent"], url],
            capture_output=True,
            text=True,
            timeout=25,
            check=True,
        )
        return res.stdout


def clean_text(raw_html: str) -> str:
    """Convert HTML fragment to clean plain text while preserving line breaks."""
    text = re.sub(r"<br\s*/?>", "\n", raw_html, flags=re.I)
    text = re.sub(r"</p\s*>", "\n\n", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    text = html_lib.unescape(text)
    text = text.replace("\xa0", " ")
    lines = [re.sub(r"[ \t]+", " ", ln).strip() for ln in text.splitlines()]
    cleaned = "\n".join(lines)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned).strip()
    return cleaned


def parse_html_table(table_html: str) -> list[list[dict]]:
    """Parse an HTML <table> into rows of cells with text, bold, and align metadata."""
    rows = []
    for tr_match in re.finditer(r"<tr[^>]*>(.*?)</tr\s*>", table_html, flags=re.S | re.I):
        tr_inner = tr_match.group(1)
        row_cells = []
        for td_match in re.finditer(r"<(td|th)([^>]*)>(.*?)</\1\s*>", tr_inner, flags=re.S | re.I):
            attrs = td_match.group(2)
            cell_html = td_match.group(3)
            is_bold = bool(re.search(r"<(strong|b)\b", cell_html, flags=re.I)) or td_match.group(1).lower() == "th"
            align = "right" if "right" in attrs.lower() else "left"
            cell_text = clean_text(cell_html).replace("\n", " ").strip()
            row_cells.append({
                "text": cell_text,
                "bold": is_bold,
                "align": align,
            })
        if row_cells and any(c["text"] for c in row_cells):
            rows.append(row_cells)
    return rows


def format_euro(amount: float, decimals: int = 0) -> str:
    """Format number as Dutch Euro string (e.g. € 100.000,- or € 1.900,18)."""
    if decimals == 0:
        int_str = f"{int(round(amount)):,}".replace(",", ".")
        return f"€ {int_str},-"
    formatted = f"{amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"€ {formatted}"


from decimal import Decimal, ROUND_UP


def excel_roundup(val: float, decimals: int = 2) -> float:
    """Exact implementation of Excel's ROUNDUP(val, decimals) for positive numbers."""
    if val <= 0:
        return 0.0
    d = Decimal(str(round(val, 10)))
    q = Decimal("10") ** (-decimals)
    return float(d.quantize(q, rounding=ROUND_UP))


def calculate_loan_cost_summary(
    principal: float,
    months: int,
    succes_fee_pct: float = 2.0,
    admin_fee_pct: float = 0.07,
) -> dict:
    """Calculate the cost breakdown matching Zekerheden!A74:B87 in Amo EBI.xlsm."""
    pub_excl = 500.0 if principal < 250000 else 950.0
    raw_succes = principal * (succes_fee_pct / 100.0)
    succes_plus_pub_excl = excel_roundup(
        (950.0 + pub_excl) if raw_succes < 950.0 else (raw_succes + pub_excl),
        2,
    )
    succes_plus_pub_btw = excel_roundup(succes_plus_pub_excl * 0.21, 2)
    succes_plus_pub_incl = round(succes_plus_pub_excl + succes_plus_pub_btw, 2)

    monthly_admin_excl = excel_roundup(principal * (admin_fee_pct / 100.0), 2)
    monthly_admin_btw = excel_roundup(monthly_admin_excl * 0.21, 2)
    monthly_admin_incl = round(monthly_admin_excl + monthly_admin_btw, 2)

    total_admin_excl = round(monthly_admin_excl * months, 2)
    total_admin_btw = round(total_admin_excl * 0.21, 2)
    total_admin_incl = round(total_admin_excl + total_admin_btw, 2)

    return {
        "publicatiekosten_excl": pub_excl,
        "succes_plus_pub_excl": succes_plus_pub_excl,
        "succes_plus_pub_btw": succes_plus_pub_btw,
        "succes_plus_pub_incl": succes_plus_pub_incl,
        "monthly_admin_excl": monthly_admin_excl,
        "monthly_admin_btw": monthly_admin_btw,
        "monthly_admin_incl": monthly_admin_incl,
        "total_admin_excl": total_admin_excl,
        "total_admin_btw": total_admin_btw,
        "total_admin_incl": total_admin_incl,
    }


def generate_amortization_schedule(
    principal: float,
    annual_rate_pct: float,
    months: int,
    grace_months: int = 0,
    balloon_amount: float = 0.0,
    succes_fee_pct: float = 2.0,
    admin_fee_pct: float = 0.07,
    loan_type: str = "Lineair",
) -> list[dict]:
    """Generate the EBI amortization schedule using the exact formulas from 'Amortisatie EBI' in Amo EBI.xlsm."""
    rows = [
        {
            "termijn": "0",
            "termijnbedrag": "",
            "terugbetaling": "",
            "rente_en_kosten": "",
            "saldo": format_euro(principal, 2),
        }
    ]
    if months <= 0 or principal <= 0:
        return rows

    # Monthly admin fee from Zekerheden!B78 & Amortisatie EBI col J & M:
    # J = ROUNDUP(Hoofdsom * 0.0007, 2), M = ROUNDUP(J * 0.21, 2)
    admin_excl = excel_roundup(principal * (admin_fee_pct / 100.0), 2)
    admin_btw = excel_roundup(admin_excl * 0.21, 2)
    monthly_admin_incl = round(admin_excl + admin_btw, 2)

    rate_decimal = annual_rate_pct / 100.0
    saldo = round(principal, 2)

    if loan_type.lower().startswith("annu"):
        # Annuitair EBI (sheet7 & sheet8 in Amo EBI.xlsm)
        n_repay = (months - grace_months) if balloon_amount <= 0 else (months - grace_months - 1)
        n_repay = max(1, n_repay)
        r_m = rate_decimal / 12.0
        if r_m > 0:
            # Excel PMT(rate, nper, -pv, fv)
            factor = (1 + r_m) ** n_repay
            pmt = (principal * factor - balloon_amount) * r_m / (factor - 1)
        else:
            pmt = (principal - balloon_amount) / n_repay

        for m in range(1, months + 1):
            exact_interest = (saldo / 12.0) * rate_decimal
            rente_roundup = excel_roundup(exact_interest, 2)
            if m == months and balloon_amount > 0:
                aflossing = round(balloon_amount, 2)
            elif m == months:
                aflossing = round(saldo, 2)
            elif m <= grace_months:
                aflossing = 0.0
            else:
                aflossing = max(0.0, round(pmt - exact_interest, 2))
            rente_kosten = round(rente_roundup + monthly_admin_incl, 2)
            termijnbedrag = round(aflossing + rente_kosten, 2)
            saldo = max(0.0, round(saldo - aflossing, 2))
            rows.append(
                {
                    "termijn": str(m),
                    "termijnbedrag": format_euro(termijnbedrag, 2),
                    "terugbetaling": "€ -" if aflossing == 0 else format_euro(aflossing, 2),
                    "rente_en_kosten": format_euro(rente_kosten, 2),
                    "saldo": "€ -" if saldo == 0 else format_euro(saldo, 2),
                }
            )
        return rows

    # Lineair EBI ('Amortisatie EBI' & 'Zekerheden!B64' in Amo EBI.xlsm):
    # B64 = IF(B49>0, (B46-B49)/(B47-1-B48)*1000, (B46-B49)/(B47-B48)*1000)
    repay_months = (months - 1 - grace_months) if balloon_amount > 0 else (months - grace_months)
    repay_months = max(1, repay_months)
    repayable_principal = max(0.0, principal - balloon_amount)
    raw_aflossing = repayable_principal / repay_months
    std_repayment = excel_roundup(raw_aflossing, 2)

    # Penultimate repayment month adjustment so final balance equals 0.00
    total_std = round(std_repayment * repay_months, 2)
    diff = round(repayable_principal - total_std, 2)
    adjust_month = (months - 2) if balloon_amount > 0 else (months - 1)

    for m in range(1, months + 1):
        # H_m = ROUNDUP((O_{m-1} / 12 * Zekerheden!$B$50), 2)
        rente = excel_roundup((saldo / 12.0) * rate_decimal, 2)
        rente_kosten = round(rente + monthly_admin_incl, 2)

        if m == months and balloon_amount > 0:
            aflossing = excel_roundup(balloon_amount, 2)
        elif m == months and grace_months >= months:
            aflossing = round(saldo, 2)
        elif m <= grace_months:
            aflossing = 0.0
        elif m == adjust_month and abs(diff) > 0.001 and repay_months > 1:
            aflossing = round(std_repayment + diff, 2)
        else:
            aflossing = std_repayment

        termijnbedrag = round(aflossing + rente_kosten, 2)
        saldo = max(0.0, round(saldo - aflossing, 2))

        rows.append(
            {
                "termijn": str(m),
                "termijnbedrag": format_euro(termijnbedrag, 2),
                "terugbetaling": "€ -" if aflossing == 0 else format_euro(aflossing, 2),
                "rente_en_kosten": format_euro(rente_kosten, 2),
                "saldo": "€ -" if saldo == 0 else format_euro(saldo, 2),
            }
        )
    return rows


def parse_pitch_html(html: str, url: str = "") -> dict:
    """Extract all EBI fields from a Collin Crowdfund pitch HTML page."""
    # 1. Loan ID / Leningnummer
    loan_id_match = (
        re.search(r'data-loan-id=["\'](\d+)["\']', html)
        or re.search(r'<div[^>]*id=["\']loanid["\'][^>]*>(\d+)</div>', html)
        or re.search(r'LoanId=(\d+)', html)
    )
    raw_loan_id = loan_id_match.group(1) if loan_id_match else ""
    leningnummer = raw_loan_id.zfill(8) if raw_loan_id else ""

    # 2. Project Title
    title_match = re.search(r'<h1[^>]*class=["\'][^"\']*entry-title[^"\']*["\'][^>]*>(.*?)</h1>', html, re.S | re.I)
    if not title_match:
        title_match = re.search(r"<title>(.*?)</title>", html, re.S | re.I)
    project_title = clean_text(title_match.group(1)) if title_match else ""
    project_title = re.sub(r"\s*-\s*Collin Crowdfund.*$", "", project_title, flags=re.I).strip()

    # 3. Extract all h2 sections inside appmx-summary or main content
    sections = {}
    for m in re.finditer(r"<h2[^>]*>(.*?)</h2>(.*?)(?=<h2\b|<div class=\"et_pb_column|<footer|\Z)", html, re.S | re.I):
        sec_title = clean_text(m.group(1)).strip()
        sec_body = m.group(2)
        if sec_title and sec_title not in sections:
            sections[sec_title] = sec_body

    # 4. Parse Bedrijfskenmerken table
    bedrijfskenmerken = {}
    if "Bedrijfskenmerken" in sections:
        tbl_m = re.search(r"<table[^>]*>(.*?)</table>", sections["Bedrijfskenmerken"], re.S | re.I)
        if tbl_m:
            for row in parse_html_table(tbl_m.group(1)):
                if len(row) >= 2:
                    bedrijfskenmerken[row[0]["text"].rstrip(":")] = row[1]["text"]

    handelsnaam = bedrijfskenmerken.get("Handelsnaam") or project_title
    rechtsvorm = bedrijfskenmerken.get("Rechtsvorm", "")
    website = bedrijfskenmerken.get("Website", "n.v.t.")
    vestigingsplaats = bedrijfskenmerken.get("Vestigingsplaats", "")
    kvk_from_table = bedrijfskenmerken.get("KvK-nummer") or bedrijfskenmerken.get("KvK nummer") or ""

    # 5. Parse Financieringsbehoefte / Investeringsoverzicht table
    investeringsoverzicht = []
    streefbedrag_str = ""
    principal_amount = 0.0

    for sec_name in ["Financieringsbehoefte", "Investeringsoverzicht"]:
        if sec_name in sections:
            tbl_m = re.search(r"<table[^>]*>(.*?)</table>", sections[sec_name], re.S | re.I)
            if tbl_m:
                investeringsoverzicht = parse_html_table(tbl_m.group(1))
                break

    for row in investeringsoverzicht:
        if len(row) >= 2 and ("collin" in row[0]["text"].lower() or "lening" in row[0]["text"].lower()):
            streefbedrag_str = row[1]["text"]
            nums = re.sub(r"[^\d]", "", streefbedrag_str)
            if nums:
                principal_amount = float(nums)

    # Fallback for principal amount from top bar
    if principal_amount <= 0:
        top_m = re.search(r"€\s*([\d.]+)\s*\|\s*<span[^>]*>([\d,]+)\s*(?:&nbsp;|\s)*%\s*rente</span>\s*\|\s*(\d+)\s*maanden", html, re.I)
        if top_m:
            principal_amount = float(top_m.group(1).replace(".", ""))
            streefbedrag_str = format_euro(principal_amount, 0)

    # 6. Parse Samenvatting, Leendoel, Ondernemer, Onderneming & Financieringsbehoefte text
    onderneming_paragraphs = []
    financieringsbehoefte_paragraphs = []

    def _is_loan_terms_para(text_str: str) -> bool:
        low = text_str.lower()
        return (
            ("lening" in low and ("maanden" in low or "rente" in low))
            or ("looptijd" in low and ("maanden" in low or "rente" in low))
            or ("financiering via collin" in low and ("maanden" in low or "rente" in low))
            or ("gevraagde financiering via collin" in low)
        )

    # 6a. Check 'Samenvatting' section first — standard pitches always have
    #     <strong>Financieringsbehoefte</strong> at the bottom of Samenvatting.
    samenvatting_html = sections.get("Samenvatting", "")
    if samenvatting_html:
        fb_sam_m = re.search(
            r"<strong>\s*Financieringsbehoefte\s*:?\s*</strong>\s*(?:<br\s*/?>)?\s*(.*?)(?=<strong>\s*Risico|</p\s*>|<ul\b|\Z)",
            samenvatting_html,
            re.S | re.I,
        )
        if fb_sam_m:
            fb_sam_text = clean_text(fb_sam_m.group(1)).strip()
            if fb_sam_text:
                financieringsbehoefte_paragraphs.append(fb_sam_text)

    # 6b. Extract Ondernemer + Onderneming (or Leendoel on Collin Direct pitches)
    for sec_key in ["Ondernemer", "Onderneming"]:
        if sec_key in sections:
            for pm in re.finditer(r"<p[^>]*>(.*?)</p\s*>", sections[sec_key], re.S | re.I):
                t = clean_text(pm.group(1))
                if t and not t.lower().startswith("de totale investering bedraagt"):
                    onderneming_paragraphs.append(t)

    if "Leendoel" in sections:
        leendoel_ps = []
        for pm in re.finditer(r"<p[^>]*>(.*?)</p\s*>", sections["Leendoel"], re.S | re.I):
            t = clean_text(pm.group(1))
            if t:
                leendoel_ps.append(t)
        has_dedicated_onderneming = bool(onderneming_paragraphs)
        for p in leendoel_ps:
            low = p.lower()
            if _is_loan_terms_para(p):
                if not financieringsbehoefte_paragraphs:
                    financieringsbehoefte_paragraphs.append(p)
            elif low.startswith("de totale investering bedraagt") or low.startswith("de totale financieringsbehoefte"):
                # Summary sentence already represented by the Investeringsoverzicht table
                continue
            elif not has_dedicated_onderneming:
                onderneming_paragraphs.append(p)

    # 6c. Also inspect 'Financieringsbehoefte' section for loan terms if not yet found
    fb_section_text = ""
    if "Financieringsbehoefte" in sections:
        fb_sec_ps = []
        for pm in re.finditer(r"<p[^>]*>(.*?)</p\s*>", sections["Financieringsbehoefte"], re.S | re.I):
            t = clean_text(pm.group(1)).strip()
            if not t:
                continue
            if t.lower().rstrip(":") in ("investeringsoverzicht", "overige financiers", "financieringsbehoefte"):
                continue
            fb_sec_ps.append(t)
        fb_section_text = "\n\n".join(fb_sec_ps)
        if not financieringsbehoefte_paragraphs and fb_sec_ps:
            financieringsbehoefte_paragraphs.extend(fb_sec_ps)

    onderneming_text = "\n\n".join(onderneming_paragraphs)
    financieringsbehoefte_text = "\n\n".join(financieringsbehoefte_paragraphs)

    # Combine all candidate sources for extracting exact loan parameters
    loan_parse_source = "\n".join(filter(None, [financieringsbehoefte_text, fb_section_text]))

    # Extract interest rate, duration, grace months, balloon from financieringsbehoefte_text / Samenvatting / top bar
    annual_rate = 7.5
    duration_months = 60
    grace_months = 0
    balloon_amount = 0.0

    top_bar = re.search(
        r"€\s*([\d.]+)\s*\|\s*<span[^>]*>([\d,]+)\s*(?:&nbsp;|\s)*%\s*rente</span>\s*\|\s*(\d+)\s*maanden",
        html,
        re.I,
    )
    if top_bar:
        annual_rate = float(top_bar.group(2).replace(",", "."))
        duration_months = int(top_bar.group(3))

    dutch_num_words = {
        "een": 1, "één": 1, "twee": 2, "drie": 3, "vier": 4, "vijf": 5,
        "zes": 6, "zeven": 7, "acht": 8, "negen": 9, "tien": 10,
        "elf": 11, "twaalf": 12, "achttien": 18, "vierentwintig": 24, "zesendertig": 36,
    }

    def _parse_int_or_word(val_str: str) -> int:
        v = (val_str or "").strip().lower()
        if v.isdigit():
            return int(v)
        return dutch_num_words.get(v, 0)

    if loan_parse_source:
        m_rate = (
            re.search(r"rente\s*(?:bedraagt|:)?\s*(?:gedurende[^0-9]*?)?([\d,]+)\s*%", loan_parse_source, re.I)
            or re.search(r"([\d,]+)\s*%\s*(?:per jaar|rente)", loan_parse_source, re.I)
        )
        if m_rate:
            annual_rate = float(m_rate.group(1).replace(",", "."))

        m_dur = (
            re.search(r"looptijd(?:\s+van|\s+is|\s+bedraagt|:)?\s*(\d+)\s*maanden", loan_parse_source, re.I)
            or re.search(r"aflossing\s+in\s*(\d+)\s*maanden", loan_parse_source, re.I)
        )
        if m_dur:
            duration_months = int(m_dur.group(1))

        m_grace = (
            re.search(r"eerste\s+(\d+|[a-zà-ÿ]+)\s+maanden\s+aflossingsvrij", loan_parse_source, re.I)
            or re.search(r"aflossingsvrije\s+periode\s+van\s+(\d+|[a-zà-ÿ]+)\s+maanden", loan_parse_source, re.I)
            or re.search(r"(\d+|[a-zà-ÿ]+)\s+maanden\s+aflossingsvrij", loan_parse_source, re.I)
        )
        if m_grace:
            grace_months = _parse_int_or_word(m_grace.group(1))

        m_slot = re.search(
            r"slottermijn(?:\s+van|\s+bedraagt|\s+ter\s+hoogte\s+van)?\s*€\s*([\d.]+)",
            loan_parse_source,
            re.I,
        )
        if m_slot:
            balloon_amount = float(m_slot.group(1).replace(".", ""))
        elif re.search(
            r"volledig\s+aflossingsvrij|aflossingsvrije\s+lening|ineens\s+aan\s+het\s+einde\s+van\s+de\s+looptijd\s+afgelost|maanden,\s*aflossingsvrij",
            loan_parse_source,
            re.I,
        ):
            balloon_amount = principal_amount
            grace_months = 0
        elif grace_months >= duration_months and duration_months > 0:
            balloon_amount = principal_amount
            grace_months = 0

    # 7. Parse Risico / Risico's section: Zekerheden en voorwaarden, D&B, Collin Credit Score, Toelichting jaarcijfers
    zekerheden_bullets = []
    db_score = "Niet van toepassing"
    cc_score = "Ruim voldoende"
    presentatie_intro = ""
    kasstroom_paragraphs = []

    db_div = re.search(r'<div[^>]*id=["\']DBRisk["\'][^>]*>(.*?)</div>', html, re.S | re.I)
    if db_div:
        db_score = clean_text(db_div.group(1)).strip()

    cc_div = re.search(r'<div[^>]*id=["\']CCScore["\'][^>]*>(.*?)</div>', html, re.S | re.I)
    if cc_div:
        cc_score = clean_text(cc_div.group(1)).strip()

    risico_html = ""
    for k, v in sections.items():
        if k.lower().startswith("risico"):
            risico_html = v
            break

    if risico_html:
        # Extract top-level <li> bullets under Zekerheden en voorwaarden
        zek_m = re.search(r"<strong>\s*Zekerheden[^<]*</strong>\s*(.*)", risico_html, re.S | re.I)
        ul_source = zek_m.group(1) if zek_m else risico_html
        ul_m = re.search(r"<ul[^>]*>(.*)", ul_source, re.S | re.I)
        if ul_m:
            for li_m in re.finditer(r"<li[^>]*>(.*?)(?=<li\b|</ul\s*>)", ul_m.group(1), re.S | re.I):
                bt = clean_text(li_m.group(1))
                if bt:
                    zekerheden_bullets.append(bt)

        # Extract Collin Credit Score basis sentence
        ccs_m = re.search(
            r"<strong>\s*Collin Credit Score\s*</strong>\s*(?:<br\s*/?>)?\s*(.*?)</p>",
            risico_html,
            re.S | re.I,
        )
        if ccs_m:
            ccs_para = clean_text(ccs_m.group(1))
            basis_m = re.search(
                r"(?:Er is bij de beoordeling gebruik gemaakt van|De (?:beoordeling|Score|Collin Credit Score) is gebaseerd op)\s+(.*?)(?:\.\s*De overall Collin Credit Score|$)",
                ccs_para,
                re.S | re.I,
            )
            if basis_m:
                presentatie_intro = (
                    "De presentatie van de belangrijkste financiële cijfers is gebaseerd op "
                    + basis_m.group(1).strip().rstrip(".")
                    + "."
                )

        # Extract Toelichting jaarcijfers / Afloscapaciteit paragraphs
        toel_m = re.search(
            r"<strong>\s*(?:Toelichting jaarcijfers|Afloscapaciteit|Kasstroom)\s*</strong>\s*(?:<br\s*/?>)?\s*(.*?)(?=<strong>\s*(?:Zekerheden|Solvabiliteit)|\Z)",
            risico_html,
            re.S | re.I,
        )
        if toel_m:
            rest_html = toel_m.group(1)
            first_p_end = rest_html.find("</p>")
            if first_p_end != -1:
                p1 = clean_text(rest_html[:first_p_end])
                if p1:
                    kasstroom_paragraphs.append(p1)
                for pm in re.finditer(r"<p[^>]*>(.*?)</p\s*>", rest_html[first_p_end + 4 :], re.S | re.I):
                    pt = clean_text(pm.group(1))
                    if pt:
                        kasstroom_paragraphs.append(pt)
            else:
                p1 = clean_text(rest_html)
                if p1:
                    kasstroom_paragraphs.append(p1)

    # Check if there is a separate Afloscapaciteit section
    if not kasstroom_paragraphs and "Afloscapaciteit" in sections:
        for pm in re.finditer(r"<p[^>]*>(.*?)</p\s*>", sections["Afloscapaciteit"], re.S | re.I):
            pt = clean_text(pm.group(1))
            if pt:
                kasstroom_paragraphs.append(pt)

    kasstroom_text = "\n\n".join(kasstroom_paragraphs)

    # 8. Extract Geldnemers and Vestigingsplaats from Zekerheden or Title
    geldnemers = []
    for b in zekerheden_bullets:
        m_gn = re.search(
            r"De geldnemers?\s+(?:zijn|is)\s+(.*?)(?<!\b[A-Z])\.(?:\s+[A-Z][a-z]|$)",
            b,
        )
        if m_gn:
            raw_names = m_gn.group(1).strip()
            parts = [p.strip() for p in re.split(r"\s+en\s+|,\s*", raw_names) if p.strip()]
            kvk_matches = re.findall(r"\b(\d{8})\b", b)
            for idx, p in enumerate(parts):
                p_clean = re.sub(r"([A-Za-z])(B\.V\.)", r"\1 \2", p)
                p_clean = re.sub(r"\s*\(in priv[eé]\)", "", p_clean, flags=re.I).strip()
                kvk_val = kvk_matches[idx] if idx < len(kvk_matches) else (kvk_from_table or "n.v.t.")
                geldnemers.append({"naam": p_clean, "kvk": kvk_val, "dnb_url": ""})
            break

    if not geldnemers:
        parts = [p.strip() for p in re.split(r"\s+en\s+", handelsnaam) if p.strip()]
        for p in parts:
            geldnemers.append({"naam": p, "kvk": kvk_from_table or "n.v.t.", "dnb_url": ""})

    if not vestigingsplaats:
        for b in zekerheden_bullets:
            # Match Dutch postcode + city, or 'te <City>'
            m_pc = re.search(r"\b\d{4}\s*[A-Z]{2}\s+(?:te\s+)?([A-ZÀ-ÿIJ][a-zà-ÿijA-Z\s'-]+?)(?:,|\.|ten behoeve)", b)
            if m_pc:
                vestigingsplaats = m_pc.group(1).strip()
                break
            m_te = re.search(r"\bte\s+([A-ZÀ-ÿIJ][a-zà-ÿijA-Z'-]+)(?:,|\.)", b)
            if m_te:
                vestigingsplaats = m_te.group(1).strip()
                break

    # 8b. Automatic KvK lookup via KVK Handelsregister API for business entities
    is_box3 = rechtsvorm.strip().lower() == "box 3"
    for gn in geldnemers:
        gn_name = gn.get("naam", "")
        has_corp_suffix = bool(re.search(r"\b(B\.?V\.?|V\.?O\.?F\.?|N\.?V\.?|C\.?V\.?|Holding|Beheer|Stichting)\b", gn_name, re.I))
        if gn.get("kvk") in ("", "n.v.t.", None) and (not is_box3 or has_corp_suffix):
            hits = lookup_kvk(gn_name, vestigingsplaats)
            if hits:
                gn["kvk"] = hits[0]["kvk"]
                if not vestigingsplaats and hits[0].get("plaats"):
                    vestigingsplaats = hits[0]["plaats"]

    # 8c. Build default Organigram structure (single or multi-layer)
    organigram_nodes = []
    struct_text = clean_text(sections.get("Structuur", "")) + "\n" + onderneming_text + "\n" + "\n".join(zekerheden_bullets)
    gn_names_lower = {g["naam"].lower() for g in geldnemers}

    # Check if explicit parent-subsidiary / holding relationship is described for the borrowers
    if len(geldnemers) >= 2 and any("holding" in g["naam"].lower() or "beheer" in g["naam"].lower() for g in geldnemers):
        top_nodes = [g for g in geldnemers if "holding" in g["naam"].lower() or "beheer" in g["naam"].lower()]
        bottom_nodes = [g for g in geldnemers if g not in top_nodes]
        if top_nodes and bottom_nodes:
            pct = "100%" if len(top_nodes) == 1 else f"{int(100 / len(top_nodes))}%"
            for t in top_nodes:
                organigram_nodes.append({"naam": t["naam"], "laag": 1, "is_geldnemer": True, "label": ""})
            for b_node in bottom_nodes:
                organigram_nodes.append({"naam": b_node["naam"], "laag": 2, "is_geldnemer": True, "label": pct})

    if not organigram_nodes:
        # Check if a single borrower is 100% owned by a holding mentioned in the text
        m_hold = re.search(
            r"aandelen\s+van\s+([A-Za-z0-9\s.&'-]+?B\.V\.)\s+worden\s+(?:voor\s+(\d+%)\s+)?gehouden\s+door\s+([A-Za-z0-9\s.&'-]+?B\.V\.)",
            struct_text,
            re.I,
        )
        if m_hold and not is_box3:
            child_bv = m_hold.group(1).strip()
            pct = m_hold.group(2) or "100%"
            parent_bv = m_hold.group(3).strip()
            organigram_nodes = [
                {
                    "naam": parent_bv,
                    "laag": 1,
                    "is_geldnemer": parent_bv.lower() in gn_names_lower,
                    "label": "",
                },
                {
                    "naam": child_bv,
                    "laag": 2,
                    "is_geldnemer": True,
                    "label": pct,
                },
            ]
        else:
            organigram_nodes = [
                {
                    "naam": handelsnaam,
                    "laag": 1,
                    "is_geldnemer": True,
                    "label": "",
                }
            ]

    # 9. Parse Jaarcijfers tables (with their preceding caption paragraphs)
    jaarcijfers_tables = []
    if "Jaarcijfers" in sections:
        jc_html = sections["Jaarcijfers"]
        # Match pairs of (<p>caption</p>)* followed by <table>...</table>
        pos = 0
        for tbl_m in re.finditer(r"<table[^>]*>(.*?)</table\s*>", jc_html, re.S | re.I):
            preceding = jc_html[pos : tbl_m.start()]
            pos = tbl_m.end()
            # Extract title and subtitle from preceding HTML
            title_lines = []
            subtitle_lines = []
            for str_m in re.finditer(r"<strong[^>]*>(.*?)</strong\s*>", preceding, re.S | re.I):
                t = clean_text(str_m.group(1)).replace("\n", " ").strip()
                if t:
                    title_lines.append(t)
            for em_m in re.finditer(r"<em[^>]*>(.*?)</em\s*>", preceding, re.S | re.I):
                t = clean_text(em_m.group(1)).replace("\n", " ").strip()
                if t:
                    subtitle_lines.append(t)
            if not title_lines:
                raw_prec = clean_text(preceding)
                if raw_prec:
                    title_lines.append(raw_prec.splitlines()[0])

            parsed_rows = parse_html_table(tbl_m.group(1))
            # By default in EBI, include non-BV-historical tables or first 3 tables (user can toggle in UI)
            full_title = " ".join(title_lines)
            include_default = not full_title.upper().startswith("BV:")
            jaarcijfers_tables.append(
                {
                    "title": full_title,
                    "subtitle": " ".join(subtitle_lines) or "In duizendtallen (x 1.000)",
                    "rows": parsed_rows,
                    "included": include_default,
                }
            )
        # Extract any D&B report links in Jaarcijfers
        for a_m in re.finditer(r'<a[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', jc_html, re.S | re.I):
            href = a_m.group(1).strip()
            link_txt = clean_text(a_m.group(2)).strip()
            if href and link_txt:
                for gn in geldnemers:
                    if gn.get("naam", "").lower() in link_txt.lower() or link_txt.lower() in gn.get("naam", "").lower():
                        gn["dnb_url"] = href
                        break

    # 10. Format Risico(scores) en voorwaarden intro line
    if db_score.lower() in ("niet van toepassing", "n.v.t.", "nvt", ""):
        db_sentence = "De kwalificatie van Dun & Bradstreet voor deze aanvraag is niet van toepassing."
    else:
        db_sentence = f"De kwalificatie van Dun & Bradstreet voor deze aanvraag is {db_score}."
    cc_sentence = f"De overall Collin Credit Score is {cc_score}."

    # 11. Historie betaalverplichtingen source (BKR for Box 3 / private individuals, D&B for companies)
    is_private = rechtsvorm.lower() == "box 3" or db_score.lower() == "niet van toepassing"
    historie_instantie = "BKR" if is_private else "Dun & Bradstreet"
    historie_text = (
        f"Heeft de Projecteigenaar de laatste vijf jaar een betalingsachterstand van 90 dagen of meer gehad? "
        f"Conform recent inzicht vanuit {historie_instantie} zijn deze niet bekend."
    )

    loan_type = "Annuïtair" if "annu" in (financieringsbehoefte_text or "").lower() else "Lineair"

    # 12. Generate default Betaalschema using exact Amo EBI.xlsm formulas
    betaalschema = generate_amortization_schedule(
        principal=principal_amount,
        annual_rate_pct=annual_rate,
        months=duration_months,
        grace_months=grace_months,
        balloon_amount=balloon_amount,
        succes_fee_pct=2.0,
        admin_fee_pct=0.07,
        loan_type=loan_type,
    )

    return {
        "source_url": url,
        "lei_code": "724500VFNB9F43XDI745",
        "leningnummer": leningnummer,
        "investeringskans": project_title,
        "projecteigenaar": handelsnaam,
        "projectnaam": project_title,
        "streefbedrag": streefbedrag_str or format_euro(principal_amount, 0),
        "deadline": "30 dagen na publicatiedatum",
        "rechtsvorm": rechtsvorm,
        "geldnemers": geldnemers,
        "website": website,
        "vestigingsplaats": vestigingsplaats,
        "organigram_text": handelsnaam,
        "organigram_nodes": organigram_nodes,
        "onderneming_text": onderneming_text,
        "presentatie_intro": presentatie_intro,
        "kasstroom_text": kasstroom_text,
        "jaarcijfers_tables": jaarcijfers_tables,
        "kvk_jaarcijfers_text": "Niet van toepassing.",
        "investeringsoverzicht": investeringsoverzicht,
        "bodembedrag": "n.v.t.",
        "inschrijfperiode": "Maximaal 30 dagen na publicatie",
        "bedenktermijn": "Vier kalenderdagen",
        "collin_invested_note": (
            "In deze investeringskans heeft/hebben de aandeelhouder(s), bestuurder(s) "
            "en/of werknemer(s) van Collin reeds geïnvesteerd."
        ),
        "kasstroom_risico_afsluiting": (
            "De aangeleverde en beoordeelde (cijfer)informatie is in belangrijke mate "
            "gebaseerd op historische, en toekomstige (cijfer)informatie. Het risico bestaat "
            "dat de toekomstige kasstroom niet toereikend is om aan de betalingsverplichting(en) "
            "te kunnen voldoen."
        ),
        "include_herfinanciering": balloon_amount > 0,
        "herfinanciering_text": (
            "Bij deze lening is er sprake van restschuld op het einde van de looptijd. "
            "Deze restschuld dient de Projecteigenaar af te lossen. Dit kan bijvoorbeeld door "
            "een herfinanciering via Collin of elders, verkoop van het onroerend goed of "
            "beschikbare liquiditeiten. Er is een risico dat de restschuld niet tijdig wordt "
            "afgelost. Indien er sprake is van een lening met een restschuld is Projecteigenaar "
            "van mening dat deze restschuld verantwoord is op basis van de huidige inzichten."
        ),
        "financieringsbehoefte_text": financieringsbehoefte_text,
        "risico_scores_text": f"{db_sentence}\n{cc_sentence}",
        "zekerheden_bullets": zekerheden_bullets,
        "loan_params": {
            "principal": principal_amount,
            "annual_rate_pct": annual_rate,
            "months": duration_months,
            "grace_months": grace_months,
            "balloon_amount": balloon_amount,
            "succes_fee_pct": 2.0,
            "admin_fee_pct": 0.07,
            "loan_type": loan_type,
        },
        "betaalschema": betaalschema,
        "historie_betaalverplichtingen": historie_text,
        "beheervergoeding": "0,1%",
    }


def parse_pitch_docx(docx_bytes: bytes, filename: str = "") -> dict:
    """Extract all EBI fields from a Collin Crowdfund Word pitch document (.docx)."""
    import base64
    import io
    import xml.etree.ElementTree as ET
    import zipfile

    ns = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    r_embed_attr = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"

    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
        doc_xml = zf.read("word/document.xml")
        rels_map = {}
        if "word/_rels/document.xml.rels" in zf.namelist():
            rels_root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
            for rel in rels_root:
                rid = rel.attrib.get("Id", "")
                target = rel.attrib.get("Target", "")
                if rid and target:
                    rels_map[rid] = "word/" + target.lstrip("/") if not target.startswith("word/") else target

        def _get_image_bytes(rid: str) -> bytes | None:
            path = rels_map.get(rid, "")
            if path and path in zf.namelist():
                return zf.read(path)
            return None

        root = ET.fromstring(doc_xml)
        body = root.find("w:body", ns)

        def _parse_p(p_elem) -> dict:
            is_list = p_elem.find(".//w:numPr", ns) is not None
            pstyle_el = p_elem.find("w:pPr/w:pStyle", ns)
            pstyle = pstyle_el.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", "") if pstyle_el is not None else ""
            if "list" in pstyle.lower() or "opsomming" in pstyle.lower():
                is_list = True

            full_parts = []
            bold_chars = 0
            total_chars = 0
            for r in p_elem.findall("w:r", ns):
                txt = "".join(t.text for t in r.findall(".//w:t", ns) if t.text)
                if not txt:
                    continue
                is_b = r.find("w:rPr/w:b", ns) is not None
                b_val = r.find("w:rPr/w:b", ns)
                if b_val is not None and b_val.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") in ("0", "false"):
                    is_b = False
                full_parts.append(txt)
                stripped_len = len(txt.strip())
                total_chars += stripped_len
                if is_b:
                    bold_chars += stripped_len

            text_str = "".join(full_parts).replace("\xa0", " ").strip()
            text_str = re.sub(r"[ \t]+", " ", text_str)
            is_bold = (total_chars > 0 and bold_chars >= total_chars * 0.7) or pstyle.lower().startswith(("heading", "kop"))
            blips = [b.attrib.get(r_embed_attr) for b in p_elem.findall(".//a:blip", ns) if b.attrib.get(r_embed_attr)]
            return {
                "type": "p",
                "text": text_str,
                "bold": is_bold,
                "is_list": is_list,
                "blips": blips,
            }

        def _parse_docx_table(tbl_elem) -> list[list[dict]]:
            rows = []
            for r_idx, tr in enumerate(tbl_elem.findall("w:tr", ns)):
                row_cells = []
                for tc in tr.findall("w:tc", ns):
                    cell_texts = []
                    cell_bold = False
                    for p in tc.findall("w:p", ns):
                        p_info = _parse_p(p)
                        if p_info["text"]:
                            cell_texts.append(p_info["text"])
                            if p_info["bold"]:
                                cell_bold = True
                    c_str = " ".join(cell_texts).strip()
                    is_num = bool(re.match(r"^(?:€\s*)?[-+]?[\d.,]+(?:\s*,-|\s*%)?$", c_str))
                    row_cells.append({
                        "text": c_str,
                        "bold": cell_bold or (r_idx == 0 and bool(c_str)) or c_str.lower().startswith("totaal"),
                        "align": "right" if is_num else "left",
                    })
                if row_cells and any(c["text"] for c in row_cells):
                    rows.append(row_cells)
            return rows

        header_kv = {}
        sections_blocks: dict[str, list[dict]] = {}

        if body is not None:
            for child in body:
                tag = child.tag.split("}")[-1]
                if tag == "tbl":
                    trs = child.findall("w:tr", ns)
                    if not trs:
                        continue
                    first_row_text = " ".join(
                        "".join(t.text for t in tc.findall(".//w:t", ns) if t.text).strip()
                        for tc in trs[0].findall("w:tc", ns)
                    ).strip()
                    if "header" in first_row_text.lower() or "gegevens bovenaan" in first_row_text.lower() or len(trs) >= 6:
                        for tr in trs[1:]:
                            tcs = tr.findall("w:tc", ns)
                            if len(tcs) >= 2:
                                k = "".join(t.text for t in tcs[0].findall(".//w:t", ns) if t.text).replace("\xa0", " ").strip()
                                v = "".join(t.text for t in tcs[1].findall(".//w:t", ns) if t.text).replace("\xa0", " ").strip()
                                if k:
                                    header_kv[k.lower()] = v
                    elif len(trs) >= 2:
                        sec_title = first_row_text.strip()
                        blocks = []
                        for tr in trs[1:]:
                            for tc in tr.findall("w:tc", ns):
                                for elem in tc:
                                    etag = elem.tag.split("}")[-1]
                                    if etag == "p":
                                        pi = _parse_p(elem)
                                        if pi["text"] or pi["blips"]:
                                            blocks.append(pi)
                                    elif etag == "tbl":
                                        blocks.append({"type": "tbl", "rows": _parse_docx_table(elem)})
                        sections_blocks[sec_title] = blocks

        # Extract Header values
        handelsnaam = ""
        leningnummer = ""
        doel_str = ""
        principal_amount = 100000.0
        streefbedrag_str = ""
        annual_rate = 7.5
        duration_months = 60
        db_score = "Niet van toepassing"
        cc_score = "Goed"

        for k, v in header_kv.items():
            if "handelsnaam" in k or "geldnemer" in k:
                handelsnaam = v
            elif "leningnummer" in k:
                leningnummer = v.strip()
            elif "doel" in k:
                doel_str = v
            elif "leenbedrag" in k:
                num_s = re.sub(r"[^\d]", "", v.split(",")[0])
                if num_s:
                    principal_amount = float(num_s)
                    streefbedrag_str = format_euro(principal_amount, 0)
            elif "rente" in k:
                m_r = re.search(r"([\d,]+)", v)
                if m_r:
                    annual_rate = float(m_r.group(1).replace(",", "."))
            elif "looptijd" in k:
                m_d = re.search(r"(\d+)", v)
                if m_d:
                    duration_months = int(m_d.group(1))
            elif "d&b" in k or "dun" in k:
                if v:
                    db_score = v
            elif "collin credit score" in k:
                if v:
                    cc_score = v

        project_title = handelsnaam or os.path.splitext(os.path.basename(filename or "Pitch"))[0]

        def _get_sec(prefix: str) -> list[dict]:
            for sk, sv in sections_blocks.items():
                if sk.lower().strip().startswith(prefix.lower()):
                    return sv
            return []

        # 1. Samenvatting -> Financieringsbehoefte & fallback Risico bullets
        financieringsbehoefte_paragraphs = []
        samenvatting_zekerheden = []
        sam_blocks = _get_sec("samenvatting")
        sam_mode = "intro"
        for b in sam_blocks:
            if b["type"] != "p":
                continue
            txt = b["text"]
            low = txt.lower().rstrip(":")
            if b["bold"] and low == "financieringsbehoefte":
                sam_mode = "fb"
                continue
            if b["bold"] and low.startswith("risico"):
                sam_mode = "risico"
                continue
            if sam_mode == "fb" and txt:
                financieringsbehoefte_paragraphs.append(txt)
            elif sam_mode == "risico" and b["is_list"] and txt:
                samenvatting_zekerheden.append(txt)

        # 2. Onderneming + Organigram image + Onderneming tables
        onderneming_paragraphs = []
        onderneming_tables = []
        organigram_image_b64 = ""

        ond_blocks = _get_sec("onderneming")
        ond_mode = "normal"
        pending_tbl_title = ""
        for b in ond_blocks:
            if b["type"] == "p":
                for rid in b.get("blips", []):
                    img_b = _get_image_bytes(rid)
                    if img_b and not organigram_image_b64:
                        organigram_image_b64 = base64.b64encode(img_b).decode("utf-8")
                txt = b["text"]
                if not txt:
                    continue
                low = txt.lower().rstrip(":")
                if b["bold"] and ("juridische structuur" in low or "totale portefeuille" in low):
                    ond_mode = "skip"
                    continue
                if b["bold"] and len(txt) < 60:
                    if low == "onroerend goed":
                        txt = "Onroerend goed o.b.v. taxatierapport"
                    pending_tbl_title = txt
                    onderneming_paragraphs.append(txt)
                    ond_mode = "normal"
                    continue
                if ond_mode != "skip":
                    onderneming_paragraphs.append(txt)
            elif b["type"] == "tbl" and b.get("rows"):
                onderneming_tables.append({
                    "title": pending_tbl_title or "Specificatie",
                    "rows": b["rows"],
                })
                pending_tbl_title = ""

        if not onderneming_paragraphs:
            for sec_name in ("leendoel", "geldnemer"):
                for b in _get_sec(sec_name):
                    if b["type"] == "p" and b["text"]:
                        onderneming_paragraphs.append(b["text"])

        onderneming_text = "\n\n".join(onderneming_paragraphs)

        # 3. Financieringsbehoefte table -> Investeringsoverzicht + loan terms + Overige financiers
        investeringsoverzicht = []
        fb_blocks = _get_sec("financieringsbehoefte")
        fb_extra_text = []
        overige_financiers_paragraphs = []
        collin_tranche_amount = 0.0
        fb_mode = "inv"
        for b in fb_blocks:
            if b["type"] == "tbl" and b.get("rows"):
                for r in b["rows"]:
                    if len(r) >= 2:
                        investeringsoverzicht.append([r[0], r[1]])
            elif b["type"] == "p" and b["text"]:
                txt = b["text"]
                low = txt.lower().rstrip(":")
                if low == "investeringsoverzicht":
                    fb_mode = "inv"
                    continue
                if low == "overige financiers":
                    fb_mode = "overige"
                    continue
                # Check if line is "<Omschrijving> € <Bedrag>"
                m_inv = re.match(r"^(.*?)\s+(€\s*[\d.,]+(?:\s*,-)?)\s*$", txt)
                if fb_mode == "inv" and m_inv and not low.startswith("leenbedrag"):
                    lbl = m_inv.group(1).strip()
                    amt = m_inv.group(2).strip()
                    is_b = b["bold"] or lbl.lower().startswith(("totale investering", "lening via collin crowdfund"))
                    investeringsoverzicht.append([
                        {"text": lbl, "bold": is_b, "align": "left"},
                        {"text": amt, "bold": is_b, "align": "right"},
                    ])
                    if lbl.lower().startswith("lening via collin crowdfund"):
                        num_c = re.sub(r"[^\d]", "", amt.split(",")[0])
                        if num_c:
                            collin_tranche_amount = float(num_c)
                elif fb_mode == "overige":
                    overige_financiers_paragraphs.append(txt)
                else:
                    fb_extra_text.append(txt)

        if not financieringsbehoefte_paragraphs and fb_extra_text:
            financieringsbehoefte_paragraphs.extend(fb_extra_text)

        financieringsbehoefte_text = "\n\n".join(financieringsbehoefte_paragraphs)
        overige_financiers_text = "\n\n".join(overige_financiers_paragraphs)

        # Check if communicerende vaten text is present in financieringsbehoefte_text for streefbedrag
        m_comm = re.search(
            r"(gevraagde financiering via Collin Crowdfund\s+€\s*[\d.,]+,-.*?communicerende vaten.*?veranderen\.)",
            financieringsbehoefte_text,
            re.S | re.I,
        )
        if m_comm:
            comm_txt = m_comm.group(1).strip()
            streefbedrag_str = "De " + comm_txt[0].lower() + comm_txt[1:]

        # Parse exact loan parameters (grace_months, balloon_amount, rate, duration)
        loan_parse_source = "\n".join(filter(None, [financieringsbehoefte_text] + fb_extra_text))
        grace_months = 0
        balloon_amount = 0.0

        dutch_num_words = {
            "een": 1, "één": 1, "twee": 2, "drie": 3, "vier": 4, "vijf": 5,
            "zes": 6, "zeven": 7, "acht": 8, "negen": 9, "tien": 10,
            "elf": 11, "twaalf": 12, "achttien": 18, "vierentwintig": 24, "zesendertig": 36,
        }

        def _parse_int_or_word(val_str: str) -> int:
            v = (val_str or "").strip().lower()
            if v.isdigit():
                return int(v)
            return dutch_num_words.get(v, 0)

        if loan_parse_source:
            m_grace = (
                re.search(r"eerste\s+(\d+|[a-zà-ÿ]+)\s+maanden\s+aflossingsvrij", loan_parse_source, re.I)
                or re.search(r"aflossingsvrije\s+periode\s+van\s+(\d+|[a-zà-ÿ]+)\s+maanden", loan_parse_source, re.I)
                or re.search(r"(\d+|[a-zà-ÿ]+)\s+maanden\s+aflossingsvrij", loan_parse_source, re.I)
            )
            if m_grace:
                grace_months = _parse_int_or_word(m_grace.group(1))

            m_slot = re.search(
                r"slottermijn(?:\s+van|\s+bedraagt|\s+ter\s+hoogte\s+van)?\s*€\s*([\d.]+)",
                loan_parse_source,
                re.I,
            )
            if m_slot:
                balloon_amount = float(m_slot.group(1).replace(".", ""))
            elif re.search(
                r"volledig\s+aflossingsvrij|aflossingsvrije\s+lening|ineens\s+aan\s+het\s+einde\s+van\s+de\s+looptijd\s+afgelost|maanden,\s*aflossingsvrij",
                loan_parse_source,
                re.I,
            ):
                balloon_amount = principal_amount
                grace_months = 0
            elif grace_months >= duration_months and duration_months > 0:
                balloon_amount = principal_amount
                grace_months = 0

        # 4. Risico table -> D&B, Collin Credit Score, Afloscapaciteit, Zekerheden en voorwaarden
        presentatie_intro = ""
        kasstroom_paragraphs = []
        zekerheden_bullets = []
        ris_blocks = _get_sec("risico")
        ris_sub = ""
        for b in ris_blocks:
            if b["type"] != "p" or not b["text"]:
                continue
            txt = b["text"]
            low = txt.lower().strip()
            if b["bold"] and len(txt) < 45:
                if "dun & bradstreet" in low or "d&b" in low:
                    ris_sub = "db"
                    continue
                elif "collin credit score" in low:
                    ris_sub = "ccs"
                    continue
                elif low.startswith(("afloscapaciteit", "toelichting jaarcijfers", "kasstroom")):
                    ris_sub = "kasstroom"
                    continue
                elif low.startswith("solvabiliteit"):
                    ris_sub = "solvabiliteit"
                    continue
                elif low.startswith("zekerheden"):
                    ris_sub = "zekerheden"
                    continue
                elif low.startswith("samenvattend"):
                    ris_sub = "samenvattend"
                    continue
            if low.startswith("samenvattend"):
                ris_sub = "samenvattend"
                continue

            if ris_sub == "ccs":
                basis_m = re.search(
                    r"(?:Er is bij de beoordeling gebruik gemaakt van|De (?:beoordeling|Score|Collin Credit Score) is gebaseerd op)\s+(.*?)(?:\.\s*De overall Collin Credit Score|$)",
                    txt,
                    re.S | re.I,
                )
                if basis_m:
                    presentatie_intro = (
                        "De presentatie van de belangrijkste financiële cijfers is gebaseerd op "
                        + basis_m.group(1).strip().rstrip(".")
                        + "."
                    )
            elif ris_sub == "kasstroom":
                kasstroom_paragraphs.append(txt)
            elif ris_sub == "zekerheden":
                if b["is_list"]:
                    zekerheden_bullets.append(txt)
                elif not low.startswith("door de ") and "kwalificeren wij de zekerheden" not in low:
                    zekerheden_bullets.append(txt)

        if not zekerheden_bullets and samenvatting_zekerheden:
            zekerheden_bullets = samenvatting_zekerheden

        kasstroom_text = "\n\n".join(kasstroom_paragraphs)

        # 5. EBI cijfers / Jaarcijfers tables
        jaarcijfers_tables = []
        ebi_blocks = _get_sec("ebi cijfers") or _get_sec("jaarcijfers")
        cur_title = ""
        cur_sub = ""
        for b in ebi_blocks:
            if b["type"] == "p" and b["text"]:
                if b["bold"] or not cur_title:
                    cur_title = b["text"]
                else:
                    cur_sub = b["text"]
            elif b["type"] == "tbl" and b.get("rows"):
                rows = b["rows"]
                # Check if table already contains € symbols (meaning it's not x 1.000)
                has_euro = any("€" in c.get("text", "") for r in rows for c in r)
                default_sub = "" if has_euro else "In duizendtallen (x 1.000)"
                jaarcijfers_tables.append({
                    "title": cur_title,
                    "subtitle": cur_sub if cur_sub else default_sub,
                    "rows": rows,
                    "included": True,
                })
                cur_title = ""
                cur_sub = ""

        # 6. Extract Geldnemers, Rechtsvorm, Vestigingsplaats & KvK lookup
        geldnemers = []
        projecteigenaar_full = handelsnaam
        for b in zekerheden_bullets:
            m_gn = re.search(
                r"De geldnemers?\s+(?:zijn|is)\s+(.*?)(?<!\b[A-Z])\.(?:\s+[A-Z][a-z]|$)",
                b,
            )
            if m_gn:
                raw_names = m_gn.group(1).strip()
                projecteigenaar_full = re.sub(r"\s*\(in priv[eé]\)", "", raw_names, flags=re.I).strip()
                parts = [p.strip() for p in re.split(r"\s+en\s+|,\s*", raw_names) if p.strip()]
                kvk_matches = re.findall(r"\b(\d{8})\b", b)
                for idx, p in enumerate(parts):
                    p_clean = re.sub(r"([A-Za-z])(B\.V\.)", r"\1 \2", p)
                    p_clean = re.sub(r"\s*\(in priv[eé]\)", "", p_clean, flags=re.I).strip()
                    kvk_val = kvk_matches[idx] if idx < len(kvk_matches) else "n.v.t."
                    geldnemers.append({"naam": p_clean, "kvk": kvk_val, "dnb_url": ""})
                break

        if not geldnemers and handelsnaam:
            geldnemers.append({"naam": handelsnaam, "kvk": "n.v.t.", "dnb_url": ""})

        vestigingsplaats = ""
        for b in zekerheden_bullets:
            m_pc = re.search(r"\b\d{4}\s*[A-Z]{2}\s+(?:te\s+)?([A-ZÀ-ÿIJ][a-zà-ÿijA-Z\s'-]+?)(?:,|\.|ten behoeve)", b)
            if m_pc:
                vestigingsplaats = m_pc.group(1).strip()
                break
            m_te = re.search(r"\bte\s+([A-ZÀ-ÿIJ][a-zà-ÿijA-Z'-]+)(?:,|\.)", b)
            if m_te:
                vestigingsplaats = m_te.group(1).strip()
                break

        has_any_corp = any(
            re.search(r"\b(B\.?V\.?|V\.?O\.?F\.?|N\.?V\.?|C\.?V\.?|Holding|Beheer|Stichting)\b", g.get("naam", ""), re.I)
            for g in geldnemers
        )
        rechtsvorm = "Besloten Vennootschap" if has_any_corp else "Box 3"

        for idx_gn, gn in enumerate(geldnemers):
            gn_name = gn.get("naam", "")
            has_corp = bool(re.search(r"\b(B\.?V\.?|V\.?O\.?F\.?|N\.?V\.?|C\.?V\.?|Holding|Beheer|Stichting)\b", gn_name, re.I))
            if gn.get("kvk") in ("", "n.v.t.", None) and has_corp:
                hits = lookup_kvk(gn_name)
                if hits:
                    gn["kvk"] = hits[0]["kvk"]
                    if idx_gn == 0 and hits[0].get("plaats"):
                        vestigingsplaats = hits[0]["plaats"]
                    elif not vestigingsplaats and hits[0].get("plaats"):
                        vestigingsplaats = hits[0]["plaats"]

        # 7. Organigram fallback nodes if no embedded image
        organigram_nodes = [{"naam": handelsnaam, "laag": 1, "is_geldnemer": True, "label": ""}]

        # 8. Risico sentences & Historie betaalverplichtingen
        if db_score.lower() in ("niet van toepassing", "n.v.t.", "nvt", ""):
            db_sentence = "De kwalificatie van Dun & Bradstreet voor deze aanvraag is niet van toepassing."
        else:
            db_sentence = f"De kwalificatie van Dun & Bradstreet voor deze aanvraag is {db_score}."
        cc_sentence = f"De overall Collin Credit Score is {cc_score}."

        is_private = rechtsvorm.lower() == "box 3" or db_score.lower() == "niet van toepassing"
        historie_instantie = "BKR" if is_private else "Dun & Bradstreet"
        historie_text = (
            f"Heeft de Projecteigenaar de laatste vijf jaar een betalingsachterstand van 90 dagen of meer gehad? "
            f"Conform recent inzicht vanuit {historie_instantie} zijn deze niet bekend."
        )

        loan_type = "Annuïtair" if "annu" in (financieringsbehoefte_text or "").lower() else "Lineair"

        betaalschema = generate_amortization_schedule(
            principal=principal_amount,
            annual_rate_pct=annual_rate,
            months=duration_months,
            grace_months=grace_months,
            balloon_amount=balloon_amount,
            succes_fee_pct=2.0,
            admin_fee_pct=0.07,
            loan_type=loan_type,
        )

        return {
            "source_url": filename or "Word Pitch Document (.docx)",
            "lei_code": "724500VFNB9F43XDI745",
            "leningnummer": leningnummer,
            "investeringskans": project_title,
            "projecteigenaar": projecteigenaar_full,
            "projectnaam": project_title,
            "streefbedrag": streefbedrag_str or format_euro(principal_amount, 0),
            "collin_tranche_amount": collin_tranche_amount,
            "deadline": "30 dagen na publicatiedatum",
            "rechtsvorm": rechtsvorm,
            "geldnemers": geldnemers,
            "website": "n.v.t.",
            "vestigingsplaats": vestigingsplaats,
            "organigram_text": handelsnaam,
            "organigram_nodes": organigram_nodes,
            "organigram_image_b64": organigram_image_b64,
            "onderneming_text": onderneming_text,
            "onderneming_tables": onderneming_tables,
            "presentatie_intro": presentatie_intro,
            "kasstroom_text": kasstroom_text,
            "overige_financiers_text": overige_financiers_text,
            "jaarcijfers_tables": jaarcijfers_tables,
            "kvk_jaarcijfers_text": "Niet van toepassing.",
            "investeringsoverzicht": investeringsoverzicht,
            "bodembedrag": "n.v.t.",
            "inschrijfperiode": "Maximaal 30 dagen na publicatie",
            "bedenktermijn": "Vier kalenderdagen",
            "collin_invested_note": (
                "In deze investeringskans heeft/hebben de aandeelhouder(s), bestuurder(s) "
                "en/of werknemer(s) van Collin reeds geïnvesteerd."
            ),
            "kasstroom_risico_afsluiting": (
                "De aangeleverde en beoordeelde (cijfer)informatie is in belangrijke mate "
                "gebaseerd op historische, en toekomstige (cijfer)informatie. Het risico bestaat "
                "dat de toekomstige kasstroom niet toereikend is om aan de betalingsverplichting(en) "
                "te kunnen voldoen."
            ),
            "include_herfinanciering": balloon_amount > 0,
            "herfinanciering_text": (
                "Bij deze lening is er sprake van restschuld op het einde van de looptijd. "
                "Deze restschuld dient de Projecteigenaar af te lossen. Dit kan bijvoorbeeld door "
                "een herfinanciering via Collin of elders, verkoop van het onroerend goed of "
                "beschikbare liquiditeiten. Er is een risico dat de restschuld niet tijdig wordt "
                "afgelost. Indien er sprake is van een lening met een restschuld is Projecteigenaar "
                "van mening dat deze restschuld verantwoord is op basis van de huidige inzichten."
            ),
            "financieringsbehoefte_text": financieringsbehoefte_text,
            "risico_scores_text": f"{db_sentence}\n{cc_sentence}",
            "zekerheden_bullets": zekerheden_bullets,
            "loan_params": {
                "principal": principal_amount,
                "annual_rate_pct": annual_rate,
                "months": duration_months,
                "grace_months": grace_months,
                "balloon_amount": balloon_amount,
                "succes_fee_pct": 2.0,
                "admin_fee_pct": 0.07,
                "loan_type": loan_type,
            },
            "betaalschema": betaalschema,
            "historie_betaalverplichtingen": historie_text,
            "beheervergoeding": "0,1%",
        }
