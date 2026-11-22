import io, json, re
import pandas as pd
import streamlit as st
from pypdf import PdfReader

st.set_page_config(page_title="FinsearchAi", page_icon="📊", layout="wide")
MODEL = "claude-sonnet-5-5"  # change here if you want a cheaper/other model


def secret(k):
    try:
        return st.secrets[k]
    except Exception:
        return ""


# ---------- simple password lock (set APP_PASSWORD in Streamlit secrets) ----------
pw = secret("APP_PASSWORD")
if pw and not st.session_state.get("ok"):
    x = st.text_input("Password", type="password")
    if x == pw:
        st.session_state.ok = True
        st.rerun()
    elif x:
        st.error("Wrong password")
    st.stop()

st.title("📊 FinsearchAi")
st.caption("Annual report extraction, checks, DCF and WACC. Amounts as reported in the document.")

key = secret("ANTHROPIC_API_KEY") or st.sidebar.text_input("Claude API key (not saved)", type="password")
st.sidebar.info("Without a key you can still read PDFs and use the DCF and WACC tabs. Extraction and Q&A need the key.")


def ask_claude(prompt, system, max_tokens=6000):
    import anthropic
    c = anthropic.Anthropic(api_key=key)
    r = c.messages.create(model=MODEL, max_tokens=max_tokens, system=system,
                          messages=[{"role": "user", "content": prompt}])
    return r.content[0].text


@st.cache_data(show_spinner="Reading PDF pages...")
def read_pdf(b):
    return [(p.extract_text() or "") for p in PdfReader(io.BytesIO(b)).pages]


def parse_pages(s, n):
    out = []
    for part in re.split(r"[,\s]+", s.strip()):
        if not part:
            continue
        a, _, b = part.partition("-")
        try:
            lo, hi = int(a), int(b or a)
        except ValueError:
            continue
        out += [p for p in range(lo, hi + 1) if 1 <= p <= n]
    return sorted(set(out))


def suggest(pages, words):
    sc = sorted(((sum(w in t.lower() for w in words), i + 1) for i, t in enumerate(pages)), reverse=True)
    top = sorted(p for s, p in sc[:3] if s >= 2)
    return ", ".join(map(str, top))


def page_text(pages, nums, limit=60000):
    return "\n".join(f"=== PAGE {p} ===\n{pages[p-1]}" for p in nums)[:limit]


STATEMENTS = {
    "Balance Sheet": ["balance sheet", "total assets", "total equity and liabilities"],
    "Profit and Loss": ["statement of profit and loss", "revenue from operations", "total income"],
    "Cash Flow": ["cash flow", "operating activities", "net cash"],
}
SYS_EXTRACT = (
    "You extract figures from Indian annual report text. Return ONLY a JSON list, no other text. "
    "Each item: {\"statement\",\"particular\",\"fy\",\"value\",\"unit\",\"basis\",\"page\"}. "
    "One item per line item per year column. value is a plain number (no commas, negatives with minus). "
    "basis is Consolidated or Standalone as printed on the page. page is the PAGE number marker. "
    "NEVER invent or estimate a number. If a figure is unclear or missing, leave it out."
)


def clean_json(t):
    t = re.sub(r"^```(?:json)?|```$", "", t.strip(), flags=re.M).strip()
    return json.loads(t)


def to_excel(df):
    b = io.BytesIO()
    with pd.ExcelWriter(b, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Source data", index=False)
        for s, g in df.groupby("statement", sort=False):
            p = g.pivot_table(index="particular", columns="fy", values="value", aggfunc="first")
            p = p.reindex(g["particular"].drop_duplicates())
            p.to_excel(w, sheet_name=re.sub(r"[\\/*?:\[\]]", "", str(s))[:30])
    return b.getvalue()


def dcf(rev, gr, mg, da, cx, nwc, tax, w, tg, mult, nd, sh):
    r, pv, rows = rev, 0.0, []
    for i, g in enumerate(gr, 1):
        prev, r = r, r * (1 + g / 100)
        eb = r * mg / 100
        d = r * da / 100
        fc = (eb - d) * (1 - tax / 100) + d - r * cx / 100 - (r - prev) * nwc / 100
        pv += fc / (1 + w / 100) ** i
        rows.append({"Year": f"Y{i}", "Revenue": r, "EBITDA": eb, "EBIT": eb - d, "FCFF": fc})
    n, disc = len(gr), (1 + w / 100) ** len(gr)
    tvg = fc * (1 + tg / 100) / ((w - tg) / 100) if w > tg else float("nan")
    out = {}
    for name, tv in (("Gordon Growth", tvg), ("Exit Multiple", eb * mult)):
        ev = pv + tv / disc
        out[name] = {"Terminal value": tv, "PV of TV": tv / disc, "Enterprise value": ev,
                     "Equity value": ev - nd, "Value per share": (ev - nd) / sh, "TV % of EV": tv / disc / ev * 100}
    return pd.DataFrame(rows), out


t1, t2, t3, t4 = st.tabs(["1. Extract from annual report", "2. Ask questions", "3. DCF and sensitivity", "4. WACC"])

with t1:
    up = st.file_uploader("Upload annual report PDF", type="pdf", key="u1")
    if up:
        pages = read_pdf(up.getvalue())
        st.success(f"{len(pages)} pages read. Check the suggested pages, fix them if wrong.")
        sel = {s: st.text_input(f"{s} pages (example 120-123)", suggest(pages, w), key="p" + s)
               for s, w in STATEMENTS.items()}
        if st.button("Extract figures", type="primary"):
            if not key:
                st.error("Add the Claude API key first (sidebar).")
            else:
                recs = []
                for s, txt in sel.items():
                    nums = parse_pages(txt, len(pages))
                    if not nums:
                        continue
                    with st.spinner(f"Extracting {s}..."):
                        try:
                            recs += clean_json(ask_claude(f"Statement: {s}\n\n{page_text(pages, nums)}", SYS_EXTRACT))
                        except Exception as e:
                            st.error(f"{s}: {e}")
                if recs:
                    df = pd.DataFrame(recs)
                    df["value"] = pd.to_numeric(df["value"].astype(str).str.replace(",", ""), errors="coerce")
                    st.session_state.df = df
        if "df" in st.session_state:
            df = st.session_state.df
            st.dataframe(df, width="stretch")
            st.subheader("Check: does the balance sheet balance?")
            a = df[df.particular.str.lower().str.contains("total assets", na=False)].groupby("fy").value.first()
            l = df[df.particular.str.lower().str.contains("total equity and liabilities", na=False)].groupby("fy").value.first()
            chk = pd.DataFrame({"Total assets": a, "Total equity and liabilities": l}).dropna()
            if len(chk):
                chk["Difference"] = chk.iloc[:, 0] - chk.iloc[:, 1]
                chk["Status"] = chk["Difference"].abs().lt(1).map({True: "OK", False: "MISMATCH - check pages"})
                st.dataframe(chk)
            else:
                st.warning("Totals not found. Check the balance sheet page numbers.")
            st.warning("Always verify key numbers against the PDF page shown in the page column.")
            st.download_button("Download Excel", to_excel(df), "extracted_financials.xlsx")

with t2:
    up2 = st.file_uploader("Upload annual report or concall PDF", type="pdf", key="u2")
    if up2:
        pg = read_pdf(up2.getvalue())
        rng = st.text_input("Pages to read (example 10-40)", "1-30")
        q = st.text_area("Your question")
        if st.button("Ask") and q:
            if not key:
                st.error("Add the Claude API key first (sidebar).")
            else:
                sysq = ("You are a finance mentor for a fresher. Use simple Indian English and ₹ crore. "
                        "Answer ONLY from the given pages and quote the page number for each fact. "
                        "If the answer is not in the pages, say so. Never invent figures.")
                try:
                    st.markdown(ask_claude(f"{page_text(pg, parse_pages(rng, len(pg)), 120000)}\n\nQuestion: {q}", sysq, 2500))
                except Exception as e:
                    st.error(e)

with t3:
    st.caption("Dummy numbers. Replace with your own. FCFF = EBIT x (1 - tax) + D&A - capex - change in NWC")
    c = st.columns(4)
    rev = c[0].number_input("Base revenue", value=10000.0)
    grs = c[1].text_input("Growth % per year", "12,11,10,9,8")
    mg = c[2].number_input("EBITDA margin %", value=22.0)
    da = c[3].number_input("D&A % revenue", value=4.0)
    c = st.columns(4)
    cx = c[0].number_input("Capex % revenue", value=5.0)
    nwc = c[1].number_input("NWC % of change in revenue", value=15.0)
    tax = c[2].number_input("Tax %", value=25.17)
    w = c[3].number_input("WACC %", value=11.5)
    c = st.columns(4)
    tg = c[0].number_input("Terminal growth %", value=5.0)
    mult = c[1].number_input("Exit EV/EBITDA", value=14.0)
    nd = c[2].number_input("Net debt", value=1500.0)
    sh = c[3].number_input("Shares (crore)", value=10.0)
    try:
        gr = [float(x) for x in grs.split(",")]
        rows, out = dcf(rev, gr, mg, da, cx, nwc, tax, w, tg, mult, nd, sh)
        st.dataframe(rows.round(1), width="stretch")
        st.dataframe(pd.DataFrame(out).round(1), width="stretch")
        if tg >= w:
            st.error("Terminal growth must be below WACC.")
        if tg > 6:
            st.warning("Terminal growth above 6% is aggressive for India.")
        if out["Gordon Growth"]["TV % of EV"] > 75:
            st.warning("Terminal value is more than 75% of EV. Result depends heavily on terminal assumptions.")
        st.subheader("Value per share: WACC vs terminal growth (Gordon)")
        ws = [w + d for d in (-1, -0.5, 0, 0.5, 1)]
        gs = [tg + d for d in (-1, -0.5, 0, 0.5, 1)]
        grid = pd.DataFrame([[dcf(rev, gr, mg, da, cx, nwc, tax, x, y, mult, nd, sh)[1]["Gordon Growth"]["Value per share"]
                              for y in gs] for x in ws], index=[f"{x:.1f}%" for x in ws], columns=[f"{y:.1f}%" for y in gs])
        st.dataframe(grid.round(1))
    except Exception as e:
        st.error(f"Check your inputs: {e}")

with t4:
    c = st.columns(4)
    rf = c[0].number_input("Risk-free rate %", value=7.0)
    erp = c[1].number_input("Equity risk premium %", value=6.0)
    bl = c[2].number_input("Levered beta", value=1.1)
    kd = c[3].number_input("Pre-tax cost of debt %", value=8.5)
    c = st.columns(3)
    E = c[0].number_input("Market value of equity", value=60000.0)
    D = c[1].number_input("Market value of debt", value=6000.0)
    tx = c[2].number_input("Tax rate %", value=25.17)
    ke = rf + bl * erp
    wacc = E / (E + D) * ke + D / (E + D) * kd * (1 - tx / 100)
    st.table(pd.DataFrame({"Item": ["Cost of equity = Rf + beta x ERP", "Unlevered beta = BL / (1+(1-t) x D/E)",
                                    "WACC = E/V x Ke + D/V x Kd x (1-t)"],
                           "Value": [f"{ke:.2f}%", f"{bl / (1 + (1 - tx / 100) * D / E):.3f}", f"{wacc:.2f}%"]}))
