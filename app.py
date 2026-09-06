import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
from fpdf import FPDF

st.set_page_config(page_title="מערכת O&M סולארית", layout="wide")

def parse_growatt_xls(file_bytes, filename):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes))
        text_data = df.to_string()
    except Exception:
        text_data = file_bytes.decode('latin-1', errors='ignore')

    sn_match = re.search(r'(FVLJ[A-Z0-9]+|UMHS[A-Z0-9]+|UNHS[A-Z0-9]+|[A-Z0-9]{10,16})', text_data)
    sn = sn_match.group(1) if sn_match else "Growatt"

    nums = re.findall(r'\b\d{3,6}\.\d{1,2}\b|\b\d{3,6}\b', text_data)
    nums = [float(n) for n in nums if float(n) > 500.0]

    m1, m2, m3 = 0.0, 0.0, 0.0
    if len(nums) >= 3:
        m1, m2, m3 = nums[-3], nums[-2], nums[-1]

    clean_name = filename.replace('.xls', '').replace('.xlsx', '').strip()
    return {"name": clean_name, "sn": sn, "m1": m1, "m2": m2, "m3": m3}

def parse_pvsyst_pdf(file_bytes):
    text = ""
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"

    kwp = 0.0
    kwp_match = re.search(r'([0-9]{2,4}(?:\.[0-9]+)?)\s*kWp', text, re.IGNORECASE)
    if kwp_match:
        kwp = float(kwp_match.group(1))

    exp_june, exp_july, exp_aug = 0.0, 0.0, 0.0
    mwh_matches = re.findall(r'(\d{1,3}\.\d{1,3})\s*MWh', text)
    if len(mwh_matches) >= 8:
        exp_june = float(mwh_matches[5]) * 1000
        exp_july = float(mwh_matches[6]) * 1000
        exp_aug = float(mwh_matches[7]) * 1000
    elif kwp > 0:
        exp_june = kwp * 165
        exp_july = kwp * 170
        exp_aug = kwp * 155

    return {"kwp": kwp, "exp_june": exp_june, "exp_july": exp_july, "exp_aug": exp_aug}

def build_pdf_report(project_name, period, tariff, sites, total_kwp, total_act, total_exp, total_diff, total_ratio, total_fin, notes_text):
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # כותרת
    pdf.set_font("Helvetica", style="B", size=16)
    pdf.cell(0, 10, f"Solar O&M Executive Report - {project_name}", new_x="LMARGIN", new_y="NEXT", align="C")
    
    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 6, f"Reporting Period: {period} | Feed-in Tariff: {tariff} ILS/kWh", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(5)
    
    # KPI Box
    pdf.set_fill_color(240, 244, 248)
    pdf.set_draw_color(43, 108, 176)
    pdf.set_line_width(0.3)
    pdf.rect(10, pdf.get_y(), 190, 18, style="DF")
    
    pdf.set_xy(10, pdf.get_y() + 2)
    pdf.set_font("Helvetica", style="B", size=9)
    pdf.cell(47, 5, "Total DC Capacity", align="C")
    pdf.cell(47, 5, "Total Actual Yield", align="C")
    pdf.cell(47, 5, "Prorated Target", align="C")
    pdf.cell(47, 5, "Cluster Performance", align="C", new_x="LMARGIN", new_y="NEXT")
    
    pdf.set_font("Helvetica", style="B", size=11)
    pdf.set_text_color(26, 54, 93)
    pdf.cell(47, 7, f"{total_kwp:.1f} kWp", align="C")
    pdf.cell(47, 7, f"{total_act:,.1f} kWh", align="C")
    pdf.cell(47, 7, f"{total_exp:,.1f} kWh", align="C")
    pdf.cell(47, 7, f"{total_ratio:.1f}%", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(8)
    
    # טבלה
    pdf.set_font("Helvetica", style="B", size=8)
    pdf.set_fill_color(26, 54, 93)
    pdf.set_text_color(255, 255, 255)
    
    col_w = [40, 20, 25, 25, 25, 25, 30]
    headers = ["Site Name", "kWp", "COD (Days)", "Target (kWh)", "Actual (kWh)", "Variance", "Financial Net"]
    
    for i, h in enumerate(headers):
        pdf.cell(col_w[i], 7, h, border=1, fill=True, align="C")
    pdf.ln()
    
    pdf.set_font("Helvetica", size=8)
    pdf.set_text_color(0, 0, 0)
    for s in sites:
        pdf.cell(col_w[0], 6, str(s['name'])[:22], border=1, align="L")
        pdf.cell(col_w[1], 6, f"{s['kwp']:.1f}", border=1, align="C")
        pdf.cell(col_w[2], 6, f"{s['cod']} ({s['active_days']}d)", border=1, align="C")
        pdf.cell(col_w[3], 6, f"{s['exp']:,.1f}", border=1, align="C")
        pdf.cell(col_w[4], 6, f"{s['act']:,.1f}", border=1, align="C")
        pdf.cell(col_w[5], 6, f"{s['diff']:+,.1f}", border=1, align="C")
        pdf.cell(col_w[6], 6, f"{int(s['fin']):+,} ILS", border=1, align="C")
        pdf.ln()
        
    # שורת סיכום
    pdf.set_font("Helvetica", style="B", size=8)
    pdf.set_fill_color(235, 248, 255)
    pdf.cell(col_w[0], 7, "Total Portfolio", border=1, fill=True, align="L")
    pdf.cell(col_w[1], 7, f"{total_kwp:.1f}", border=1, fill=True, align="C")
    pdf.cell(col_w[2], 7, "-", border=1, fill=True, align="C")
    pdf.cell(col_w[3], 7, f"{total_exp:,.1f}", border=1, fill=True, align="C")
    pdf.cell(col_w[4], 7, f"{total_act:,.1f}", border=1, fill=True, align="C")
    pdf.cell(col_w[5], 7, f"{total_diff:+,.1f}", border=1, fill=True, align="C")
    pdf.cell(col_w[6], 7, f"{int(total_fin):+,} ILS", border=1, fill=True, align="C")
    pdf.ln(10)
    
    # הערות
    pdf.set_font("Helvetica", style="B", size=9)
    pdf.cell(0, 6, "O&M Findings and Executive Summary:", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", size=8)
    pdf.multi_cell(0, 5, notes_text)
    pdf.ln(10)
    
    # חתימות
    pdf.set_font("Helvetica", size=9)
    pdf.cell(60, 6, "Lead O&M Engineer: ________________", align="L")
    pdf.cell(60, 6, "Date: ______________", align="C")
    pdf.cell(70, 6, "Stamp & Signature: ________________", align="R")
    
    return bytes(pdf.output())

st.title("☀️ מערכת O&M וקיזוז מצרפי סולארי")
st.markdown("טען דוחות גולמיים לקבלת ניתוח הנדסי מצרפי והפקת דוח רשמי")

col_setup1, col_setup2, col_setup3 = st.columns(3)
with col_setup1:
    project_name = st.text_input("שם הפרויקט / אשכול", value="אשכול גבעת אלה")
with col_setup2:
    period = st.text_input("תקופת הדיווח", value="רבעון 3 (יוני - אוגוסט 2026)")
with col_setup3:
    tariff = st.number_input("תעריף הסדרה (₪ לקוט\"ש)", value=0.45, step=0.01)

st.divider()

c1, c2 = st.columns(2)
with c1:
    growatt_files = st.file_uploader("📂 העלה קובצי Growatt (XLS)", type=["xls", "xlsx"], accept_multiple_files=True)
with c2:
    pvsyst_files = st.file_uploader("📑 העלה דוחות PVsyst (PDF)", type=["pdf"], accept_multiple_files=True)

sites = []

if growatt_files:
    for f in growatt_files:
        g_data = parse_growatt_xls(f.read(), f.name)
        
        matched_pvsyst = None
        if pvsyst_files:
            for pf in pvsyst_files:
                if any(part in pf.name for part in g_data["name"].split()):
                    matched_pvsyst = parse_pvsyst_pdf(pf.getvalue())
                    break
        
        if not matched_pvsyst:
            matched_pvsyst = {"kwp": 50.0, "exp_june": 7500.0, "exp_july": 7800.0, "exp_aug": 7200.0}

        active_days_june = 30
        if "ספורט" in g_data["name"]:
            active_days_june = 23
            cod = "08/06/2026"
        elif "מזכירות" in g_data["name"]:
            active_days_june = 24
            cod = "07/06/2026"
        elif "נוער" in g_data["name"]:
            active_days_june = 22
            cod = "09/06/2026"
        elif "צרכני" in g_data["name"]:
            active_days_june = 18
            cod = "13/06/2026"
        else:
            cod = "01/06/2026"

        prorated_june = matched_pvsyst["exp_june"] * (active_days_june / 30.0)
        tot_exp = prorated_june + matched_pvsyst["exp_july"] + matched_pvsyst["exp_aug"]
        tot_act = g_data["m1"] + g_data["m2"] + g_data["m3"]
        diff = tot_act - tot_exp
        ratio = (tot_act / tot_exp * 100) if tot_exp > 0 else 0
        fin_impact = diff * tariff

        sites.append({
            "name": g_data["name"],
            "kwp": matched_pvsyst["kwp"],
            "inv": g_data["sn"],
            "cod": cod,
            "active_days": active_days_june,
            "m1": g_data["m1"],
            "m2": g_data["m2"],
            "m3": g_data["m3"],
            "act": tot_act,
            "exp": tot_exp,
            "diff": diff,
            "ratio": ratio,
            "fin": fin_impact
        })

if sites:
    df_display = pd.DataFrame(sites)
    
    st.subheader("📊 טבלת קיזוז מצרפי מרוכזת (Portfolio Netting)")
    st.dataframe(
        df_display[["name", "kwp", "cod", "active_days", "m1", "m2", "m3", "act", "exp", "diff", "ratio", "fin"]].rename(columns={
            "name": "מתקן",
            "kwp": "הספק (kWp)",
            "cod": "תאריך הפעלה",
            "active_days": "ימי יוני",
            "m1": "יוני [kWh]",
            "m2": "יולי [kWh]",
            "m3": "אוגוסט [kWh]",
            "act": "בפועל [kWh]",
            "exp": "צפי מותאם [kWh]",
            "diff": "עודף/(חסר)",
            "ratio": "עמידה ביעד",
            "fin": "משמעות כספית (₪)"
        }),
        use_container_width=True
    )

    total_kwp = sum(s["kwp"] for s in sites)
    total_act = sum(s["act"] for s in sites)
    total_exp = sum(s["exp"] for s in sites)
    total_diff = total_act - total_exp
    total_ratio = (total_act / total_exp * 100) if total_exp > 0 else 0
    total_fin = total_diff * tariff

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("סך הספק מותקן", f"{total_kwp:.1f} kWp")
    k2.metric("סך ייצור בפועל", f"{total_act:,.1f} kWh")
    k3.metric("צפי מותאם מצרפי", f"{total_exp:,.1f} kWh")
    k4.metric("עמידה מצרפית ביעד", f"{total_ratio:.1f}%", delta=f"{total_diff:,.1f} kWh")

    st.subheader("📝 מסקנות והערות O&M")
    notes_text = (
        f"1. Netting Mechanism Efficiency: Total actual generation reached {total_act:,.1f} kWh "
        f"vs prorated target of {total_exp:,.1f} kWh ({total_ratio:.1f}% target realization). "
        f"Outperforming sites successfully offset variance across the portfolio.\n"
        f"2. COD Proration: June target values were calculated on a pro-rata basis according to actual days of grid connection."
    )
    edited_notes = st.text_area("הערות ומסקנות לדוח:", value=notes_text, height=100)

    if st.button("📄 הפק קובץ PDF להורדה"):
        pdf_data = build_pdf_report(
            project_name, period, tariff, sites, 
            total_kwp, total_act, total_exp, total_diff, total_ratio, total_fin, edited_notes
        )
        st.download_button(
            label="⬇️ לחץ להורדת קובץ ה-PDF",
            data=pdf_data,
            file_name=f"OM_Report_{project_name}.pdf",
            mime="application/pdf"
        )
else:
    st.info("נא להעלות קובצי XLS ו-PDF למעלה להפעלת מנוע החישוב.")
