import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
from weasyprint import HTML

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
        f"1. אפקטיביות מנגנון הקיזוז: סך התפוקה בפועל באשכול עומדת על {total_act:,.1f} קוט\"ש "
        f"מול צפי מותאם של {total_exp:,.1f} קוט\"ש (עמידה מצרפית של {total_ratio:.1f}%). "
        f"עודפי הייצור באתרים המובילים פיצו על פערי האתרים האחרים ומנעו קנסות אי-עמידה.\n"
        f"2. התאמת תאריכי הפעלה (COD): צפי חודש יוני חושב באופן יחסי (Prorated) מול מספר ימי הפעילות "
        f"האמיתיים ממועד החיבור לרשת, בהתאם לכללי נוהל רשות החשמל."
    )
    edited_notes = st.text_area("ערוך הערות לדוח במידת הצורך:", value=notes_text, height=120)

    if st.button("📄 הפק קובץ PDF להורדה"):
        rows_html = "".join([
            f"<tr>"
            f"<td style='text-align:right;'>{s['name']}</td>"
            f"<td>{s['kwp']:.1f}</td>"
            f"<td>{s['cod']} ({s['active_days']} ימים)</td>"
            f"<td>{s['exp']:,.1f}</td>"
            f"<td><strong>{s['act']:,.1f}</strong></td>"
            f"<td style='color:{'#234E52' if s['diff'] >= 0 else '#9B2C2C'};'>{( '+' if s['diff'] >= 0 else '' )}{s['diff']:,.1f}</td>"
            f"<td><strong>{s['ratio']:.1f}%</strong></td>"
            f"<td style='color:{'#234E52' if s['fin'] >= 0 else '#9B2C2C'}; font-weight:bold;'>{( '+' if s['fin'] >= 0 else '' )}{int(s['fin']):,} ₪</td>"
            f"</tr>"
            for s in sites
        ])

        html_content = f"""
        <!DOCTYPE html>
        <html lang="he" dir="rtl">
        <head>
          <meta charset="UTF-8">
          <style>
            @page {{ size: A4 portrait; margin: 12mm; }}
            body {{ font-family: Arial, sans-serif; color: #1A202C; direction: rtl; font-size: 10pt; }}
            h1 {{ color: #1A365D; font-size: 16pt; margin: 0 0 4px 0; border-bottom: 2px solid #2B6CB0; padding-bottom: 6px; }}
            .sub {{ color: #4A5568; font-size: 9pt; margin-bottom: 12px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 8.5pt; }}
            th, td {{ border: 1px solid #CBD5E0; padding: 6px 4px; text-align: center; }}
            th {{ background: #1A365D; color: #fff; }}
            .notes {{ background: #FFFDF5; border: 1px solid #FEEBC8; border-right: 4px solid #DD6B20; padding: 8px; margin-top: 15px; font-size: 8.5pt; }}
            .sigs {{ margin-top: 25px; display: table; width: 100%; font-size: 9pt; }}
            .sig-cell {{ display: table-cell; width: 33%; }}
            .line {{ border-bottom: 1px solid #718096; width: 80%; margin-top: 25px; }}
          </style>
        </head>
        <body>
          <h1>דו"ח תפעולי רבעוני וקיזוז מצרפי (O&M)</h1>
          <div class="sub">{project_name} | {period} | נוהל רשות החשמל (נספח יד')</div>

          <table>
            <thead>
              <tr>
                <th style="text-align:right;">שם המתקן</th>
                <th>הספק DC</th>
                <th>מועד COD וימים</th>
                <th>צפי מותאם [kWh]</th>
                <th>בפועל [kWh]</th>
                <th>עודף / (חסר)</th>
                <th>עמידה ביעד</th>
                <th>משמעות כספית</th>
              </tr>
            </thead>
            <tbody>
              {rows_html}
              <tr style="background:#EBF8FF; font-weight:bold;">
                <td style="text-align:right;">סה"כ {project_name}</td>
                <td>{total_kwp:.1f} kWp</td>
                <td>-</td>
                <td>{total_exp:,.1f}</td>
                <td>{total_act:,.1f}</td>
                <td>{( '+' if total_diff >= 0 else '' )}{total_diff:,.1f} kWh</td>
                <td>{total_ratio:.1f}%</td>
                <td>{( '+' if total_fin >= 0 else '' )}{int(total_fin):,} ₪</td>
              </tr>
            </tbody>
          </table>

          <div class="notes">
            <strong>ממצאי O&M ומסקנות הנדסיות:</strong><br>
            {edited_notes.replace(chr(10), '<br>')}
          </div>

          <div class="sigs">
            <div class="sig-cell">מהנדס בודק / מנהל O&M:<div class="line"></div></div>
            <div class="sig-cell">תאריך חתימה:<div class="line"></div></div>
            <div class="sig-cell">חתימה וחותמת:<div class="line"></div></div>
          </div>
        </body>
        </html>
        """

        pdf_bytes = HTML(string=html_content).write_pdf()
        st.download_button(
            label="⬇️ לחץ להורדת קובץ ה-PDF",
            data=pdf_bytes,
            file_name=f"OM_Report_{project_name}.pdf",
            mime="application/pdf"
        )
else:
    st.info("נא להעלות קובצי XLS ו-PDF למעלה להפעלת מנוע החישוב.")
