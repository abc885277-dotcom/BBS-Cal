import io
import os
from typing import Dict, List

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

# Groq is optional at import time so the app can still run without an API key.
try:
    from groq import Groq
except ImportError:
    Groq = None


st.set_page_config(
    page_title="BBS Calculator | Streamlit",
    page_icon="🏗️",
    layout="wide",
)


# -----------------------------
# BBS calculation engine
# -----------------------------
SHAPES = {
    "00": "Straight",
    "11": "L / two-leg",
    "21": "U / three-leg",
    "51": "Closed / four-leg",
}


def unit_weight_kg_m(d_mm: float) -> float:
    """Nominal rebar unit weight, kg/m."""
    return (d_mm ** 2) / 162.2


def cut_length_m(shape_code: str, d_mm: float, a: float, b: float, c: float) -> float:
    d_m = d_mm / 1000.0
    if shape_code == "00":
        length = a
    elif shape_code == "11":
        length = a + b - 2 * d_m
    elif shape_code == "21":
        length = a + b + c - 4 * d_m
    elif shape_code == "51":
        length = 2 * (a + b) + 16 * d_m
    else:
        raise ValueError(f"Unsupported shape code: {shape_code}")
    return max(0.0, length)


def calculate_bbs(items: List[Dict], waste_pct: float) -> Dict:
    rows = []
    diameter_summary: Dict[float, float] = {}
    net_weight = 0.0

    for item in items:
        cl = cut_length_m(
            item["shape_code"], item["diameter_mm"],
            item["a_m"], item["b_m"], item["c_m"]
        )
        total_length = cl * item["quantity"]
        unit_wt = unit_weight_kg_m(item["diameter_mm"])
        weight = total_length * unit_wt

        rows.append({
            "Bar Mark": item["bar_mark"],
            "Dia (mm)": item["diameter_mm"],
            "Shape Code": item["shape_code"],
            "Shape": SHAPES[item["shape_code"]],
            "Qty": item["quantity"],
            "A (m)": item["a_m"],
            "B (m)": item["b_m"],
            "C (m)": item["c_m"],
            "Cut Length (m)": round(cl, 3),
            "Total Length (m)": round(total_length, 3),
            "Unit Weight (kg/m)": round(unit_wt, 3),
            "Weight (kg)": round(weight, 2),
        })

        diameter_summary[item["diameter_mm"]] = (
            diameter_summary.get(item["diameter_mm"], 0.0) + weight
        )
        net_weight += weight

    waste_weight = net_weight * waste_pct / 100.0
    gross_weight = net_weight + waste_weight

    return {
        "schedule": rows,
        "diameter_breakdown_kg": {
            f"{int(d) if float(d).is_integer() else d:g} mm": round(w, 2)
            for d, w in sorted(diameter_summary.items())
        },
        "net_total_kg": round(net_weight, 2),
        "waste_factor_pct": waste_pct,
        "waste_weight_kg": round(waste_weight, 2),
        "gross_total_kg": round(gross_weight, 2),
        "gross_total_tonnes": round(gross_weight / 1000.0, 3),
    }


def sample_items() -> List[Dict]:
    return [
        {"bar_mark": "B01", "diameter_mm": 16.0, "shape_code": "00",
         "quantity": 10, "a_m": 4.0, "b_m": 0.0, "c_m": 0.0},
        {"bar_mark": "B02", "diameter_mm": 10.0, "shape_code": "11",
         "quantity": 24, "a_m": 4.0, "b_m": 0.5, "c_m": 0.0},
    ]


# -----------------------------
# Export functions
# -----------------------------
def make_excel(result: Dict) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        schedule = pd.DataFrame(result["schedule"])
        schedule.to_excel(writer, sheet_name="BBS Schedule", index=False)

        summary = [
            {"Metric": f"Steel Dia {dia}", "Value": wt}
            for dia, wt in result["diameter_breakdown_kg"].items()
        ]
        summary += [
            {"Metric": "Net Total Weight (kg)", "Value": result["net_total_kg"]},
            {"Metric": f"Waste ({result['waste_factor_pct']}%)", "Value": result["waste_weight_kg"]},
            {"Metric": "Gross Total Weight (kg)", "Value": result["gross_total_kg"]},
            {"Metric": "Gross Total Weight (tonnes)", "Value": result["gross_total_tonnes"]},
        ]
        pd.DataFrame(summary).to_excel(writer, sheet_name="Summary", index=False)

        # Basic professional formatting.
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 28)

    buffer.seek(0)
    return buffer.getvalue()


def make_pdf(result: Dict) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        rightMargin=12 * mm, leftMargin=12 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm,
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "BbsTitle", parent=styles["Heading1"], fontSize=18,
        spaceAfter=8, textColor=colors.HexColor("#1e293b")
    )
    story = [Paragraph("Bar Bending Schedule (BBS) Report", title)]

    headers = [
        "Bar Mark", "Dia", "Shape", "Qty", "A (m)", "B (m)", "C (m)",
        "Cut Length", "Total Length", "Unit Wt.", "Weight (kg)"
    ]
    table_data = [headers]
    for r in result["schedule"]:
        table_data.append([
            r["Bar Mark"], str(r["Dia (mm)"]), r["Shape Code"], str(r["Qty"]),
            f'{r["A (m)"]:.3f}', f'{r["B (m)"]:.3f}', f'{r["C (m)"]:.3f}',
            f'{r["Cut Length (m)"]:.3f}', f'{r["Total Length (m)"]:.3f}',
            f'{r["Unit Weight (kg/m)"]:.3f}', f'{r["Weight (kg)"]:.2f}'
        ])

    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f172a")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd5e1")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
    ]))
    story += [table, Spacer(1, 10)]

    summary_data = [["Summary", "Value"]]
    for dia, wt in result["diameter_breakdown_kg"].items():
        summary_data.append([f"Diameter {dia}", f"{wt:.2f} kg"])
    summary_data += [
        ["Net Total", f'{result["net_total_kg"]:.2f} kg'],
        [f'Waste ({result["waste_factor_pct"]:.1f}%)', f'{result["waste_weight_kg"]:.2f} kg'],
        ["Gross Total", f'{result["gross_total_kg"]:.2f} kg'],
        ["Gross Total", f'{result["gross_total_tonnes"]:.3f} tonnes'],
    ]
    summary_table = Table(summary_data, colWidths=[65 * mm, 55 * mm], hAlign="LEFT")
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#334155")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, -4), (-1, -1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#cbd5e1")),
    ]))
    story.append(summary_table)
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------------
# Optional Groq AI assistant
# -----------------------------
def ask_groq(question: str, result: Dict) -> str:
    if Groq is None:
        return "Groq package is not installed. Add `groq` to requirements.txt."
    api_key = None
    try:
        api_key = st.secrets.get("GROQ_API_KEY")
    except Exception:
        pass
    api_key = api_key or os.getenv("GROQ_API_KEY")

    if not api_key:
        return "GROQ_API_KEY is not configured. Add it to Streamlit Secrets or your local environment."

    client = Groq(api_key=api_key)
    model = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")
    context = (
        f"BBS result: net steel={result['net_total_kg']} kg, "
        f"waste={result['waste_factor_pct']}%, gross steel={result['gross_total_kg']} kg. "
        f"Schedule rows={len(result['schedule'])}."
    )
    response = client.chat.completions.create(
        model=model,
        temperature=0.2,
        max_tokens=700,
        messages=[
            {
                "role": "system",
                "content": (
                    "You are a construction quantity/BBS assistant. "
                    "Explain calculations clearly. Do not claim a design is code-compliant "
                    "without complete project/design information. "
                    "Distinguish calculation assistance from structural engineering approval."
                ),
            },
            {"role": "user", "content": context + "\nQuestion: " + question},
        ],
    )
    return response.choices[0].message.content


# -----------------------------
# Streamlit UI
# -----------------------------
st.title("🏗️ Bar Bending Schedule (BBS) Calculator")
st.caption("Streamlit version — no FastAPI server required. Calculate, review, export and optionally ask Groq AI.")

with st.sidebar:
    st.header("Settings")
    waste_pct = st.number_input("Steel waste (%)", min_value=0.0, max_value=20.0, value=3.0, step=0.5)
    st.info("Groq is optional. The BBS calculator works without an API key.")

    if st.button("Load sample data", use_container_width=True):
        st.session_state["items"] = sample_items()

if "items" not in st.session_state:
    st.session_state["items"] = sample_items()

st.subheader("Bar Inputs")
st.write("Enter one row for each bar mark. Dimensions are in metres.")

input_df = pd.DataFrame(st.session_state["items"])
edited_df = st.data_editor(
    input_df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "bar_mark": st.column_config.TextColumn("Bar Mark"),
        "diameter_mm": st.column_config.NumberColumn("Diameter (mm)", min_value=1.0, step=1.0),
        "shape_code": st.column_config.SelectboxColumn("Shape Code", options=list(SHAPES.keys())),
        "quantity": st.column_config.NumberColumn("Quantity", min_value=1, step=1),
        "a_m": st.column_config.NumberColumn("A (m)", min_value=0.0, step=0.01),
        "b_m": st.column_config.NumberColumn("B (m)", min_value=0.0, step=0.01),
        "c_m": st.column_config.NumberColumn("C (m)", min_value=0.0, step=0.01),
    },
    hide_index=True,
    key="bbs_editor",
)

# Normalize editor values and calculate.
try:
    items = []
    for _, row in edited_df.iterrows():
        mark = str(row["bar_mark"]).strip()
        if not mark:
            continue
        shape = str(row["shape_code"]).strip()
        if shape not in SHAPES:
            raise ValueError(f"{mark}: invalid shape code {shape}.")
        item = {
            "bar_mark": mark,
            "diameter_mm": float(row["diameter_mm"]),
            "shape_code": shape,
            "quantity": int(row["quantity"]),
            "a_m": float(row["a_m"]),
            "b_m": float(row["b_m"]),
            "c_m": float(row["c_m"]),
        }
        if item["diameter_mm"] <= 0 or item["quantity"] <= 0:
            raise ValueError(f"{mark}: diameter and quantity must be greater than zero.")
        items.append(item)

    if not items:
        raise ValueError("Add at least one bar item.")

    result = calculate_bbs(items, waste_pct)
except (ValueError, TypeError, KeyError) as exc:
    st.error(str(exc))
    st.stop()

st.session_state["items"] = items

st.subheader("Results")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Net Steel", f"{result['net_total_kg']:.2f} kg")
c2.metric("Waste", f"{result['waste_weight_kg']:.2f} kg")
c3.metric("Gross Steel", f"{result['gross_total_kg']:.2f} kg")
c4.metric("Gross Tonnes", f"{result['gross_total_tonnes']:.3f} t")

st.dataframe(pd.DataFrame(result["schedule"]), use_container_width=True, hide_index=True)

col1, col2 = st.columns(2)
with col1:
    st.download_button(
        "⬇️ Download Excel",
        data=make_excel(result),
        file_name="Bar_Bending_Schedule.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col2:
    st.download_button(
        "⬇️ Download PDF",
        data=make_pdf(result),
        file_name="Bar_Bending_Schedule.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

st.subheader("Groq AI Assistant")
question = st.text_area(
    "Ask about the current BBS",
    placeholder="Example: Explain how the steel weight was calculated.",
)
if st.button("Ask Groq", type="primary", disabled=not question.strip()):
    with st.spinner("Asking Groq..."):
        try:
            st.write(ask_groq(question.strip(), result))
        except Exception as exc:
            st.error(f"Groq request failed: {exc}")

st.caption("Note: Verify cutting-length, anchorage, development-length, bend and lap rules against the project drawings and the governing code before construction.")
