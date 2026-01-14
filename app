import streamlit as st
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from io import BytesIO
from datetime import datetime

# --- HELPER FUNCTIONS ---
def extract_quantities(value):
    if pd.isna(value): 
        return 0.0, 0.0
    val_str = str(value).strip()
    if '+' in val_str:
        parts = val_str.split('+')
        try:
            base_qty = float(parts[0]) if parts[0].strip() else 0.0
            free_qty = float(parts[1]) if parts[1].strip() else 0.0
            return base_qty, free_qty
        except:
            return 0.0, 0.0
    else:
        try:
            return float(val_str), 0.0
        except:
            return 0.0, 0.0

# --- UI SETUP ---
st.set_page_config(page_title="Ferrite Agencies Report", layout="centered")
st.title("📦 Ferrite Agencies")
st.subheader("Order Report Generator")

uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file is not None:
    try:
        # 1. Process Excel
        df = pd.read_excel(uploaded_file, sheet_name='Item Details', usecols="D,G,H,K,L")
        df.columns = ['Item Name', 'Category', 'MRP', 'Raw_Qty', 'Unit']

        qty_data = df['Raw_Qty'].apply(extract_quantities)
        df['Quantity'] = qty_data.apply(lambda x: x[0])
        df['Free_Quantity'] = qty_data.apply(lambda x: x[1])
        
        df['MRP'] = pd.to_numeric(df['MRP'], errors='coerce').fillna(0)
        df['Unit'] = df['Unit'].fillna("-").astype(str).str.strip()
        df['Category'] = df['Category'].fillna("Uncategorized").astype(str).str.strip()
        
        df = df.groupby(['Category', 'Item Name', 'Unit'], as_index=False).agg({
            'Quantity': 'sum',
            'Free_Quantity': 'sum',
            'MRP': 'first'
        }).sort_values(by=['Category', 'Item Name'])

        # 2. Generate PDF in Memory
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        elements = []
        styles = getSampleStyleSheet()
        
        cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=9, leading=11)
        title_style = ParagraphStyle('T', fontSize=24, alignment=TA_CENTER)
        
        elements.append(Paragraph("Ferrite Agencies", title_style))
        elements.append(Paragraph(f"Generated: {datetime.now().strftime('%d-%m-%Y %I:%M %p')}", styles['Normal']))
        elements.append(Spacer(1, 15))
        
        table_data = [['MRP', 'CATEGORY', 'ITEM NAME', 'UNIT', 'QTY', 'FREE QTY']]
        total_qty = df['Quantity'].sum()
        total_free = df['Free_Quantity'].sum()
        
        for _, row in df.iterrows():
            mrp_disp = f"{row['MRP']:.2f}" if row['MRP'] != 0 else ""
            table_data.append([
                mrp_disp,
                Paragraph(row['Category'], cell_style),
                Paragraph(row['Item Name'], cell_style),
                row['Unit'],
                int(row['Quantity']) if row['Quantity'].is_integer() else row['Quantity'],
                int(row['Free_Quantity']) if row['Free_Quantity'].is_integer() else "" if row['Free_Quantity'] == 0 else row['Free_Quantity']
            ])
            
        table_data.append(['', '', 'TOTAL', '', total_qty, total_free])
        
        t = Table(table_data, colWidths=[50, 80, 190, 60, 45, 55])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.darkblue),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        
        elements.append(t)
        doc.build(elements)
        
        # 3. Download Button
        st.success("Report Ready!")
        st.download_button(
            label="📩 Download PDF Report",
            data=buffer.getvalue(),
            file_name=f"Order_Report_{datetime.now().strftime('%H%M%S')}.pdf",
            mime="application/pdf"
        )

    except Exception as e:
        st.error(f"Error: {e}")
