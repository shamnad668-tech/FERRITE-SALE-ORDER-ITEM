import streamlit as st
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from io import BytesIO
from datetime import datetime, timedelta

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

# --- PAGE CONFIG ---
st.set_page_config(page_title="Ferrite Agencies", page_icon="📦")

st.title("Ferrite Agencies")
st.subheader("Order Report System")

uploaded_file = st.file_uploader("Choose Excel File", type=['xlsx'])

if uploaded_file:
    try:
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

        # --- PDF GENERATION ---
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=30)
        elements = []
        styles = getSampleStyleSheet()
        
        # FIXED STYLES TO PREVENT OVERLAP
        title_style = ParagraphStyle(
            'Title', 
            fontSize=28, 
            alignment=TA_CENTER, 
            fontName='Helvetica-Bold', 
            leading=34,       
            spaceAfter=12     
        )
        
        sub_title_style = ParagraphStyle(
            'Sub', 
            fontSize=18, 
            alignment=TA_CENTER, 
            textColor=colors.grey, 
            leading=22, 
            spaceAfter=20
        )
        
        cell_style = ParagraphStyle('Cell', fontSize=9, leading=11, alignment=TA_LEFT)
        
        # --- CORRECT TIME FOR INDIA (IST) ---
        # Adjusting for UTC to IST (+5 hours 30 minutes)
        ist_time = datetime.utcnow() + timedelta(hours=5, minutes=30)
        
        # Add Header
        elements.append(Paragraph("Ferrite Agencies", title_style))
        elements.append(Paragraph("Order Report", sub_title_style))
        elements.append(Paragraph(f"Generated on: {ist_time.strftime('%d-%m-%Y %I:%M %p')}", styles['Normal']))
        elements.append(Spacer(1, 20))
        
        # Table Data
        table_data = [['MRP', 'CATEGORY', 'ITEM NAME', 'UNIT', 'QTY', 'FREE QTY']]
        t_qty, t_free = 0, 0
        
        for _, row in df.iterrows():
            t_qty += row['Quantity']
            t_free += row['Free_Quantity']
            
            mrp_disp = f"{row['MRP']:.2f}" if row['MRP'] != 0 else ""
            qty_disp = int(row['Quantity']) if row['Quantity'].is_integer() else f"{row['Quantity']:.2f}"
            free_disp = int(row['Free_Quantity']) if row['Free_Quantity'] > 0 else ""
            
            table_data.append([
                mrp_disp,
                Paragraph(row['Category'], cell_style),
                Paragraph(row['Item Name'], cell_style),
                row['Unit'],
                qty_disp,
                free_disp
            ])
            
        table_data.append(['', '', Paragraph('TOTAL ITEMS', cell_style), '', 
                           int(t_qty) if t_qty.is_integer() else t_qty, 
                           int(t_free) if t_free.is_integer() else t_free])
        
        # FIXED COLUMN WIDTHS & ALIGNMENT
        t = Table(table_data, colWidths=[50, 85, 185, 65, 45, 55], repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2c3e50")), 
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), 
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 9),
            ('TOPPADDING', (0,0), (-1,-1), 6),    
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.whitesmoke, colors.white]),
            ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
            ('BACKGROUND', (0,-1), (-1,-1), colors.lightgrey),
        ]))
        
        elements.append(t)
        doc.build(elements)
        
        st.success(f"PDF Generated Successfully at {ist_time.strftime('%I:%M %p')} IST!")
        st.download_button(
            label="📩 DOWNLOAD FINAL PDF REPORT",
            data=buffer.getvalue(),
            file_name=f"Ferrite_Order_{ist_time.strftime('%H%M%S')}.pdf",
            mime="application/pdf"
        )

    except Exception as e:
        st.error(f"Error: {e}")
