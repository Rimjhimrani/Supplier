import streamlit as st
import pandas as pd
import os
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph, PageBreak
from reportlab.lib.units import cm, inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from io import BytesIO
import subprocess
import sys
import tempfile
from datetime import datetime

# Define label dimensions (similar to your original)
LABEL_WIDTH = 18 * cm  # Wider for shipping label
LABEL_HEIGHT = 12 * cm  # Shorter height
LABEL_PAGESIZE = (LABEL_WIDTH, LABEL_HEIGHT)

# Check for required libraries
try:
    import qrcode
    from reportlab.graphics.barcode import code128
    BARCODE_AVAILABLE = True
except ImportError:
    BARCODE_AVAILABLE = False
    st.write("Installing required libraries...")
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'qrcode', 'reportlab'])
    import qrcode
    from reportlab.graphics.barcode import code128
    BARCODE_AVAILABLE = True

def generate_barcode(data_string, width=4*cm, height=1*cm):
    """Generate Code128 barcode"""
    try:
        from reportlab.graphics.shapes import Drawing
        from reportlab.graphics.barcode import code128
        
        # Create barcode
        barcode = code128.Code128(data_string, barWidth=0.8, barHeight=height)
        
        # Create drawing
        d = Drawing(width, height)
        d.add(barcode)
        
        return d
    except Exception as e:
        st.error(f"Error generating barcode: {e}")
        return None

def find_column_by_keywords(df_columns, keywords):
    """Find column by matching keywords"""
    cols = [str(col).upper() for col in df_columns]
    
    for keyword in keywords:
        for i, col in enumerate(cols):
            if keyword.upper() in col:
                return df_columns[i]
    return None

def generate_shipping_labels(excel_file_path, output_pdf_path, status_callback=None):
    """Generate shipping labels from Excel data"""
    if status_callback:
        status_callback(f"Processing file: {excel_file_path}")
    
    # Load the Excel data
    try:
        if excel_file_path.lower().endswith('.csv'):
            df = pd.read_csv(excel_file_path)
        else:
            try:
                df = pd.read_excel(excel_file_path)
            except Exception as e:
                df = pd.read_excel(excel_file_path, engine='openpyxl')
        
        if status_callback:
            status_callback(f"Successfully read file with {len(df)} rows")
    except Exception as e:
        error_msg = f"Error reading file: {e}"
        if status_callback:
            status_callback(error_msg)
        return None
    
    # Define column mappings based on the label format
    column_mappings = {
        'asn_no': find_column_by_keywords(df.columns, ['ASN', 'ASN_NO', 'ASN NO', 'ADVANCE_SHIPMENT']),
        'document_date': find_column_by_keywords(df.columns, ['DATE', 'DOC_DATE', 'DOCUMENT_DATE', 'SHIP_DATE']),
        'part_no': find_column_by_keywords(df.columns, ['PART', 'PART_NO', 'PART NO', 'ITEM']),
        'description': find_column_by_keywords(df.columns, ['DESC', 'DESCRIPTION', 'ITEM_DESC', 'PART_DESC']),
        'quantity': find_column_by_keywords(df.columns, ['QTY', 'QUANTITY', 'QTY_SHIPPED']),
        'net_weight': find_column_by_keywords(df.columns, ['NET_WT', 'NET_WEIGHT', 'NET WEIGHT']),
        'gross_weight': find_column_by_keywords(df.columns, ['GROSS_WT', 'GROSS_WEIGHT', 'GROSS WEIGHT']),
        'shipper': find_column_by_keywords(df.columns, ['SHIPPER', 'VENDOR', 'SUPPLIER', 'FROM']),
        'receiver': find_column_by_keywords(df.columns, ['RECEIVER', 'CUSTOMER', 'TO', 'CONSIGNEE'])
    }
    
    # Show detected columns
    if status_callback:
        for key, col in column_mappings.items():
            if col:
                status_callback(f"Detected {key}: {col}")
    
    # Create document
    doc = SimpleDocTemplate(output_pdf_path, pagesize=LABEL_PAGESIZE,
                          topMargin=0.5*cm, bottomMargin=0.5*cm,
                          leftMargin=0.5*cm, rightMargin=0.5*cm)
    
    all_elements = []
    
    # Process each row
    for index, row in df.iterrows():
        if status_callback:
            status_callback(f"Creating label {index+1} of {len(df)}")
        
        elements = []
        
        # Extract data with fallbacks
        asn_no = str(row.get(column_mappings['asn_no'], f"ASN{2024070100 + index}"))
        document_date = str(row.get(column_mappings['document_date'], datetime.now().strftime('%d-%m-%Y')))
        part_no = str(row.get(column_mappings['part_no'], f"PART{index+1}"))
        description = str(row.get(column_mappings['description'], "Description"))
        quantity = str(row.get(column_mappings['quantity'], "1"))
        net_weight = str(row.get(column_mappings['net_weight'], "480 KG"))
        gross_weight = str(row.get(column_mappings['gross_weight'], "500 KG"))
        shipper = str(row.get(column_mappings['shipper'], "Shipper Name"))
        receiver = str(row.get(column_mappings['receiver'], "Pinnacle Mobility Solutions Pvt Ltd"))
        
        # Header with Receiver
        header_style = ParagraphStyle(name='Header', fontName='Helvetica-Bold', fontSize=12, alignment=TA_CENTER)
        receiver_para = Paragraph(f"<b>Receiver</b><br/>{receiver}", header_style)
        
        # Main table structure
        main_table_data = [
            # Header row
            [receiver_para, "", ""],
            
            # ASN No and Document Date row
            [
                Paragraph("<b>ASN No.</b>", ParagraphStyle(name='Label', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT)),
                Paragraph(asn_no, ParagraphStyle(name='Value', fontName='Helvetica', fontSize=10, alignment=TA_CENTER)),
                [
                    Paragraph("<b>Document Date</b>", ParagraphStyle(name='Label', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT)),
                    Paragraph(document_date, ParagraphStyle(name='Value', fontName='Helvetica', fontSize=10, alignment=TA_CENTER))
                ]
            ],
            
            # ASN No Barcode
            ["", generate_barcode(asn_no) if BARCODE_AVAILABLE else Paragraph(asn_no, ParagraphStyle(name='Barcode', fontName='Courier', fontSize=8, alignment=TA_CENTER)), ""],
            
            # Part No and Description
            [
                Paragraph("<b>Part No / Desc</b>", ParagraphStyle(name='Label', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT)),
                Paragraph(f"{part_no}", ParagraphStyle(name='Value', fontName='Helvetica', fontSize=10, alignment=TA_CENTER)),
                Paragraph(description, ParagraphStyle(name='Value', fontName='Helvetica', fontSize=10, alignment=TA_CENTER))
            ],
            
            # Part No Barcode
            ["", generate_barcode(part_no) if BARCODE_AVAILABLE else Paragraph(part_no, ParagraphStyle(name='Barcode', fontName='Courier', fontSize=8, alignment=TA_CENTER)), ""],
            
            # Quantity
            [
                Paragraph("<b>Quantity</b>", ParagraphStyle(name='Label', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT)),
                Paragraph(quantity, ParagraphStyle(name='Value', fontName='Helvetica', fontSize=10, alignment=TA_CENTER)),
                ""
            ],
            
            # Quantity Barcode
            ["", generate_barcode(quantity) if BARCODE_AVAILABLE else Paragraph(quantity, ParagraphStyle(name='Barcode', fontName='Courier', fontSize=8, alignment=TA_CENTER)), ""],
            
            # Net Weight and Gross Weight
            [
                [
                    Paragraph("<b>Net Wt.</b>", ParagraphStyle(name='Label', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT)),
                    Paragraph(net_weight, ParagraphStyle(name='Value', fontName='Helvetica', fontSize=10, alignment=TA_CENTER))
                ],
                "",
                [
                    Paragraph("<b>Gross Wt.</b>", ParagraphStyle(name='Label', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT)),
                    Paragraph(gross_weight, ParagraphStyle(name='Value', fontName='Helvetica', fontSize=10, alignment=TA_CENTER))
                ]
            ],
            
            # Shipper
            [
                Paragraph("<b>Shipper</b>", ParagraphStyle(name='Label', fontName='Helvetica-Bold', fontSize=10, alignment=TA_LEFT)),
                Paragraph("V12345", ParagraphStyle(name='Value', fontName='Helvetica', fontSize=10, alignment=TA_CENTER)),
                Paragraph(shipper, ParagraphStyle(name='Value', fontName='Helvetica', fontSize=10, alignment=TA_CENTER))
            ]
        ]
        
        # Create the main table
        main_table = Table(main_table_data,
                          colWidths=[4*cm, 8*cm, 6*cm],
                          rowHeights=[1.5*cm, 1*cm, 1*cm, 1*cm, 1*cm, 1*cm, 1*cm, 1.5*cm, 1*cm])
        
        # Style the table
        main_table.setStyle(TableStyle([
            # Grid
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            
            # Header styling
            ('SPAN', (0, 0), (2, 0)),  # Merge header cells
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            
            # Background colors for label cells
            ('BACKGROUND', (0, 1), (0, 1), colors.lightgrey),
            ('BACKGROUND', (0, 3), (0, 3), colors.lightgrey),
            ('BACKGROUND', (0, 5), (0, 5), colors.lightgrey),
            ('BACKGROUND', (0, 8), (0, 8), colors.lightgrey),
            
            # Font sizes
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ]))
        
        elements.append(main_table)
        
        # Add to all elements
        all_elements.extend(elements)
        
        # Add page break except for last item
        if index < len(df) - 1:
            all_elements.append(PageBreak())
    
    # Build the PDF
    try:
        doc.build(all_elements)
        if status_callback:
            status_callback(f"PDF generated successfully: {output_pdf_path}")
        return output_pdf_path
    except Exception as e:
        error_msg = f"Error building PDF: {e}"
        if status_callback:
            status_callback(error_msg)
        return None

def main():
    """Main Streamlit application"""
    st.set_page_config(page_title="Shipping Label Generator", page_icon="📦", layout="wide")
    
    st.title("📦 Shipping Label Generator")
    st.markdown(
        "<p style='font-size:18px; font-style:italic; margin-top:-10px; text-align:left;'>"
        "Designed and Developed by Agilomatrix</p>",
        unsafe_allow_html=True
    )

    st.markdown("---")
    
    # File upload
    st.header("📁 File Upload")
    uploaded_file = st.file_uploader(
        "Choose an Excel or CSV file",
        type=['xlsx', 'xls', 'csv'],
        help="Upload your Excel or CSV file containing shipping information"
    )
    
    if uploaded_file is not None:
        # Create temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            temp_input_path = tmp_file.name
        
        # Display file info
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        
        # Preview data
        try:
            if uploaded_file.name.lower().endswith('.csv'):
                preview_df = pd.read_csv(temp_input_path).head(5)
            else:
                preview_df = pd.read_excel(temp_input_path).head(5)
            
            st.subheader("📊 Data Preview (First 5 rows)")
            st.dataframe(preview_df, use_container_width=True)
            
        except Exception as e:
            st.error(f"Error previewing file: {e}")
            return
        
        # Generate labels section
        st.subheader("🚀 Generate Shipping Labels")
        
        if st.button("📦 Generate PDF Labels", type="primary", use_container_width=True):
            # Create progress container
            progress_container = st.empty()
            status_container = st.empty()
            
            # Create temporary output file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_output:
                temp_output_path = tmp_output.name
            
            # Progress tracking
            def update_status(message):
                status_container.info(f"📊 {message}")
            
            try:
                # Generate the PDF
                update_status("Starting label generation...")
                
                result_path = generate_shipping_labels(
                    temp_input_path, 
                    temp_output_path,
                    status_callback=update_status
                )
                
                if result_path:
                    # Success - provide download
                    with open(result_path, 'rb') as pdf_file:
                        pdf_data = pdf_file.read()
                    
                    status_container.success("✅ Labels generated successfully!")
                    
                    # Download button
                    st.download_button(
                        label="📥 Download PDF Labels",
                        data=pdf_data,
                        file_name=f"shipping_labels_{uploaded_file.name.split('.')[0]}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                    
                    # Show file size
                    file_size = len(pdf_data) / 1024  # KB
                    st.info(f"📄 PDF size: {file_size:.1f} KB")
                    
                else:
                    status_container.error("❌ Failed to generate labels")
                    
            except Exception as e:
                status_container.error(f"❌ Error: {str(e)}")
                st.exception(e)
            
            finally:
                # Cleanup temporary files
                try:
                    if os.path.exists(temp_input_path):
                        os.unlink(temp_input_path)
                    if os.path.exists(temp_output_path):
                        os.unlink(temp_output_path)
                except:
                    pass
        
        # Label information
        st.subheader("ℹ️ Label Information")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **Label Features:**
            - 📦 Shipping label format
            - 📊 Barcode generation
            - 📍 Receiver information
            - 🔢 ASN tracking
            - ⚖️ Weight information
            """)
        
        with col2:
            st.markdown("""
            **Expected Columns:**
            - ASN No / ASN_NO
            - Document Date / DATE
            - Part No / PART_NO
            - Description / DESC
            - Quantity / QTY
            - Net Weight / NET_WT
            - Gross Weight / GROSS_WT
            - Shipper / VENDOR
            - Receiver / CUSTOMER
            """)
    
    else:
        # Instructions
        st.info("👆 Please upload an Excel or CSV file to get started")
        
        st.subheader("📋 Instructions")
        st.markdown("""
        1. **Upload your file** - Excel (.xlsx, .xls) or CSV format
        2. **Review data preview** - Check if your data looks correct
        3. **Generate labels** - Click the button to create your shipping labels
        4. **Download** - Get your professional shipping labels with barcodes
        """)
        
        # Sample data format
        st.subheader("📊 Sample Data Format")
        sample_data = pd.DataFrame({
            'ASN_NO': ['2024070100', '2024070101', '2024070102'],
            'Document_Date': ['11-07-2024', '11-07-2024', '11-07-2024'],
            'Part_No': ['1234567890', '9876543210', '1122334455'],
            'Description': ['Chassis Harness', 'Engine Component', 'Brake Assembly'],
            'Quantity': [100, 50, 75],
            'Net_Weight': ['480 KG', '250 KG', '350 KG'],
            'Gross_Weight': ['500 KG', '270 KG', '380 KG'],
            'Shipper': ['XYZ Pvt Ltd', 'ABC Corp', 'DEF Industries'],
            'Receiver': ['Pinnacle Mobility Solutions Pvt Ltd', 'Pinnacle Mobility Solutions Pvt Ltd', 'Pinnacle Mobility Solutions Pvt Ltd']
        })
        st.dataframe(sample_data, use_container_width=True)

if __name__ == "__main__":
    main()
