import streamlit as st
import pandas as pd
import io
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfgen import canvas
import tempfile
import os
import barcode
from barcode import Code128
from barcode.writer import ImageWriter
from PIL import Image

# Set page configuration
st.set_page_config(
    page_title="Supplier Label Generator",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for styling
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%);
        color: white;
        padding: 30px;
        text-align: center;
        border-radius: 20px;
        margin-bottom: 30px;
    }
    
    .main-header h1 {
        font-size: 2.5em;
        margin-bottom: 10px;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    
    .subtitle {
        font-size: 1.2em;
        opacity: 0.9;
        font-style: italic;
    }
    
    .upload-section {
        background: #f8f9fa;
        border: 3px dashed #dee2e6;
        border-radius: 15px;
        padding: 40px;
        text-align: center;
        margin-bottom: 30px;
    }
    
    .info-section {
        background: #f8f9fa;
        border-radius: 10px;
        padding: 20px;
        margin-top: 20px;
    }
    
    .info-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 20px;
        margin-top: 20px;
    }
    
    .info-card {
        background: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    
    .info-card h3 {
        color: #2c3e50;
        margin-bottom: 15px;
    }
    
    .success-message {
        background: #d4edda;
        color: #155724;
        border: 1px solid #c3e6cb;
        padding: 15px;
        border-radius: 10px;
        margin: 20px 0;
    }
    
    .error-message {
        background: #f8d7da;
        color: #721c24;
        border: 1px solid #f5c6cb;
        padding: 15px;
        border-radius: 10px;
        margin: 20px 0;
    }
    
    .info-message {
        background: #cce7ff;
        color: #004085;
        border: 1px solid #b8daff;
        padding: 15px;
        border-radius: 10px;
        margin: 20px 0;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class="main-header">
    <h1>📦 Supplier Label Generator</h1>
    <p class="subtitle" style="font-size: 1.4em;">Designed and Developed by Agilomatrix</p>
</div>
""", unsafe_allow_html=True)

# Initialize session state
if 'uploaded_data' not in st.session_state:
    st.session_state.uploaded_data = None
if 'column_mappings' not in st.session_state:
    st.session_state.column_mappings = {}

def detect_columns(headers):
    """Detect column mappings based on header names"""
    mappings = {
        'document_date': ['DATE', 'DOC_DATE', 'DOCUMENT_DATE', 'SHIP_DATE', 'DOCUMENT DATE', 'Document Date'],
        'asn_no': ['ASN', 'ASN_NO', 'ASN NO', 'ADVANCE_SHIPMENT', 'ASN NUMBER', 'ASN No.', 'ASN No'],
        'part_no': ['PART', 'PART_NO', 'PART NO', 'ITEM', 'PART NUMBER', 'PartNo'],
        'description': ['DESC', 'DESCRIPTION', 'ITEM_DESC', 'PART_DESC', 'ITEM DESCRIPTION', 'Part Description', 'Description'],
        'quantity': ['QTY', 'QUANTITY', 'QTY_SHIPPED', 'SHIPPED QTY', 'Quantity', 'Qty'],
        'net_weight': ['NET_WT', 'NET_WEIGHT', 'NET WEIGHT', 'Net Weight(KG)', 'NET WT', 'Net Wt.', 'Net Wt', 'NetWt'],
        'gross_weight': ['GROSS_WT', 'GROSS_WEIGHT', 'GROSS WEIGHT','Gross Weight(KG)', 'GROSS WT', 'GROSS WT.', 'Gross Wt.', 'Gross Wt', 'Gross wt.'],
        'shipper_id': ['SHIPPER_PART', 'VENDOR_PART', 'SUPPLIER_PART', 'VENDOR PART', 'SHIPPER PART', 'Shipper ID', 'ID', 'id', 'Shipper_ID', 'Delivery Partner ID', 'SHIPPER_ID'],
        'shipper_name': ['SHIPPER', 'VENDOR', 'SUPPLIER', 'FROM', 'VENDOR NAME', 'SUPPLIER NAME', 'SHIPPER NAME', 'Shipper Name', 'shipper name', 'SHIPPER_NAME']
    }
    column_mappings = {}
    
    for key, keywords in mappings.items():
        found = None
        for header in headers:
            header_upper = header.upper()
            if any(keyword in header_upper for keyword in keywords):
                found = header
                break
        if found:
            column_mappings[key] = found
    
    return column_mappings

def get_value_with_fallback(row, column_name, default_value, allow_blank=False):
    if not column_name:
        return default_value if not allow_blank else ""
    if column_name in row and pd.notna(row[column_name]):
        value = row[column_name]
        if isinstance(value, pd.Timestamp):
            return value.strftime('%d-%m-%y')
        value_str = str(value).strip()
        return value_str if value_str else ("" if allow_blank else default_value)
    
    return "" if allow_blank else default_value

def draw_centered_text(canvas, text, x, y, width):
    """Helper function to draw centered text"""
    text_width = canvas.stringWidth(text, canvas._fontname, canvas._fontsize)
    center_x = x + width / 2 - text_width / 2
    canvas.drawString(center_x, y, text)

def generate_barcode_image(data, width_cm=3.5, height_cm=0.8):
    """Generate a barcode image"""
    if not data or str(data).strip() == "":
        return None
    try:
        from barcode import Code128
        from barcode.writer import ImageWriter
        
        code128 = Code128(str(data), writer=ImageWriter())
        barcode_buffer = io.BytesIO()
        code128.write(barcode_buffer, options={
            'module_width': 0.2,
            'module_height': 10,
            'quiet_zone': 1,
            'text_distance': 6,
            'font_size': 8,
            'write_text': False
        })
        
        barcode_buffer.seek(0)
        barcode_image = Image.open(barcode_buffer)
        
        temp_barcode = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
        barcode_image.save(temp_barcode.name, 'PNG')
        temp_barcode.close()
        return temp_barcode.name
    except Exception as e:
        print(f"Error generating barcode: {e}")
        return None

def draw_barcode(canvas, data, x, y, width_cm, height_cm):
    """Draw barcode on canvas"""
    # Don't draw barcode for empty data
    if not data or str(data).strip() == "":
        return
        
    barcode_file = generate_barcode_image(data, width_cm, height_cm)
    if barcode_file:
        try:
            canvas.drawImage(barcode_file, x, y, width=width_cm, height=height_cm)
            # Clean up temporary file
            os.unlink(barcode_file)
        except Exception as e:
            # If barcode fails, fall back to text
            draw_centered_text(canvas, str(data), x, y + height_cm/2, width_cm)
    else:
        # For empty data, don't draw anything (leave blank)
        pass

def create_label_pdf(data, column_mappings):
    """Create PDF with shipping labels"""
    # Create a temporary file
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_filename = temp_file.name
    temp_file.close()
    
    # Create PDF with custom page size (10cm x 15cm)
    page_width = 10 * cm
    page_height = 15 * cm
    
    c = canvas.Canvas(temp_filename, pagesize=(page_width, page_height))
    
    for index, row in data.iterrows():
        if index > 0:
            c.showPage()
        
        # Extract data with fallbacks - ASN can be blank
        document_date = get_value_with_fallback(row, column_mappings.get('document_date'), '11-07-24')
        asn_no = get_value_with_fallback(row, column_mappings.get('asn_no'), '', allow_blank=True)
        part_no = get_value_with_fallback(row, column_mappings.get('part_no'), f'PART{index + 1}')
        description = get_value_with_fallback(row, column_mappings.get('description'), 'Description')
        quantity = get_value_with_fallback(row, column_mappings.get('quantity'), '1')
        net_weight = get_value_with_fallback(row, column_mappings.get('net_weight'), '480')
        gross_weight = get_value_with_fallback(row, column_mappings.get('gross_weight'), '500')
        shipper_id = get_value_with_fallback(row, column_mappings.get('shipper_id'), 'V12345')
        shipper_name = get_value_with_fallback(row, column_mappings.get('shipper_name'), 'Shipper Name')
        
        # Create the label with correct mapping
        create_single_label(c, document_date, asn_no, part_no, description, quantity, 
                          net_weight, gross_weight, shipper_id, shipper_name, page_width, page_height)
    
    c.save()
    return temp_filename

def create_single_label(c, document_date, asn_no, part_no, description, quantity, 
                       net_weight, gross_weight, shipper_id, shipper_name, page_width, page_height):
    """Create a single label with correct shipper information"""
    
    # Set up dimensions (same as original)
    row_height = 1.0 * cm
    start_y = page_height - 0.5 * cm - row_height
    
    # Column widths
    col1_width = 2.5 * cm
    col2_width = 3.0 * cm
    col3_width = 3.7 * cm
    
    # Set line width and font - darker border
    c.setLineWidth(1.0)
    c.setFont('Helvetica', 11)
    
    # Row 1: EKA Mobility, Document Date Header, Date Value
    current_y = start_y
    eka_col_width = 5.5 * cm
    doc_header_width = 1.4 * cm
    doc_value_width = 2.3 * cm
    
    # Draw rectangles
    c.rect(0.5 * cm, current_y, eka_col_width, row_height)
    c.rect(0.5 * cm + eka_col_width, current_y, doc_header_width, row_height)
    c.rect(0.5 * cm + eka_col_width + doc_header_width, current_y, doc_value_width, row_height)
    
    # Add text
    c.setFont('Helvetica-Bold', 11)
    center_y = current_y + row_height / 2

    # First line: slightly above the center
    draw_centered_text(c, 'Pinnacle Mobility Solutions', 0.5 * cm, center_y + 0.15 * cm, eka_col_width)
    # Second line: slightly below the center
    draw_centered_text(c, 'Pvt. Ltd.', 0.5 * cm, center_y - 0.25 * cm, eka_col_width)

    c.setFont('Helvetica-Bold', 11)
    draw_centered_text(c, 'Date', 0.5 * cm + eka_col_width, current_y + row_height / 2 - 0.15 * cm, doc_header_width)
    
    c.setFont('Helvetica', 11)
    draw_centered_text(c, document_date, 0.5 * cm + eka_col_width + doc_header_width, current_y + row_height / 2 - 0.15 * cm, doc_value_width)
    
    # Row 2: ASN No Header, ASN Value, Barcode
    current_y -= row_height
    c.rect(0.5 * cm, current_y, col1_width, row_height)
    c.rect(0.5 * cm + col1_width, current_y, col2_width, row_height)
    c.rect(0.5 * cm + col1_width + col2_width, current_y, col3_width, row_height)
    
    c.setFont('Helvetica-Bold', 11)
    draw_centered_text(c, 'ASN No', 0.5 * cm, current_y + row_height / 2 - 0.15 * cm, col1_width)
    c.setFont('Helvetica', 11)
    # Only draw ASN value if it's not empty
    if asn_no and asn_no.strip():
        draw_centered_text(c, asn_no, 0.5 * cm + col1_width, current_y + row_height / 2 - 0.15 * cm, col2_width)
        # Draw barcode for ASN number only if ASN is not empty
        draw_barcode(c, asn_no, 0.5 * cm + col1_width + col2_width + 0.1 * cm, current_y + 0.1 * cm, col3_width - 0.2 * cm, row_height - 0.2 * cm)
    
    # Row 3: Part No Header, Part Value, Barcode
    current_y -= row_height
    c.rect(0.5 * cm, current_y, col1_width, row_height)
    c.rect(0.5 * cm + col1_width, current_y, col2_width, row_height)
    c.rect(0.5 * cm + col1_width + col2_width, current_y, col3_width, row_height)
    
    c.setFont('Helvetica-Bold', 11)
    draw_centered_text(c, 'Part No', 0.5 * cm, current_y + row_height / 2 - 0.15 * cm, col1_width)
    c.setFont('Helvetica', 11)
    draw_centered_text(c, part_no, 0.5 * cm + col1_width, current_y + row_height / 2 - 0.15 * cm, col2_width)
    # Draw barcode for part number
    draw_barcode(c, part_no, 0.5 * cm + col1_width + col2_width + 0.1 * cm, current_y + 0.1 * cm, col3_width - 0.2 * cm, row_height - 0.2 * cm)
    
    # Row 4: Description Header, Description Value
    current_y -= row_height
    c.rect(0.5 * cm, current_y, col1_width, row_height)
    c.rect(0.5 * cm + col1_width, current_y, col2_width + col3_width, row_height)
    
    c.setFont('Helvetica-Bold', 11)
    draw_centered_text(c, 'Description', 0.5 * cm, current_y + row_height / 2 - 0.15 * cm, col1_width)
    c.setFont('Helvetica', 11)
    # Truncate description if too long
    if len(description) > 25:
        description = description[:22] + "..."
    c.drawString(0.5 * cm + col1_width + 0.2 * cm, current_y + row_height / 2 - 0.15 * cm, description)
    
    # Row 5: Quantity Header, Quantity Value, Barcode
    current_y -= row_height
    c.rect(0.5 * cm, current_y, col1_width, row_height)
    c.rect(0.5 * cm + col1_width, current_y, col2_width, row_height)
    c.rect(0.5 * cm + col1_width + col2_width, current_y, col3_width, row_height)
    
    c.setFont('Helvetica-Bold', 11)
    draw_centered_text(c, 'Quantity', 0.5 * cm, current_y + row_height / 2 - 0.15 * cm, col1_width)
    c.setFont('Helvetica', 11)
    draw_centered_text(c, quantity, 0.5 * cm + col1_width, current_y + row_height / 2 - 0.15 * cm, col2_width)
    # Draw barcode for quantity
    draw_barcode(c, quantity, 0.5 * cm + col1_width + col2_width + 0.1 * cm, current_y + 0.1 * cm, col3_width - 0.2 * cm, row_height - 0.2 * cm)
    
    # Row 6: Net Wt Header, Net Wt Value, Gross Wt Header, Gross Wt Value
    current_y -= row_height
    header_width = 2.5 * cm
    value_width = 2.1 * cm
    
    c.rect(0.5 * cm, current_y, header_width, row_height)
    c.rect(0.5 * cm + header_width, current_y, value_width, row_height)
    c.rect(0.5 * cm + header_width + value_width, current_y, header_width, row_height)
    c.rect(0.5 * cm + header_width * 2 + value_width, current_y, value_width, row_height)
    
    c.setFont('Helvetica-Bold', 11)
    draw_centered_text(c, 'Net Wt', 0.5 * cm, current_y + row_height / 2 - 0.15 * cm, header_width)
    c.setFont('Helvetica', 11)
    draw_centered_text(c, net_weight, 0.5 * cm + header_width, current_y + row_height / 2 - 0.15 * cm, value_width)
    c.setFont('Helvetica-Bold', 11)
    draw_centered_text(c, 'Gross Wt', 0.5 * cm + header_width + value_width, current_y + row_height / 2 - 0.15 * cm, header_width)
    c.setFont('Helvetica', 11)
    draw_centered_text(c, gross_weight, 0.5 * cm + header_width * 2 + value_width, current_y + row_height / 2 - 0.15 * cm, value_width)
    
    # Row 7: Shipper Header, Shipper ID, Shipper Name - FIXED
    # Row 7: Shipper Header, Shipper ID, Shipper Name - UPDATED WIDTHS
    current_y -= row_height
    row7_height = 1.0 * cm

    # NEW fixed widths
    shipper_label_width = 2.5 * cm
    shipper_id_width = 2.5 * cm
    shipper_name_width = 4.2 * cm

    # Draw rectangles
    c.rect(0.5 * cm, current_y, shipper_label_width, row7_height)
    c.rect(0.5 * cm + shipper_label_width, current_y, shipper_id_width, row7_height)
    c.rect(0.5 * cm + shipper_label_width + shipper_id_width, current_y, shipper_name_width, row7_height)

    # Add text
    c.setFont('Helvetica-Bold', 11)
    draw_centered_text(c, 'Shipper', 0.5 * cm, current_y + row7_height / 2 - 0.15 * cm, shipper_label_width)

    c.setFont('Helvetica', 11)
    draw_centered_text(c, shipper_id, 0.5 * cm + shipper_label_width, current_y + row7_height / 2 - 0.15 * cm, shipper_id_width)

    # Truncate shipper name if too long
    display_shipper_name = shipper_name
    if len(display_shipper_name) > 15:
        display_shipper_name = display_shipper_name[:12] + "..."
    draw_centered_text(c, display_shipper_name, 0.5 * cm + shipper_label_width + shipper_id_width, current_y + row7_height / 2 - 0.15 * cm, shipper_name_width)


# File upload section
st.markdown("""
<div class="upload-section">
    <div style="font-size: 4em; margin-bottom: 20px;">📁</div>
    <h3>Upload Your Excel or CSV File</h3>
    <p>Choose a file containing your shipping information</p>
    <p><small>Supported formats: .xlsx, .xls, .csv</small></p>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None:
    try:
        # Read the file
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # Clean column names
        df.columns = df.columns.astype(str).str.strip()
        
        # Detect column mappings
        column_mappings = detect_columns(df.columns.tolist())
        st.session_state.column_mappings = column_mappings
        st.session_state.uploaded_data = df
        
        # Show success message
        st.markdown(f"""
        <div class="success-message">
            ✅ File loaded successfully! {len(df)} records found.
        </div>
        """, unsafe_allow_html=True)
        
        # Show preview
        st.subheader("📊 Data Preview")
        st.dataframe(df.head(), use_container_width=True)
        
        # Show column mappings
        st.subheader("🔗 Column Mappings Detected")
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Mapped Columns:**")
            for key, value in column_mappings.items():
                if value:
                    st.write(f"• {key.replace('_', ' ').title()}: `{value}`")
        
        with col2:
            st.write("**Available Columns:**")
            for col in df.columns:
                st.write(f"• {col}")
        
        # Generate PDF button
        if st.button("🚀 Generate PDF Labels", type="primary", use_container_width=True):
            with st.spinner("Generating PDF labels..."):
                try:
                    pdf_file = create_label_pdf(df, column_mappings)
                    
                    # Read the PDF file
                    with open(pdf_file, 'rb') as f:
                        pdf_bytes = f.read()
                    
                    # Clean up temporary file
                    os.unlink(pdf_file)
                    
                    # Provide download button
                    st.download_button(
                        label="📥 Download PDF Labels",
                        data=pdf_bytes,
                        file_name="shipping_labels.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                    
                    st.markdown("""
                    <div class="success-message">
                        ✅ PDF generated successfully! Click the download button above to save your labels.
                    </div>
                    """, unsafe_allow_html=True)
                    
                except Exception as e:
                    st.markdown(f"""
                    <div class="error-message">
                        ❌ Error generating PDF: {str(e)}
                    </div>
                    """, unsafe_allow_html=True)
    
    except Exception as e:
        st.markdown(f"""
        <div class="error-message">
            ❌ Error reading file: {str(e)}
        </div>
        """, unsafe_allow_html=True)

# Information section
st.markdown("""
<div class="info-section">
    <h3>ℹ️ Label Information</h3>
    <div class="info-grid">
        <div class="info-card">
            <h3>Label Features</h3>
            <ul>
                <li>📦 Custom 10cm x 15cm label format</li>
                <li>📊 Real Code128 barcodes (scannable)</li>
                <li>📍 Fixed EKA Mobility header</li>
                <li>🔢 ASN tracking number with barcode (blank if no data)</li>
                <li>🔧 Part number with barcode</li>
                <li>📈 Quantity with barcode</li>
                <li>⚖️ Weight information included</li>
                <li>📋 7-row structured layout</li>
                <li>🔲 Darker borders for better visibility</li>
                <li>🏷️ Shipper ID (column 2) and Name (column 3)</li>
            </ul>
        </div>
        <div class="info-card">
            <h3>Expected Columns</h3>
            <ul>
                <li>Document Date / DATE</li>
                <li>ASN No / ASN_NO (can be blank)</li>
                <li>Part No / PART_NO</li>
                <li>Description / DESC</li>
                <li>Quantity / QTY</li>
                <li>Net Weight / NET_WT</li>
                <li>Gross Weight / GROSS_WT</li>
                <li>Shipper Name / VENDOR</li>
                <li>Shipper ID / SHIPPER_PART</li>
            </ul>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Footer
st.markdown("---")
st.markdown("**Designed and Developed by Agilomatrix**")
