import streamlit as st
import openpyxl
from PIL import Image, ImageDraw, ImageFont
import io
import base64
import zipfile
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import tempfile
import os
import json
import time
import qrcode
import re
import uuid
import pandas as pd

# Set page config
st.set_page_config(
    page_title="🤖CertificateSaathi.AI",
    page_icon="🎓",
    layout="wide",
)

# Initialize session state variables
if 'text_elements' not in st.session_state:
    st.session_state.text_elements = []
if 'template_size' not in st.session_state:
    st.session_state.template_size = None
if 'template_file' not in st.session_state:
    st.session_state.template_file = None
if 'excel_headers' not in st.session_state:
    st.session_state.excel_headers = []
if 'excel_df' not in st.session_state:
    st.session_state.excel_df = None
if 'certificates_generated' not in st.session_state:
    st.session_state.certificates_generated = False
if 'certificate_files' not in st.session_state:
    st.session_state.certificate_files = {}
if 'active_tab' not in st.session_state:
    st.session_state.active_tab = 0
if 'errors' not in st.session_state:
    st.session_state.errors = []
if 'drag_update_data' not in st.session_state:
    st.session_state.drag_update_data = None
if 'email_sent_status' not in st.session_state:
    st.session_state.email_sent_status = {}

# Function to handle drag updates
def process_drag_update():
    if st.session_state.drag_update_data:
        try:
            data = json.loads(st.session_state.drag_update_data)
            # Update position in session state
            for i, element in enumerate(st.session_state.text_elements):
                if element['id'] == data['elementId']:
                    element['x_pos'] = data['xPos']
                    element['y_pos'] = data['yPos']
                    element['actual_x'] = int(st.session_state.template_size[0] * data['xPos'] / 100)
                    element['actual_y'] = int(st.session_state.template_size[1] * data['yPos'] / 100)
            # Clear the update data
            st.session_state.drag_update_data = None
        except Exception as e:
            st.session_state.errors.append(f"Error processing drag update: {str(e)}")

# Function to validate email
def is_valid_email(email):
    if not email or pd.isna(email):
        return False
    pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(pattern, str(email)) is not None

# Function to test email connection
def test_email_connection(email, password):
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(email, password)
        return True, "Connection successful"
    except smtplib.SMTPAuthenticationError:
        return False, "Authentication failed. Check your email and app password."
    except smtplib.SMTPException as e:
        return False, f"SMTP error: {str(e)}"
    except Exception as e:
        return False, f"Connection error: {str(e)}"

# Function to send an email with attachment
def send_email(sender_email, password, recipient_email, subject, body, attachment_path=None):
    try:
        # Validate inputs
        if not all([sender_email, password, recipient_email]):
            return False, "Missing required email parameters"
            
        if not is_valid_email(sender_email) or not is_valid_email(recipient_email):
            return False, "Invalid email address format"

        # Create message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        
        # Add body text
        msg.attach(MIMEText(body, 'plain'))
        
        # Add attachment if provided
        if attachment_path and os.path.exists(attachment_path):
            attachment_filename = os.path.basename(attachment_path)
            with open(attachment_path, 'rb') as file:
                attachment = MIMEApplication(file.read(), Name=attachment_filename)
                attachment['Content-Disposition'] = f'attachment; filename="{attachment_filename}"'
                msg.attach(attachment)
        
        # Create secure SSL context and send email
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.send_message(msg)
        return True, "Email sent successfully"
    except smtplib.SMTPAuthenticationError:
        return False, "Authentication failed. Check your email and app password."
    except Exception as e:
        return False, f"Error sending email: {str(e)}"

# Clean dataframe to ensure proper types
def clean_dataframe(df):
    for col in df.columns:
        # Convert all columns to string to avoid Arrow serialization issues
        try:
            df[col] = df[col].astype(str)
        except:
            df[col] = df[col].apply(str)
    return df

# Process any pending drag updates
process_drag_update()

# Title and description
st.title("🏆🎓CertificateSaathi.AI🤖 ")
st.subheader("Generate & Send Bulk Certificates Instantly!🏆⚡🎯")
st.markdown("""
"Create Stunning Certificates in Seconds! Upload your Excel participant list and certificate template, 
easily position text fields with drag-and-drop, and download all certificates at once. Quick, easy, and professional!
""")

# Show errors if any
if st.session_state.errors:
    with st.expander("Error Log", expanded=True):
        for error in st.session_state.errors:
            st.error(error)
        if st.button("Clear Errors", key="clear_errors"):
            st.session_state.errors = []
            st.rerun()

# Create tabs for workflow
tab_names = [
    "1. Upload Files",
    "2. Design Certificate",
    "3. Generate & Send",
    "5. Certificate Analytics",
    "6. Customization Library",
    "7. Certificate Personalization"
]
tabs = st.tabs(tab_names)

# Tab 1: Upload Files
with tabs[0]:
    st.header("Upload Your Files")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Upload certificate template
        st.subheader("Certificate Template")
        template_file = st.file_uploader("Upload certificate template image", type=["png", "jpg", "jpeg"], key="template_uploader")
        
        if template_file:
            try:
                # Display the template
                image = Image.open(template_file)
                st.session_state.template_size = image.size
                st.session_state.template_file = template_file
                st.image(image, caption=f"Certificate Template ({image.width}x{image.height} px)")
                st.success(f"✅ Template loaded: {image.width}x{image.height} pixels")
            except Exception as e:
                error_msg = f"Error loading template image: {str(e)}"
                st.session_state.errors.append(error_msg)
                st.error(error_msg)
    
    with col2:
        # Upload Excel file
        st.subheader("Participant Data")
        excel_file = st.file_uploader("Upload Excel file with participant data", type=["xlsx", "xls"], key="excel_uploader")
        
        if excel_file:
            try:
                # Load Excel data using pandas for better type handling
                df = pd.read_excel(excel_file, engine='openpyxl')
                
                # Clean the dataframe
                df = clean_dataframe(df)
                st.session_state.excel_df = df
                st.session_state.excel_headers = df.columns.tolist()
                
                # Show Excel preview
                st.dataframe(df.head(), use_container_width=True)
                st.success(f"✅ Excel loaded: {len(df)} participants, {len(df.columns)} columns")
                
                # Email configuration
                st.subheader("Email Configuration (Optional)")
                email_col = st.selectbox("Select email column", options=["None"] + df.columns.tolist(), key="email_column_select")
                
                if email_col != "None":
                    st.session_state.email_column = email_col
                    col_a, col_b = st.columns(2)
                    with col_a:
                        sender_email = st.text_input("Your email address", 
                                              value=st.session_state.get('sender_email', ''),
                                              key="sender_email_input")
                        if sender_email and not is_valid_email(sender_email):
                            st.warning("Please enter a valid email address")
                        else:
                            st.session_state.sender_email = sender_email
                    
                    with col_b:
                        st.session_state.email_password = st.text_input(
                            "App password", 
                            value=st.session_state.get('email_password', ''),
                            type="password",
                            key="email_password_input",
                            help="To send emails using Gmail, you'll need an App Password. This is a 16-character code that gives permission to apps. Get it from Google Account > Security > 2-Step Verification > App passwords."
                        )
                    
                    # Test connection button
                    if st.session_state.get('sender_email') and st.session_state.get('email_password'):
                        if st.button("Test Email Connection", key="test_email_connection"):
                            success, message = test_email_connection(
                                st.session_state.sender_email,
                                st.session_state.email_password
                            )
                            if success:
                                st.success(message)
                            else:
                                st.error(message)
                                st.info("If using Gmail, ensure you've set up an App Password: [Learn how](https://support.google.com/accounts/answer/185833)")
                        
            except Exception as e:
                error_msg = f"Error loading Excel file: {str(e)}"
                st.session_state.errors.append(error_msg)
                st.error(error_msg)
    
    # Navigation buttons
    if st.session_state.get('template_size') and st.session_state.excel_headers:
        if st.button("Continue to Design ➡️", use_container_width=True, key="continue_to_design"):
            st.session_state.active_tab = 1
            st.rerun()

# Validate email column
if 'email_column' in st.session_state:
    email_col = st.session_state.email_column
    invalid_emails = st.session_state.excel_df[email_col].apply(lambda x: not is_valid_email(x) if pd.notna(x) else True)
    if invalid_emails.any():
        st.error("Invalid email addresses found in the Excel file. Please correct them and try again.")

# Tab 2: Design Certificate
with tabs[1]:
    st.header("Design Your Certificate")
    
    # Check if files are loaded
    if not st.session_state.get('template_size') or not st.session_state.excel_headers:
        st.warning("⚠️ Please upload your template and Excel file in the previous tab first")
    else:
        col1, col2 = st.columns([1, 2])
        
        with col1:
            st.subheader("Add Text Fields")
            
            # Field selection
            field = st.selectbox("Select field from Excel", options=st.session_state.excel_headers, key="field_select")
            
            # Text properties
            font_size = st.slider("Font size", min_value=10, max_value=100, value=36, key="font_size_slider")
            color = st.color_picker("Text color", "#000000", key="color_picker")
            
            # Add field button
            if st.button("Add Field", use_container_width=True, key="add_field_button"):
                # Generate unique ID for this element
                element_id = str(uuid.uuid4())
                
                # Default position at center
                x_pos = 50
                y_pos = 50
                
                # Calculate actual coordinates
                actual_x = int(st.session_state.template_size[0] * x_pos / 100)
                actual_y = int(st.session_state.template_size[1] * y_pos / 100)
                
                # Add to session state with width and height for dragging
                st.session_state.text_elements.append({
                    'field': field,
                    'font_size': font_size,
                    'color': color,
                    'x_pos': x_pos,
                    'y_pos': y_pos,
                    'actual_x': actual_x,
                    'actual_y': actual_y,
                    'width': font_size * len(field) * 0.6,  # Approximate width
                    'height': font_size * 1.2,  # Approximate height
                    'id': element_id
                })
                st.success(f"✅ Added field: {field}")
                st.rerun()
            
            # List current fields
            st.subheader("Current Fields")
            if not st.session_state.text_elements:
                st.info("No fields added yet")
            else:
                for i, element in enumerate(st.session_state.text_elements):
                    col_a, col_b, col_c = st.columns([2, 1, 1])
                    with col_a:
                        st.write(f"**{element['field']}** (Size: {element['font_size']})")
                    with col_b:
                        if st.button("Edit", key=f"edit_{i}"):
                            st.session_state.editing_element = i
                            st.rerun()
                    with col_c:
                        if st.button("🗑️", key=f"del_{i}"):
                            st.session_state.text_elements.pop(i)
                            st.rerun()
                    
                    # Show edit form if editing this element
                    if st.session_state.get('editing_element') == i:
                        with st.form(f"edit_element_{i}"):
                            new_font_size = st.slider("Font size", 
                                                    min_value=10, 
                                                    max_value=100,
                                                    value=element['font_size'],
                                                    key=f"edit_size_{i}")
                            
                            new_color = st.color_picker("Text color", 
                                                      element['color'],
                                                      key=f"edit_color_{i}")
                            
                            new_x_pos = st.slider("X position (%)", 
                                               min_value=0, 
                                               max_value=100,
                                               value=element['x_pos'],
                                               key=f"edit_x_{i}")
                            
                            new_y_pos = st.slider("Y position (%)", 
                                               min_value=0, 
                                               max_value=100,
                                               value=element['y_pos'],
                                               key=f"edit_y_{i}")
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                if st.form_submit_button("Save Changes"):
                                    # Update element properties
                                    element['font_size'] = new_font_size
                                    element['color'] = new_color
                                    element['x_pos'] = new_x_pos
                                    element['y_pos'] = new_y_pos
                                    element['actual_x'] = int(st.session_state.template_size[0] * new_x_pos / 100)
                                    element['actual_y'] = int(st.session_state.template_size[1] * new_y_pos / 100)
                                    element['width'] = new_font_size * len(element['field']) * 0.6
                                    element['height'] = new_font_size * 1.2
                                    del st.session_state.editing_element
                                    st.rerun()
                            with col2:
                                if st.form_submit_button("Cancel"):
                                    del st.session_state.editing_element
                                    st.rerun()
            
            # Clear all button
            if st.session_state.text_elements:
                if st.button("Clear All Fields", use_container_width=True, key="clear_all_fields"):
                    st.session_state.text_elements = []
                    st.rerun()
        
        with col2:
            st.subheader("Certificate Preview")
            
            # Load the template
            if st.session_state.template_file:
                try:
                    # Create preview image
                    image = Image.open(st.session_state.template_file)
                    draw = ImageDraw.Draw(image)
                    
                    # Draw each text element
                    for element in st.session_state.text_elements:
                        x = element['actual_x']
                        y = element['actual_y']
                        
                        # Use default font for preview
                        try:
                            font = ImageFont.truetype("arial.ttf", element['font_size'])
                        except:
                            try:
                                font = ImageFont.truetype("Arial.ttf", element['font_size'])
                            except:
                                font = ImageFont.load_default()
                        
                        # Draw text
                        draw.text((x, y), element['field'], fill=element['color'], font=font, anchor="mm")
                    
                    # Convert to base64 for the canvas
                    img_buffer = io.BytesIO()
                    image.save(img_buffer, format="PNG")
                    img_data = base64.b64encode(img_buffer.getvalue()).decode("utf-8")
                    
                    # Show preview with interactive canvas
                    st.markdown("### Drag fields to position them on the certificate")
                    
                    # Create an empty container for receiving dragged element data
                    drag_receiver = st.empty()
                    
                    # Create interactive canvas with HTML/JS
                    canvas_html = f"""
                    <div style="position: relative; margin-bottom: 20px;">
                        <img src="data:image/png;base64,{img_data}" style="width: 100%; max-width: 800px;" id="certificate-img">
                        <div id="canvas-container" style="position: absolute; top: 0; left: 0; width: 100%; height: 100%;">
                            <!-- Elements will be dynamically added here -->
                        </div>
                    </div>
                    
                    <script>
                    // Wait for the image to load
                    document.getElementById('certificate-img').onload = function() {{
                        const container = document.getElementById('canvas-container');
                        container.style.width = this.offsetWidth + 'px';
                        container.style.height = this.offsetHeight + 'px';
                        
                        // Get current elements from JSON
                        const elements = {json.dumps(st.session_state.text_elements)};
                        const imageWidth = {st.session_state.template_size[0]};
                        const imageHeight = {st.session_state.template_size[1]};
                        const displayWidth = this.offsetWidth;
                        const displayHeight = this.offsetHeight;
                        
                        // Scale factors
                        const scaleX = displayWidth / imageWidth;
                        const scaleY = displayHeight / imageHeight;
                        
                        // Function to update Streamlit
                        function updateStreamlit(data) {{
                            const hiddenInput = document.createElement('input');
                            hiddenInput.type = 'text';
                            hiddenInput.id = 'drag-data-input';
                            hiddenInput.style.position = 'absolute';
                            hiddenInput.style.visibility = 'hidden';
                            hiddenInput.value = JSON.stringify(data);
                            
                            // Create a form to submit the data
                            const form = document.createElement('form');
                            form.method = 'POST';
                            form.appendChild(hiddenInput);
                            
                            // Add the form to body and submit
                            document.body.appendChild(form);
                            
                            // Create a submit button and click it
                            const submitBtn = document.createElement('input');
                            submitBtn.type = 'submit';
                            form.appendChild(submitBtn);
                            submitBtn.click();
                            
                            // Remove the form after submission
                            setTimeout(() => form.remove(), 100);
                        }}
                        
                        // Add each element
                        elements.forEach((el) => {{
                            // Create draggable div
                            const elem = document.createElement('div');
                            elem.id = 'element-' + el.id;
                            elem.className = 'draggable-element';
                            elem.style.position = 'absolute';
                            elem.style.cursor = 'move';
                            elem.style.backgroundColor = 'rgba(255,255,255,0.3)';
                            elem.style.border = '1px dashed #333';
                            elem.style.padding = '5px';
                            elem.style.borderRadius = '4px';
                            elem.style.fontSize = Math.max(10, el.font_size * scaleY * 0.8) + 'px';
                            elem.style.color = el.color;
                            
                            // Calculate element width and height in display scale
                            const elemWidth = el.width * scaleX;
                            const elemHeight = el.height * scaleY;
                            
                            // Position the element (center-based)
                            const left = (el.x_pos * displayWidth / 100) - (elemWidth / 2);
                            const top = (el.y_pos * displayHeight / 100) - (elemHeight / 2);
                            
                            elem.style.left = left + 'px';
                            elem.style.top = top + 'px';
                            elem.style.minWidth = '50px';
                            elem.style.textAlign = 'center';
                            elem.innerHTML = el.field;
                            
                            // Add element to container
                            container.appendChild(elem);
                            
                            // Make draggable
                            let isDragging = false;
                            let offsetX, offsetY;
                            
                            elem.addEventListener('mousedown', function(e) {{
                                isDragging = true;
                                offsetX = e.clientX - this.getBoundingClientRect().left;
                                offsetY = e.clientY - this.getBoundingClientRect().top;
                                this.style.zIndex = 1000;
                                e.preventDefault(); // Prevent text selection
                            }});
                            
                            document.addEventListener('mousemove', function(e) {{
                                if (!isDragging) return;
                                
                                // Calculate new position
                                const containerRect = container.getBoundingClientRect();
                                const x = e.clientX - offsetX - containerRect.left;
                                const y = e.clientY - offsetY - containerRect.top;
                                
                                // Keep within bounds
                                const boundedX = Math.max(0, Math.min(container.offsetWidth - 50, x));
                                const boundedY = Math.max(0, Math.min(container.offsetHeight - 20, y));
                                
                                elem.style.left = boundedX + 'px';
                                elem.style.top = boundedY + 'px';
                            }});
                            
                            document.addEventListener('mouseup', function() {{
                                if (isDragging) {{
                                    isDragging = false;
                                    elem.style.zIndex = 'auto';
                                    
                                    // Calculate center position of the element
                                    const elemRect = elem.getBoundingClientRect();
                                    const containerRect = container.getBoundingClientRect();
                                    
                                    // Calculate center of element in container coordinates
                                    const centerX = (elemRect.left - containerRect.left) + (elemRect.width / 2);
                                    const centerY = (elemRect.top - containerRect.top) + (elemRect.height / 2);
                                    
                                    // Convert to percentage
                                    const xPos = Math.round((centerX / container.offsetWidth) * 100);
                                    const yPos = Math.round((centerY / container.offsetHeight) * 100);
                                    
                                    // Send update to Streamlit
                                    updateStreamlit({{
                                        elementId: el.id,
                                        xPos: xPos,
                                        yPos: yPos
                                    }});
                                }}
                            }});
                        }});
                    }};
                    </script>
                    """
                    
                    st.components.v1.html(canvas_html, height=600)
                    
                    # Check if form was submitted with drag data
                    if "drag-data-input" in st.query_params:
                        try:
                            drag_data = st.query_params["drag-data-input"]
                            st.session_state.drag_update_data = drag_data
                            st.rerun()
                        except Exception as e:
                            st.session_state.errors.append(f"Error handling drag update: {str(e)}")
                        
                except Exception as e:
                    error_msg = f"Error generating preview: {str(e)}"
                    st.session_state.errors.append(error_msg)
                    st.error(error_msg)
            else:
                st.info("Please upload a template image in the previous tab")
        
        # Navigation
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("⬅️ Back to Upload", use_container_width=True, key="back_to_upload"):
                st.session_state.active_tab = 0
                st.rerun()
        with col_b:
            if len(st.session_state.text_elements) > 0:
                if st.button("Continue to Generate ➡️", use_container_width=True, key="continue_to_generate"):
                    st.session_state.active_tab = 2
                    st.rerun()
            else:
                st.button("Add fields to continue", disabled=True, use_container_width=True, key="disabled_continue")

# Tab 3: Generate & Send
with tabs[2]:
    st.header("Generate & Send Certificates")
    
    # Check if design is ready
    if not st.session_state.get('template_size') or not st.session_state.excel_headers or not st.session_state.text_elements:
        st.warning("⚠️ Please complete the previous steps first")
    else:
        # Generate button
        if st.button("🎓 GENERATE CERTIFICATES", type="primary", use_container_width=True, key="generate_certificates"):
            with st.spinner("Generating certificates..."):
                try:
                    # Load template and Excel
                    template_image = Image.open(st.session_state.template_file)
                    df = st.session_state.excel_df
                    
                    if df is None:
                        st.error("Excel data not found. Please go back to the first tab and upload it again.")
                    else:
                        # Create temporary directory
                        temp_dir = tempfile.mkdtemp()
                        
                        # Create a zip file in memory
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                            # Generate each certificate
                            certificate_count = 0
                            cert_files = {}
                            st.session_state.email_sent_status = {}  # Reset sent status
                            
                            # Progress bar
                            progress_bar = st.progress(0)
                            total_rows = len(df)
                            
                            for idx, row in df.iterrows():
                                # Update progress
                                progress_percent = min(int((idx+1) / total_rows * 100), 100)
                                progress_bar.progress(progress_percent)
                                
                                # Skip empty rows
                                if row.isnull().all():
                                    continue
                                    
                                # Create a copy of the template
                                certificate = template_image.copy()
                                draw = ImageDraw.Draw(certificate)
                                
                                # Add all fields
                                for element in st.session_state.text_elements:
                                    field_name = element['field']
                                    if field_name in df.columns:
                                        text = str(row[field_name]) if pd.notna(row[field_name]) else ""
                                        
                                        # Use best available font
                                        try:
                                            font = ImageFont.truetype("arial.ttf", element['font_size'])
                                        except:
                                            try:
                                                font = ImageFont.truetype("Arial.ttf", element['font_size'])
                                            except:
                                                font = ImageFont.load_default()
                                        
                                        # Draw text
                                        draw.text(
                                            (element['actual_x'], element['actual_y']),
                                            text,
                                            fill=element['color'],
                                            font=font,
                                            anchor="mm"
                                        )
                                
                                # Save the certificate
                                filename = f"certificate_{idx+1}.png"
                                
                                # Use name if available (assume first column is name)
                                if pd.notna(row[0]):
                                    name = str(row[0]).replace(" ", "_")
                                    filename = f"{name}_certificate.png"
                                
                                # Save to temporary file
                                cert_path = os.path.join(temp_dir, filename)
                                certificate.save(cert_path)
                                
                                # Add to zip
                                zip_file.write(cert_path, filename)
                                
                                # Store for email
                                if 'email_column' in st.session_state:
                                    email = row[st.session_state.email_column] if st.session_state.email_column in row and pd.notna(row[st.session_state.email_column]) else None
                                    if email and is_valid_email(email):
                                        cert_files[email] = cert_path
                                        st.session_state.email_sent_status[email] = False  # Initialize as not sent
                                
                                certificate_count += 1
                            
                            st.session_state.certificate_files = cert_files
                        
                        # Complete the progress bar
                        progress_bar.progress(100)
                        
                        # Prepare download link
                        zip_buffer.seek(0)
                        b64 = base64.b64encode(zip_buffer.read()).decode()
                        href = f'<a href="data:application/zip;base64,{b64}" download="certificates.zip" class="download-button">📥 Download All Certificates</a>'
                        
                        # Success message
                        st.success(f"✅ Successfully generated {certificate_count} certificates!")
                        st.markdown(href, unsafe_allow_html=True)
                        
                        # Set flag for email section
                        st.session_state.certificates_generated = True  # Ensure this is set

                except Exception as e:
                    error_msg = f"Error generating certificates: {str(e)}"
                    st.session_state.errors.append(error_msg)
                    st.error(error_msg)

        # Email section
        if 'email_column' in st.session_state and st.session_state.certificates_generated:
            st.header("Send Certificates by Email")
            
            if not st.session_state.get('sender_email') or not st.session_state.get('email_password'):
                st.warning("⚠️ Please enter your email credentials in the first tab to send emails")
            else:
                # Email settings
                with st.form("email_form"):
                    subject = st.text_input("Email Subject", "Your Certificate", key="email_subject")
                    
                    # Email body
                    email_body = st.text_area(
                        "Email Body",
                        """Dear Participant,

Please find attached your certificate.

Best regards,
Certificate Team""",
                        key="email_body"
                    )
                    
                    # Test email option
                    test_email = st.text_input("Send test email to", 
                                              placeholder="Enter email for testing",
                                              key="test_email_input",
                                              help="Send a test email to verify your configuration")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        test_submit = st.form_submit_button("Send Test Email")
                    with col2:
                        submit_all = st.form_submit_button("Send All Emails")
                
                # Process test email
                if test_submit and test_email:
                    if not is_valid_email(test_email):
                        st.error("Please enter a valid email address")
                    else:
                        # Try sending test email
                        with st.spinner("Sending test email..."):
                            try:
                                # Create test certificate
                                test_cert_path = None
                                if st.session_state.template_file:
                                    # Generate simple test certificate
                                    template_image = Image.open(st.session_state.template_file)
                                    test_cert = template_image.copy()
                                    draw = ImageDraw.Draw(test_cert)
                                    
                                    # Add test text
                                    for element in st.session_state.text_elements:
                                        try:
                                            font = ImageFont.truetype("arial.ttf", element['font_size'])
                                        except:
                                            try:
                                                font = ImageFont.truetype("Arial.ttf", element['font_size'])
                                            except:
                                                font = ImageFont.load_default()
                                        
                                        draw.text(
                                            (element['actual_x'], element['actual_y']),
                                            f"Sample {element['field']}",
                                            fill=element['color'],
                                            font=font,
                                            anchor="mm"
                                        )
                                    
                                    # Save test certificate
                                    test_cert_path = os.path.join(tempfile.gettempdir(), "test_certificate.png")
                                    test_cert.save(test_cert_path)
                                
                                # Send test email
                                success, message = send_email(
                                    st.session_state.sender_email,
                                    st.session_state.email_password,
                                    test_email,
                                    subject,
                                    email_body,
                                    test_cert_path
                                )
                                
                                if success:
                                    st.success(f"✅ Test email sent to {test_email}")
                                else:
                                    st.error(message)
                                    st.info("If using Gmail, ensure you've set up an App Password: [Learn how](https://support.google.com/accounts/answer/185833)")
                                
                            except Exception as e:
                                error_msg = f"Error sending test email: {str(e)}"
                                st.session_state.errors.append(error_msg)
                                st.error(error_msg)
                
                # Process all emails
                if submit_all:
                    if not st.session_state.certificate_files:
                        st.error("No certificates generated yet or no valid email addresses found")
                    else:
                        # Send emails
                        with st.spinner(f"Sending {len(st.session_state.certificate_files)} emails..."):
                            try:
                                # Progress bar
                                progress_bar = st.progress(0)
                                sent_count = 0
                                failed_emails = []
                                EMAIL_SEND_DELAY = 1  # seconds between emails
                                
                                for idx, (email, cert_path) in enumerate(st.session_state.certificate_files.items()):
                                    try:
                                        # Skip if already sent
                                        if st.session_state.email_sent_status.get(email, False):
                                            continue
                                            
                                        # Update progress
                                        progress_percent = min(int((idx+1) / len(st.session_state.certificate_files) * 100), 100)
                                        progress_bar.progress(progress_percent)
                                        
                                        if not is_valid_email(email):
                                            failed_emails.append((email, "Invalid email format"))
                                            continue
                                            
                                        # Send email with certificate
                                        success, message = send_email(
                                            st.session_state.sender_email,
                                            st.session_state.email_password,
                                            email,
                                            subject,
                                            email_body,
                                            cert_path
                                        )
                                        
                                        if success:
                                            sent_count += 1
                                            st.session_state.email_sent_status[email] = True
                                        else:
                                            failed_emails.append((email, message))
                                        
                                        # Prevent rate limiting
                                        if idx < len(st.session_state.certificate_files) - 1:
                                            time.sleep(EMAIL_SEND_DELAY)
                                        
                                    except Exception as e:
                                        failed_emails.append((email, str(e)))
                                
                                # Complete progress bar
                                progress_bar.progress(100)
                                
                                # Log failed emails
                                if failed_emails:
                                    st.warning(f"✅ Sent {sent_count} out of {len(st.session_state.certificate_files)} emails. {len(failed_emails)} failed.")
                                    with st.expander("Failed Emails"):
                                        for email, error in failed_emails:
                                            st.write(f"- {email}: {error}")
                                else:
                                    st.success(f"✅ Successfully sent all {sent_count} emails!")
                                
                            except Exception as e:
                                error_msg = f"Error sending emails: {str(e)}"
                                st.session_state.errors.append(error_msg)
                                st.error(error_msg)

        # Navigation
        if st.button("⬅️ Back to Design", use_container_width=True, key="back_to_design"):
            st.session_state.active_tab = 1
            st.rerun()

# Tab 5: Certificate Analytics
with tabs[3]:
    st.header("Certificate Analytics")
    
    # Check if certificates have been generated
    if not st.session_state.get('certificates_generated', False):
        st.warning("No certificates have been generated yet. Please complete the previous steps.")
    else:
        # Display key statistics
        total_certificates = len(st.session_state.certificate_files)
        total_emails = len(st.session_state.email_sent_status)
        sent_emails = sum(status for status in st.session_state.email_sent_status.values())
        failed_emails = total_emails - sent_emails
        
        st.subheader("Summary")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Certificates", total_certificates)
        with col2:
            st.metric("Emails Sent", sent_emails)
        with col3:
            st.metric("Failed Emails", failed_emails)
        
        # Display detailed table
        st.subheader("Detailed Report")
        if st.session_state.email_sent_status:
            report_data = [
                {"Email": email, "Status": "Sent" if status else "Failed"}
                for email, status in st.session_state.email_sent_status.items()
            ]
            report_df = pd.DataFrame(report_data)
            st.dataframe(report_df, use_container_width=True)
        else:
            st.info("No email data available.")
        
        # Option to retry failed emails
        if failed_emails > 0:
            if st.button("Retry Failed Emails"):
                with st.spinner("Retrying failed emails..."):
                    failed_emails_list = [
                        email for email, status in st.session_state.email_sent_status.items() if not status
                    ]
                    for email in failed_emails_list:
                        cert_path = st.session_state.certificate_files.get(email)
                        if cert_path:
                            success, message = send_email(
                                st.session_state.sender_email,
                                st.session_state.email_password,
                                email,
                                "Your Certificate",
                                "Please find your certificate attached.",
                                cert_path
                            )
                            if success:
                                st.session_state.email_sent_status[email] = True
                            else:
                                st.error(f"Failed to send email to {email}: {message}")
                    st.success("Retry process completed.")

# Tab 6: QR Code Validation
with tabs[4]:
    st.header("QR Code Validation")

    # Define folder to save certificates with QR codes
    qr_cert_folder = "certificates_with_qr/"
    if not os.path.exists(qr_cert_folder):
        os.makedirs(qr_cert_folder)  # Create folder if it doesn't exist

    # Generate QR codes for existing certificates
    st.subheader("Generate QR Codes for Certificates")
    if st.session_state.get('certificate_files', {}):
        if st.button("Generate QR Codes", key="generate_qr_codes_button"):
            with st.spinner("Generating QR codes for certificates..."):
                try:
                    for email, cert_path in st.session_state.certificate_files.items():
                        # Load the certificate
                        certificate = Image.open(cert_path)
                        draw = ImageDraw.Draw(certificate)

                        # Generate a unique QR code
                        certificate_id = f"cert_{uuid.uuid4()}"
                        qr_data = f"https://your-validation-url.com/validate?cert_id={certificate_id}"
                        qr_code = qrcode.make(qr_data)

                        # Resize and paste the QR code onto the certificate
                        qr_code_size = (150, 150)  # Adjust size as needed
                        qr_code = qr_code.resize(qr_code_size)
                        certificate.paste(qr_code, (certificate.width - 160, certificate.height - 160))  # Adjust position

                        # Save the updated certificate with QR code
                        qr_cert_path = os.path.join(qr_cert_folder, os.path.basename(cert_path))
                        certificate.save(qr_cert_path)

                    st.success("QR codes generated and added to certificates successfully!")
                except Exception as e:
                    st.error(f"Error generating QR codes: {str(e)}")
    else:
        st.info("No certificates found. Please generate certificates first in Tab 3.")

    # Preview certificates with QR codes
    st.subheader("Preview Certificates with QR Codes")
    qr_certificates = [f for f in os.listdir(qr_cert_folder) if f.endswith(('.png', '.jpg', '.jpeg'))]
    if qr_certificates:
        selected_qr_cert = st.selectbox("Select a certificate to preview", qr_certificates, key="preview_qr_cert_select")
        if selected_qr_cert:
            qr_cert_path = os.path.join(qr_cert_folder, selected_qr_cert)
            st.image(qr_cert_path, caption=f"Preview: {selected_qr_cert}")
    else:
        st.info("No certificates with QR codes found. Generate QR codes first.")

    # Download certificates with QR codes
    st.subheader("Download Certificates with QR Codes")
    if qr_certificates:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
            for qr_cert in qr_certificates:
                qr_cert_path = os.path.join(qr_cert_folder, qr_cert)
                zip_file.write(qr_cert_path, qr_cert)
        zip_buffer.seek(0)
        b64 = base64.b64encode(zip_buffer.read()).decode()
        href = f'<a href="data:application/zip;base64,{b64}" download="certificates_with_qr.zip" class="download-button">📥 Download All Certificates with QR Codes</a>'
        st.markdown(href, unsafe_allow_html=True)
    else:
        st.info("No certificates with QR codes available for download.")

# Add custom CSS for better styling
st.markdown("""
<style>
    /* Custom styling for download button */
    .download-button {
        display: inline-block;
        padding: 12px 20px;
        background-color: #4CAF50;
        color: white !important;
        text-decoration: none;
        border-radius: 4px;
        font-weight: bold;
        margin: 20px 0;
        text-align: center;
        transition: background-color 0.3s;
    }
    .download-button:hover {
        background-color: #45a049;
    }
    
    /* Make text elements prettier */
    h1 {
        color: #1E88E5;
    }
    h2, h3 {
        color: #0D47A1;
    }
    
    /* Better spacing */
    .block-container {
        padding-top: 2rem;
    }
    
    /* Fix for drag-and-drop elements */
    .draggable-element:hover {
        background-color: rgba(255,255,255,0.5) !important;
    }
    
    /* Better form elements */
    .stTextInput input, .stTextArea textarea {
        border: 1px solid #ddd !important;
    }
    
    /* Better buttons */
    .stButton>button {
        transition: all 0.3s;
    }
    .stButton>button:hover {
        transform: translateY(-1px);
        box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    }
</style>
""", unsafe_allow_html=True)

# Show correct tab on load
if st.session_state.active_tab > 0:
    # This script forces the appropriate tab to be selected
    tab_js = f"""
    <script>
        setTimeout(function() {{
            // Click on the appropriate tab
            document.querySelectorAll('button[data-baseweb="tab"]')[{st.session_state.active_tab}].click();
        }}, 100);
    </script>
    """
    
    st.components.v1.html(tab_js, height=0)

# Tab 7: Certificate Personalization
with tabs[5]:  # Adjusted index to match the correct tab
    st.header("Certificate Personalization")

    # Form for participant details
    st.subheader("Enter Your Details")
    participant_name = st.text_input("Your Name", key="participant_name_input")
    participant_email = st.text_input("Your Email", key="participant_email_input")
    participant_photo = st.file_uploader("Upload Your Photo (Optional)", type=["png", "jpg", "jpeg"], key="participant_photo_uploader")

    # Real-time certificate preview
    st.subheader("Certificate Preview")
    if st.session_state.get('template_file'):
        try:
            # Load the certificate template
            template_image = Image.open(st.session_state.template_file)
            certificate = template_image.copy()
            draw = ImageDraw.Draw(certificate)

            # Add sliders for photo position (after certificate is defined)
            st.subheader("Adjust Photo Position")
            photo_x = st.slider("Photo X Position", min_value=0, max_value=certificate.width, value=50, step=10, key="photo_x_slider")
            photo_y = st.slider("Photo Y Position", min_value=0, max_value=certificate.height, value=certificate.height - 200, step=10, key="photo_y_slider")

            # Add participant name
            if participant_name.strip():
                font_size = 48
                try:
                    font = ImageFont.truetype("arial.ttf", font_size)
                except:
                    font = ImageFont.load_default()
                
                text_x = certificate.width // 2
                text_y = certificate.height // 2
                draw.text((text_x, text_y), participant_name, fill="black", font=font, anchor="mm")

            # Add participant email
            if participant_email.strip():
                font_size = 36
                try:
                    font = ImageFont.truetype("arial.ttf", font_size)
                except:
                    font = ImageFont.load_default()
                
                text_x = certificate.width // 2
                text_y = certificate.height // 2 + 60  # Adjust position below the name
                draw.text((text_x, text_y), participant_email, fill="black", font=font, anchor="mm")

            # Add participant photo
            if participant_photo:
                photo = Image.open(participant_photo)
                photo = photo.resize((150, 150))  # Resize photo
                certificate.paste(photo, (photo_x, photo_y))  # Use sliders for position

            # Display the preview
            st.image(certificate, caption="Live Certificate Preview", use_column_width=True)

            # Generate and download the certificate
            if st.button("Download Certificate", key="download_participant_certificate"):
                cert_buffer = io.BytesIO()
                certificate.save(cert_buffer, format="PNG")
                cert_buffer.seek(0)

                # Provide download link
                b64 = base64.b64encode(cert_buffer.read()).decode()
                href = f'<a href="data:image/png;base64,{b64}" download="{participant_name}_certificate.png" class="download-button">📥 Download Your Certificate</a>'
                st.markdown(href, unsafe_allow_html=True)
                st.success("Your certificate has been generated successfully!")
        except Exception as e:
            st.error(f"Error generating certificate preview: {str(e)}")
    else:
        st.error("No certificate template found. Please upload a template in Tab 1.")
        
# Convert the image to Base64
def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

# Embed the developer's picture and name at the bottom center
image_base64 = get_base64_image("pic.png")
st.markdown(f"""
    <div style="position: fixed; bottom: 40px; width: 100%; text-align: center; font-size: 15px; color: grey;">
        <img src="data:image/jpeg;base64,{image_base64}" alt="Developer" style="width: 60px; height: 60px; border-radius: 50%; margin-bottom: 5px;">
        <br>
        Developer: <strong>Abhishek Yadav</strong>
    </div>
""", unsafe_allow_html=True)



# Convert the logo image to Base64
def get_base64_logo(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

# Embed the app logo at the top center
logo_base64 = get_base64_logo("image.png")
st.markdown(f"""
    <div style="text-align: center; margin-bottom: 20px;">
        <img src="data:image/png;base64,{logo_base64}" alt="App Logo" style="width: 120px; height: auto; border-radius: 10px; box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.1);">
        <h1 style="margin-top: 10px; color: #1E88E5;">CertificateSaathi.AI</h1>
    </div>
""", unsafe_allow_html=True)
