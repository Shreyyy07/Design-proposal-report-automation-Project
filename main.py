import streamlit as st
import pandas as pd
import pdfplumber
import cv2
import tempfile
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from pptx import Presentation
from pptx.util import Inches
import os
import subprocess
import pyautogui
import pygetwindow as gw
import time
import shutil
import traceback


# ---------- PDF Extraction ----------
def extract_pdf_info(pdf_path):
    content = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    content.append(text)
        return content[:3]
    except Exception as e:
        return [f"PDF error: {e}"]

# ---------- Excel Reading ----------
def read_excel_data(excel_path):
    try:
        return pd.read_excel(excel_path)
    except Exception as e:
        st.error(f"Excel error: {e}")
        return pd.DataFrame()

# ---------- Wear Image Processing ----------
def analyze_wear_image(image_path):
    try:
        img = cv2.imread(image_path, 0)
        _, thresh = cv2.threshold(img, 127, 255, cv2.THRESH_BINARY)
        contours, _ = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
        return len(contours)
    except Exception as e:
        st.error(f"Image error: {e}")
        return 0

# ---------- Open AutoCAD, Take Screenshot, and Close AutoCAD ----------
def open_autocad_and_capture_screenshot(dwg_path):
    try:
        # Start AutoCAD process
        autocad_process = subprocess.Popen(["C:/Program Files/Autodesk/AutoCAD 2023/acad.exe", dwg_path])

        st.info("Opening AutoCAD... Waiting for file to load.")
        time.sleep(15)  # Wait for AutoCAD to open and load the DWG

        # Maximize AutoCAD window
        autocad_window = None
        for w in gw.getWindowsWithTitle('AutoCAD'):
            if not w.isMaximized:
                w.maximize()
            w.activate()
            autocad_window = w
            break

        time.sleep(2)

        # Send "Zoom Extents" command
        pyautogui.press('esc')
        time.sleep(1)
        pyautogui.typewrite('zoom\n')
        time.sleep(1)
        pyautogui.typewrite('e\n')
        time.sleep(5)  # Wait more after zoom

        # Take full screen screenshot
        screenshot = pyautogui.screenshot()

        # Crop the central area (adjust cropping if needed)
        screen_width, screen_height = pyautogui.size()
        left = screen_width * 0.15
        top = screen_height * 0.10
        right = screen_width * 0.85
        bottom = screen_height * 0.90

        cropped_screenshot = screenshot.crop((left, top, right, bottom))

        # Save cropped screenshot
        output_image_path = os.path.join(tempfile.gettempdir(), "cad_screenshot.png")
        cropped_screenshot.save(output_image_path)

        # Close AutoCAD after capturing
        if autocad_window:
            try:
                # Send ALT+F4 to close AutoCAD
                autocad_window.activate()
                pyautogui.hotkey('alt', 'f4')
                time.sleep(1)
                # Handle any save dialog by pressing 'n' for No
                pyautogui.press('n')
            except Exception as e:
                st.warning(f"Could not gracefully close AutoCAD: {e}")
                # Force kill the process if graceful close failed
                try:
                    autocad_process.kill()
                except:
                    pass

        return output_image_path

    except Exception as e:
        st.error(f"Error opening AutoCAD or capturing screenshot: {e}")
        return None
    
def open_nx_and_capture_views(prt_file_path, manual_file_open=True):
    try:
        # Launch NX with optimized startup
        nx_path = r"D:\abcde\NXBIN\ugraf.exe"  # Adjust this to your NX installation path
        nx_process = subprocess.Popen([nx_path], shell=True)
        st.info("Opening NX... Please wait for the application to load.")
        time.sleep(20)  # Reduced wait time for NX to initialize
        
        # Find and activate the NX window - more efficient window detection
        nx_window = None
        attempts = 0
        while attempts < 3 and not nx_window:  # Reduced attempts
            for w in gw.getWindowsWithTitle('NX'):
                if 'NX' in w.title:
                    if not w.isMaximized:
                        w.maximize()
                    w.activate()
                    nx_window = w
                    break
            if not nx_window:
                time.sleep(3)  # Reduced wait time
                attempts += 1
                
        if not nx_window:
            st.error("Could not find NX window. Please check if NX launched properly.")
            return {}

        # Dismiss any startup dialogs
        pyautogui.press('escape')
        time.sleep(1)  # Reduced wait time
        
        # Click on empty area and ensure NX is active
        screen_width, screen_height = pyautogui.size()
        pyautogui.click(screen_width//2, screen_height//2)
        nx_window.activate()

        # Display instructions for manual file opening
        st.warning(f"Please manually open the file: {prt_file_path}")
        st.info("After opening the file, the script will automatically capture views and close NX.")
        
        # Wait for the user to manually open the file
        time.sleep(15)  # Reduced wait time
        
        st.info("Now capturing different views...")
        
        # Simplified view method using function keys (most reliable across NX versions)
        function_keys = {
            'top': 'f3',
            'front': 'f7',
            'right': 'f6',
            'isometric': 'f9'
        }
        
        # Capture only essential views to reduce time
        screenshots = {}
        
        # Capture each view
        for view, key in function_keys.items():
            nx_window.activate()
            pyautogui.press(key)
            time.sleep(1)  # Wait briefly for view to change
            
            # Fit view to screen
            pyautogui.press('f')
            time.sleep(1)
            
            # Capture screenshot
            screenshot = pyautogui.screenshot()
            img_path = os.path.join(tempfile.gettempdir(), f"nxview_{view}.png")
            screenshot.save(img_path)
            screenshots[view] = img_path
            st.success(f"Captured {view} view")

        # Close NX immediately after capturing
        nx_window.activate()
        pyautogui.hotkey('alt', 'f4')
        time.sleep(1)
        
        # Handle any save dialog by pressing 'n' (No)
        pyautogui.press('n')
        
        # Ensure process is terminated
        try:
            nx_process.terminate()
        except:
            pass

        st.success("Completed NX session and closed application!")
        return screenshots

    except Exception as e:
        st.error(f"Error in NX automation: {e}")
        # Try to close NX forcefully if there was an error
        try:
            for w in gw.getWindowsWithTitle('NX'):
                w.close()
        except:
            pass
        return {}
    
# ---------- PDF Report Generator ----------
def generate_pdf(pdf_info, excel_df, image_path, cad_image_path, output_path):
    doc = SimpleDocTemplate(output_path, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []
    elements.append(Paragraph("<b>Tyre Design Proposal Report</b>", styles["Title"]))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("Extracted PDF Content:", styles["Heading2"]))
    for pg in pdf_info:
        elements.append(Paragraph(pg.replace('\n', '<br/>'), styles["Normal"]))
        elements.append(Spacer(1, 10))

    elements.append(Spacer(1, 20))
    elements.append(Paragraph("Wear Image Analysis:", styles["Heading2"]))
    if image_path:
        elements.append(Image(image_path, width=300, height=200))

    if cad_image_path:
        elements.append(Spacer(1, 20))
        elements.append(Paragraph("2D CAD File Screenshot:", styles["Heading2"]))
        elements.append(Image(cad_image_path, width=300, height=200))

    if not excel_df.empty:
        elements.append(Spacer(1, 20))
        elements.append(Paragraph("Tyre Specifications:", styles["Heading2"]))
        data = [excel_df.columns.tolist()] + excel_df.astype(str).values.tolist()
        table = Table(data)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ]))
        elements.append(table)
    
    # Fixed NX views inclusion - removed duplicate code
    if 'nx_views' in st.session_state:
        elements.append(Spacer(1, 20))
        elements.append(Paragraph("3D Model Views (NX):", styles["Heading2"]))
        for view, path in st.session_state['nx_views'].items():
            elements.append(Paragraph(f"{view.capitalize()} View:", styles["Normal"]))
            elements.append(Image(path, width=300, height=200))

    doc.build(elements)

# ---------- PPTX Report Generator ----------
def generate_pptx(pdf_info, excel_df, image_path, cad_image_path, output_path):
    prs = Presentation()
    slide1 = prs.slides.add_slide(prs.slide_layouts[0])
    slide1.shapes.title.text = "Tyre Design Proposal Report"
    slide1.placeholders[1].text = "Generated using Streamlit + Python"

    slide2 = prs.slides.add_slide(prs.slide_layouts[1])
    slide2.shapes.title.text = "PDF Content (Specs)"
    tf = slide2.placeholders[1].text_frame
    tf.text = "\n--- PAGE SPLIT ---\n".join(pdf_info)

    slide3 = prs.slides.add_slide(prs.slide_layouts[5])
    slide3.shapes.title.text = "Wear Image Analysis"
    if image_path:
        slide3.shapes.add_picture(image_path, Inches(1), Inches(1.5), width=Inches(4.5))

    if cad_image_path:
        slide5 = prs.slides.add_slide(prs.slide_layouts[5])
        slide5.shapes.title.text = "2D CAD File Screenshot"
        slide5.shapes.add_picture(cad_image_path, Inches(1), Inches(1.5), width=Inches(6))

    if not excel_df.empty:
        slide4 = prs.slides.add_slide(prs.slide_layouts[5])
        slide4.shapes.title.text = "Tyre Specifications"
        rows, cols = excel_df.shape
        table = slide4.shapes.add_table(rows+1, cols, Inches(0.5), Inches(1.5), Inches(8.5), Inches(0.8 + rows * 0.3)).table
        for i, col in enumerate(excel_df.columns):
            table.cell(0, i).text = str(col)
        for r in range(rows):
            for c in range(cols):
                table.cell(r+1, c).text = str(excel_df.iloc[r, c])

    # Fixed NX views inclusion - removed duplicate code
    if 'nx_views' in st.session_state:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = "NX 3D Model Views"
        left = Inches(0.5)
        top = Inches(1.5)
        for i, (view, path) in enumerate(st.session_state['nx_views'].items()):
            if os.path.exists(path):
                slide.shapes.add_picture(path, left + Inches((i % 2) * 4.5), top + Inches((i // 2) * 2.5), width=Inches(4))

    prs.save(output_path)

# ---------- Streamlit GUI ----------
st.title("🛞 Tyre Report Generator with PDF & PPT Export (AutoCAD Drawing Capture)")

# Initialize session state variables if they don't exist
if 'cad_screenshot_captured' not in st.session_state:
    st.session_state['cad_screenshot_captured'] = False
    
if 'cad_screenshot_path' not in st.session_state:
    st.session_state['cad_screenshot_path'] = None

if 'cad_file_uploaded' not in st.session_state:
    st.session_state['cad_file_uploaded'] = False
    
if 'processing_cad' not in st.session_state:
    st.session_state['processing_cad'] = False

if 'nx_views' not in st.session_state:
    st.session_state['nx_views'] = {}

# Create tabs with reordered sections: AutoCAD, NX, Reports (as requested)
tab1, tab3, tab2 = st.tabs(["📸 Capture CAD Drawing", "📐 Capture NX 3D Views", "📄 Generate Reports"])

with tab1:
    cad_file = st.file_uploader("Upload 2D CAD File (.dwg)", type=["dwg"], 
                             key="cad_uploader",
                             on_change=lambda: setattr(st.session_state, 'cad_file_uploaded', True))
    
    # This block automatically processes the CAD file when uploaded
    if st.session_state['cad_file_uploaded'] and not st.session_state['processing_cad'] and not st.session_state['cad_screenshot_captured']:
        st.session_state['processing_cad'] = True
        cad_temp_dir = tempfile.mkdtemp()
        cad_path = os.path.join(cad_temp_dir, cad_file.name)

        with open(cad_path, "wb") as f:
            f.write(cad_file.read())

        with st.spinner('Opening AutoCAD and capturing drawing... Please wait...'):
            cad_screenshot_path = open_autocad_and_capture_screenshot(cad_path)

        if cad_screenshot_path:
            st.session_state['cad_screenshot_path'] = cad_screenshot_path
            st.session_state['cad_screenshot_captured'] = True
            st.success("Drawing captured successfully!")
        else:
            st.error("Failed to capture drawing.")
        
        st.session_state['processing_cad'] = False
        st.session_state['cad_file_uploaded'] = False
        # Force refresh to show success message and image
        st.rerun()
    
    # Show the captured screenshot if it exists - this is shown only once
    if st.session_state['cad_screenshot_captured']:
        st.success("✅ CAD drawing has been captured and is ready for report generation")
        with st.container():
            st.image(st.session_state['cad_screenshot_path'], caption="Captured CAD Drawing")
        
        if st.button("Capture Again"):
            st.session_state['cad_screenshot_captured'] = False
            st.session_state['cad_screenshot_path'] = None
            st.rerun()

with tab3:
    st.header("📐 NX 3D View Capture")
    prt_file = st.file_uploader("Upload 3D Model File (.prt or .step)", type=["prt", "step"], key="nx_uploader")

    if st.button("Capture 3D Views", key="nx_capture_btn") and prt_file:
        temp_prt_path = os.path.join(tempfile.gettempdir(), prt_file.name)
        with open(temp_prt_path, "wb") as f:
            f.write(prt_file.read())

        with st.spinner("Opening NX and capturing views... This will close automatically"):
            views = open_nx_and_capture_views(temp_prt_path)

        if views:
            st.success("✅ Captured 3D views from NX successfully!")
            st.session_state['nx_views'] = views
            
            # Show a collapsible section with the captured views
            with st.expander("View Captured Screenshots"):
                for name, path in views.items():
                    st.image(path, caption=f"{name.capitalize()} View")
        else:
            st.error("Failed to capture 3D views from NX.")

with tab2:
    pdf_file = st.file_uploader("Upload PDF File (Specs)", type=["pdf"])
    excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    image_file = st.file_uploader("Upload Wear Image", type=["jpg", "png"])
    
    # Check if CAD drawing has been captured
    if not st.session_state['cad_screenshot_captured']:
        st.warning("⚠️ Please capture a CAD drawing in the 'Capture CAD Drawing' tab before generating reports")
    
    # --------- Generate Report Button ---------
    if st.button("Generate Report"):
        if pdf_file and excel_file and image_file and st.session_state['cad_screenshot_captured']:
            cad_screenshot_path = st.session_state['cad_screenshot_path']
            with tempfile.TemporaryDirectory() as tmpdir:
                pdf_path = os.path.join(tmpdir, pdf_file.name)
                excel_path = os.path.join(tmpdir, excel_file.name)
                img_path = os.path.join(tmpdir, image_file.name)
                pptx_path = os.path.join(tmpdir, "Tyre_Report.pptx")
                pdf_out_path = os.path.join(tmpdir, "Tyre_Report.pdf")

                with open(pdf_path, "wb") as f: f.write(pdf_file.read())
                with open(excel_path, "wb") as f: f.write(excel_file.read())
                with open(img_path, "wb") as f: f.write(image_file.read())

                pdf_info = extract_pdf_info(pdf_path)
                excel_df = read_excel_data(excel_path)
                wear_zones = analyze_wear_image(img_path)
                pdf_info.append(f"Wear Zones Detected: {wear_zones}")

                # Generate reports directly using existing data without reopening NX
                generate_pdf(pdf_info, excel_df, img_path, cad_screenshot_path, pdf_out_path)
                generate_pptx(pdf_info, excel_df, img_path, cad_screenshot_path, pptx_path)

                col1, col2 = st.columns(2)
                with col1:
                    with open(pdf_out_path, "rb") as f:
                        st.download_button("📄 Download PDF Report", f, file_name="Tyre_Report.pdf", mime="application/pdf")

                with col2:
                    with open(pptx_path, "rb") as f:
                        st.download_button("📊 Download PPT Report", f, file_name="Tyre_Report.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
                
                st.success("Reports generated successfully!")
        else:
            if not st.session_state['cad_screenshot_captured']:
                st.error("CAD drawing has not been captured yet. Please go to the 'Capture CAD Drawing' tab first.")
            else:
                st.error("Please ensure all files (PDF, Excel, and Wear Image) are uploaded before generating reports.")