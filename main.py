import streamlit as st
import pandas as pd
import pdfplumber
import cv2
import tempfile
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
import os
import subprocess
import pyautogui
import pygetwindow as gw
import time
import shutil
import base64
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import textwrap


def get_apollo_logo_base64():
    """
    Convert the Apollo Tyres logo to base64 for embedding
    You'll need to save your logo image as 'apollo_logo.png' in the same directory
    """
    try:
        with open("apollo_logo.png", "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except FileNotFoundError:
        # Fallback if logo file not found
        return None

def add_logo_to_streamlit():
    """
    Add Apollo Tyres logo to Streamlit UI at the top center
    """
    logo_base64 = get_apollo_logo_base64()
    if logo_base64:
        st.markdown(
            f"""
            <div style="display: flex; justify-content: center; margin-bottom: 20px;">
                <img src="data:image/png;base64,{logo_base64}" width="200" alt="Apollo Tyres Logo">
            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        # Fallback text logo if image not found
        st.markdown(
            """
            <div style="text-align: center; margin-bottom: 20px;">
                <h2 style="color: #8A2BE2; margin: 0;">🏎️ APOLLO TYRES LTD</h2>
            </div>
            """,
            unsafe_allow_html=True
        )

# ---------- SESSION STATE INITIALIZATION ----------

def initialize_session_state():
    if 'cad_screenshots_captured' not in st.session_state:
        st.session_state['cad_screenshots_captured'] = False
    if 'cad_screenshot_paths' not in st.session_state:
        st.session_state['cad_screenshot_paths'] = []
    if 'nx_screenshots_captured' not in st.session_state:
        st.session_state['nx_screenshots_captured'] = False
    if 'nx_screenshot_paths' not in st.session_state:
        st.session_state['nx_screenshot_paths'] = []

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

# ---------- AutoCAD Functions ----------
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

        # Take full screen screenshot first
        screenshot = pyautogui.screenshot()

        screen_width, screen_height = pyautogui.size()
        
        # More precise cropping to capture only the main drawing area
        # These values focus on the central drawing viewport, excluding toolbars, ribbons, and panels
        left = screen_width * 0.15   # Remove left panel/toolbars more aggressively
        top = screen_height * 0.20   # Remove ribbon, title bar, and top toolbars
        right = screen_width * 0.85  # Remove right panels and properties
        bottom = screen_height * 0.75  # Remove command line, status bar, and bottom elements

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
    
# ---------- NX Functions ----------
def open_nx_and_capture_views_manual(prt_file_path="", manual_file_open=True):
    """
    Enhanced NX automation with precise viewport capture and clear user signals
    """
    try:
        # Launch NX with optimized startup
        nx_path = r"D:\abcde\NXBIN\ugraf.exe"  # Adjust this to your NX installation path
        nx_process = subprocess.Popen([nx_path], shell=True)
        st.info("Opening NX... Please wait for the application to load.")
        time.sleep(20)  # Wait for NX to initialize
        
        # Find and activate the NX window
        nx_window = None
        attempts = 0
        while attempts < 3 and not nx_window:
            for w in gw.getWindowsWithTitle('NX'):
                if 'NX' in w.title:
                    if not w.isMaximized:
                        w.maximize()
                    w.activate()
                    nx_window = w
                    break
            if not nx_window:
                time.sleep(3)
                attempts += 1
                
        if not nx_window:
            st.error("Could not find NX window. Please check if NX launched properly.")
            return {}

        # Dismiss any startup dialogs
        pyautogui.press('escape')
        time.sleep(1)
        
        # Click on empty area and ensure NX is active
        screen_width, screen_height = pyautogui.size()
        pyautogui.click(screen_width//2, screen_height//2)
        nx_window.activate()

        # Display instructions for manual file opening
        st.warning("Please manually open your 3D file in NX now.")
        st.info("After opening the file, follow the instructions below for each view.")
        
        # Wait for the user to manually open the file
        time.sleep(20)  # Increased wait time for manual file opening
        
        # Manual view capture with enhanced user feedback
        views_to_capture = ['Top', 'Front', 'Right', 'Isometric']
        screenshots = {}
        
        # Create placeholders for dynamic updates
        status_placeholder = st.empty()
        progress_placeholder = st.empty()
        
        for i, view in enumerate(views_to_capture):
            # Update status with clear instructions
            status_placeholder.info(f"📍 **Step {i+1}/4**: Please manually rotate/change to **{view}** view in NX")
            progress_placeholder.progress((i) / len(views_to_capture))
            
            # Enhanced countdown with visual feedback
            countdown_placeholder = st.empty()
            for countdown in range(10, 0, -1):
                countdown_placeholder.warning(f"⏰ Capturing {view} view in {countdown} seconds... Please set your view now!")
                time.sleep(1)
            
            countdown_placeholder.empty()
            
            # CAPTURE PREPARATION SIGNAL
            prepare_placeholder = st.empty()
            prepare_placeholder.error("🔴 **PREPARING TO CAPTURE** - Hold your view steady!")
            time.sleep(2)
            prepare_placeholder.empty()
            
            # Ensure NX window is active
            nx_window.activate()
            time.sleep(1)
            
            # Fit view to screen for better capture
            pyautogui.press('f')
            time.sleep(2)
            
            # CAPTURE MOMENT SIGNAL
            capture_placeholder = st.empty()
            capture_placeholder.success("📸 **CAPTURING NOW** - Screenshot being taken!")
            
            # UPDATED: Capture only the 3D viewport area (not entire NX window)
            screenshot = pyautogui.screenshot()

            left = screen_width * 0.12   # Remove left toolbar area
            top = screen_height * 0.12   # Remove ribbon and title bar
            right = screen_width * 0.88  # Remove right panels
            bottom = screen_height * 0.88  # Remove bottom status/command area
            
            cropped_screenshot = screenshot.crop((left, top, right, bottom))
            
            # Save the cropped viewport screenshot
            img_path = os.path.join(tempfile.gettempdir(), f"nxview_{view.lower()}.png")
            cropped_screenshot.save(img_path)
            screenshots[view.lower()] = img_path
            
            capture_placeholder.empty()
            
            # SUCCESS CONFIRMATION SIGNAL with Enhanced Feedback
            success_placeholder = st.empty()
            success_placeholder.success(f"✅ **{view} VIEW CAPTURED SUCCESSFULLY!** 🎉")
            
            # Add a brief visual confirmation
            time.sleep(3)  # Hold success message longer
            success_placeholder.empty()
            
            # Audio-like feedback through rapid status changes (simulating beep)
            if i < len(views_to_capture) - 1:  # Don't show for last capture
                beep_placeholder = st.empty()
                for beep in range(3):
                    beep_placeholder.info("🔊 BEEP!")
                    time.sleep(0.3)
                    beep_placeholder.empty()
                    time.sleep(0.2)
                
                # Pause message before next view
                next_view_placeholder = st.empty()
                next_view_placeholder.warning(f"⏳ Get ready for next view: **{views_to_capture[i+1]}** in 5 seconds...")
                time.sleep(5)
                next_view_placeholder.empty()

        # Final status update with celebration
        status_placeholder.success("🎉 **ALL VIEWS CAPTURED SUCCESSFULLY!** 🎉")
        progress_placeholder.progress(1.0)
        
        # Final success confirmation
        final_placeholder = st.empty()
        final_placeholder.balloons()  # Streamlit balloons animation
        
        # Close NX after capturing all views
        st.info("Closing NX application...")
        nx_window.activate()
        pyautogui.hotkey('alt', 'f4')
        time.sleep(2)
        
        # Handle any save dialog by pressing 'n' (No)
        pyautogui.press('n')
        
        # Ensure process is terminated
        try:
            nx_process.terminate()
        except:
            pass

        st.success("✅ Completed NX session and closed application!")
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

# ---------- PowerPoint Generation Functions ----------
def generate_ppt(pdf_info, excel_df, cad_image_paths, nx_model_groups, output_path):
    """
    Generate PowerPoint presentation with proper content alignment and pagination
    Sequence: PDF Content -> Excel -> AutoCAD -> NX
    """
    # Create presentation
    prs = Presentation()
    
    # Set slide dimensions (16:9 aspect ratio)
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    
    # Title Slide
    create_title_slide(prs)
    
    # 1. PDF Content Slides (FIRST in sequence)
    create_pdf_content_slides(prs, pdf_info)
    
    # 2. Excel Data Slides (SECOND in sequence)
    if not excel_df.empty:
        create_excel_slides(prs, excel_df)
    
    # 3. CAD Screenshots Slides (THIRD in sequence)
    if cad_image_paths:
        create_cad_slides(prs, cad_image_paths)
    
    # 4. NX Model Slides (FOURTH in sequence)
    if nx_model_groups:
        create_nx_model_slides(prs, nx_model_groups)
    
    # Save presentation
    prs.save(output_path)

def create_title_slide(prs):
    """Create title slide with Apollo Tyres branding - logo at top, title centered"""
    # Use blank layout to have full control
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(slide_layout)
    
    # Add logo at the top center if available
    try:
        if os.path.exists("apollo_logo.png"):
            # Logo positioned at top center
            logo_left = Inches(5.5)  # Center horizontally
            logo_top = Inches(0.5)   # Top of slide
            logo_width = Inches(2.33)
            logo_height = Inches(1.17)
            slide.shapes.add_picture("apollo_logo.png", logo_left, logo_top, logo_width, logo_height)
    except:
        # Fallback text logo at top
        logo_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.33), Inches(1))
        logo_frame = logo_box.text_frame
        logo_frame.text = "🏎️ APOLLO TYRES LTD"
        logo_frame.paragraphs[0].font.size = Pt(28)
        logo_frame.paragraphs[0].font.bold = True
        logo_frame.paragraphs[0].font.color.rgb = RGBColor(138, 43, 226)
        logo_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Main title centered on slide
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.33), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.text = "Enhanced Tyre Design Proposal Report"
    title_frame.paragraphs[0].font.size = Pt(44)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].font.color.rgb = RGBColor(138, 43, 226)
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    
    # Subtitle below main title
    subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(11.33), Inches(1.5))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = "Comprehensive Analysis & Design Documentation"
    subtitle_frame.paragraphs[0].font.size = Pt(24)
    subtitle_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

def create_pdf_content_slides(prs, pdf_info):
    """Create slides for PDF content with proper text wrapping and pagination"""
    if not pdf_info:
        return
    
    # Use blank layout to avoid "Click to add title" placeholders
    slide_layout = prs.slide_layouts[6]  # Blank layout
    
    for page_idx, page_content in enumerate(pdf_info):
        # Split content into manageable chunks (max 800 characters per slide)
        content_chunks = split_text_for_slides(page_content, max_chars=800)
        
        for chunk_idx, chunk in enumerate(content_chunks):
            slide = prs.slides.add_slide(slide_layout)
            
            # Add custom title using textbox (no placeholder)
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12.33), Inches(0.8))
            title_frame = title_box.text_frame
            title_text = f"PDF Content - Page {page_idx + 1}"
            if len(content_chunks) > 1:
                title_text += f" (Part {chunk_idx + 1})"
            title_frame.text = title_text
            title_frame.paragraphs[0].font.size = Pt(32)
            title_frame.paragraphs[0].font.bold = True
            title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            
            # Add content using textbox (no placeholder)
            content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(12.33), Inches(5.8))
            content_frame = content_box.text_frame
            
            # Split into paragraphs and add to slide
            paragraphs = chunk.split('\n')
            for para_idx, paragraph in enumerate(paragraphs):
                if paragraph.strip():
                    if para_idx > 0:
                        p = content_frame.add_paragraph()
                    else:
                        p = content_frame.paragraphs[0]
                    
                    # Wrap long lines
                    wrapped_lines = textwrap.wrap(paragraph, width=80)
                    p.text = '\n'.join(wrapped_lines)
                    p.font.size = Pt(14)
                    p.space_after = Pt(6)

def create_cad_slides(prs, cad_image_paths):
    """Create slides for CAD screenshots using blank layout"""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    
    for i, cad_path in enumerate(cad_image_paths):
        if cad_path and os.path.exists(cad_path):
            slide = prs.slides.add_slide(slide_layout)
            
            # Add custom title using textbox
            title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12.33), Inches(0.8))
            title_frame = title_box.text_frame
            title_frame.text = f"2D CAD Drawing {i + 1}"
            title_frame.paragraphs[0].font.size = Pt(32)
            title_frame.paragraphs[0].font.bold = True
            title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            
            # Add image - centered and properly sized
            try:
                img_left = Inches(1.5)
                img_top = Inches(1.5)
                img_width = Inches(10.33)
                img_height = Inches(5.5)
                
                slide.shapes.add_picture(cad_path, img_left, img_top, img_width, img_height)
            except Exception as e:
                # Add error text if image fails to load
                error_box = slide.shapes.add_textbox(Inches(4), Inches(3), Inches(5), Inches(1))
                error_frame = error_box.text_frame
                error_frame.text = f"Error loading CAD image {i + 1}"
                error_frame.paragraphs[0].font.size = Pt(16)
                error_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

def create_nx_model_slides(prs, nx_model_groups):
    """Create slides for NX models with 2x2 grid layout using blank layout"""
    slide_layout = prs.slide_layouts[6]  # Blank layout
    
    for group_idx, model_group in enumerate(nx_model_groups):
        slide = prs.slides.add_slide(slide_layout)
        
        # Add custom title using textbox
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.1), Inches(12.33), Inches(0.7))
        title_frame = title_box.text_frame
        title_frame.text = f"3D NX Model {group_idx + 1} - Multiple Views"
        title_frame.paragraphs[0].font.size = Pt(28)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Create 2x2 grid for up to 4 images
        view_names = ['Top View', 'Front View', 'Right View', 'Isometric View']
        positions = [
            (Inches(0.8), Inches(1.2)),   # Top-left
            (Inches(6.8), Inches(1.2)),   # Top-right
            (Inches(0.8), Inches(4.2)),   # Bottom-left
            (Inches(6.8), Inches(4.2))    # Bottom-right
        ]
        
        img_width = Inches(5.5)
        img_height = Inches(2.8)
        
        for i, (img_path, view_name, (left, top)) in enumerate(zip(model_group[:4], view_names, positions)):
            if os.path.exists(img_path):
                try:
                    # Add image
                    slide.shapes.add_picture(img_path, left, top, img_width, img_height)
                    
                    # Add view label using textbox
                    label_box = slide.shapes.add_textbox(left, top - Inches(0.3), img_width, Inches(0.25))
                    label_frame = label_box.text_frame
                    label_frame.text = view_name
                    label_frame.paragraphs[0].font.size = Pt(12)
                    label_frame.paragraphs[0].font.bold = True
                    label_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                    
                except Exception as e:
                    # Add placeholder if image fails
                    placeholder = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE, left, top, img_width, img_height
                    )
                    placeholder.fill.solid()
                    placeholder.fill.fore_color.rgb = RGBColor(200, 200, 200)
                    
                    # Add error text using textbox
                    error_box = slide.shapes.add_textbox(left, top + Inches(1), img_width, Inches(0.5))
                    error_frame = error_box.text_frame
                    error_frame.text = f"Image Error\n{view_name}"
                    error_frame.paragraphs[0].font.size = Pt(10)
                    error_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

def create_excel_slides(prs, excel_df):
    """Create slides for Excel data with proper table formatting and pagination using blank layout"""
    if excel_df.empty:
        return
    
    slide_layout = prs.slide_layouts[6]  # Blank layout
    
    # Calculate rows per slide (accounting for header and slide dimensions)
    max_rows_per_slide = 12  # Conservative estimate for readability
    
    # Split dataframe into chunks
    total_rows = len(excel_df)
    num_slides = (total_rows + max_rows_per_slide - 1) // max_rows_per_slide
    
    for slide_idx in range(num_slides):
        slide = prs.slides.add_slide(slide_layout)
        
        # Add custom title using textbox
        title_text = "Tyre Specifications"
        if num_slides > 1:
            title_text += f" (Page {slide_idx + 1}/{num_slides})"
            
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12.33), Inches(0.8))
        title_frame = title_box.text_frame
        title_frame.text = title_text
        title_frame.paragraphs[0].font.size = Pt(28)
        title_frame.paragraphs[0].font.bold = True
        title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
        
        # Calculate data slice for this slide
        start_row = slide_idx * max_rows_per_slide
        end_row = min(start_row + max_rows_per_slide, total_rows)
        df_slice = excel_df.iloc[start_row:end_row]
        
        # Create table
        rows = len(df_slice) + 1  # +1 for header
        cols = len(df_slice.columns)
        
        # Add table shape
        table_left = Inches(0.5)
        table_top = Inches(1.2)
        table_width = Inches(12.33)
        table_height = Inches(5.8)
        
        table_shape = slide.shapes.add_table(rows, cols, table_left, table_top, table_width, table_height)
        table = table_shape.table
        
        # Set column widths
        col_width = table_width / cols
        for col_idx in range(cols):
            table.columns[col_idx].width = int(col_width)
        
        # Add header row
        for col_idx, column_name in enumerate(df_slice.columns):
            cell = table.cell(0, col_idx)
            cell.text = str(column_name)
            cell.text_frame.paragraphs[0].font.bold = True
            cell.text_frame.paragraphs[0].font.size = Pt(10)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(138, 43, 226)  # Purple header
            cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)  # White text
        
        # Add data rows
        for row_idx, (_, row_data) in enumerate(df_slice.iterrows()):
            for col_idx, value in enumerate(row_data):
                cell = table.cell(row_idx + 1, col_idx)
                # Truncate long values to fit
                cell_text = str(value)
                if len(cell_text) > 20:
                    cell_text = cell_text[:17] + "..."
                cell.text = cell_text
                cell.text_frame.paragraphs[0].font.size = Pt(9)

def split_text_for_slides(text, max_chars=800):
    """Split long text into chunks suitable for slides"""
    if len(text) <= max_chars:
        return [text]
    
    # Split by paragraphs first
    paragraphs = text.split('\n')
    chunks = []
    current_chunk = ""
    
    for paragraph in paragraphs:
        # If adding this paragraph would exceed the limit, start a new chunk
        if len(current_chunk) + len(paragraph) + 1 > max_chars and current_chunk:
            chunks.append(current_chunk.strip())
            current_chunk = paragraph
        else:
            if current_chunk:
                current_chunk += "\n" + paragraph
            else:
                current_chunk = paragraph
    
    # Add the last chunk
    if current_chunk:
        chunks.append(current_chunk.strip())
    
    return chunks

# ---------- PDF Report Generator ----------
def generate_pdf(pdf_info, excel_df, cad_image_paths, nx_model_groups, output_path):
    """
    Enhanced PDF generation with grid layout for NX screenshots
    """
    doc = SimpleDocTemplate(output_path, pagesize=A4, 
                          topMargin=72, bottomMargin=72, leftMargin=72, rightMargin=72)
    styles = getSampleStyleSheet()
    elements = []

    # Add Apollo Tyres Logo at the top center
    try:
        if os.path.exists("apollo_logo.png"):
            logo = Image("apollo_logo.png", width=150, height=75)
            logo.hAlign = 'CENTER'
            elements.append(logo)
            elements.append(Spacer(1, 20))
        else:
            elements.append(Paragraph("<b>APOLLO TYRES LTD</b>", styles["Title"]))
            elements.append(Spacer(1, 10))
    except Exception as e:
        elements.append(Paragraph("<b>APOLLO TYRES LTD</b>", styles["Title"]))
        elements.append(Spacer(1, 10))

    # Title
    elements.append(Paragraph("<b>Enhanced Tyre Design Proposal Report</b>", styles["Title"]))
    elements.append(Spacer(1, 20))
    
    # PDF Content Section
    elements.append(Paragraph("Extracted PDF Content:", styles["Heading2"]))
    elements.append(Spacer(1, 6))
    for i, pg in enumerate(pdf_info):
        elements.append(Paragraph(pg.replace('\n', '<br/>'), styles["Normal"]))
        if i < len(pdf_info) - 1:
            elements.append(Spacer(1, 6))

    elements.append(Spacer(1, 15))

    # Excel data Section
    if not excel_df.empty:
        elements.append(Paragraph("Tyre Specifications:", styles["Heading2"]))
        elements.append(Spacer(1, 6))
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
        elements.append(Spacer(1, 15))

    # Multiple CAD Screenshots Section
    if cad_image_paths:
        elements.append(Paragraph("2D CAD File Screenshots:", styles["Heading2"]))
        elements.append(Spacer(1, 6))
        
        for i, cad_path in enumerate(cad_image_paths):
            if cad_path and os.path.exists(cad_path):
                elements.append(Paragraph(f"CAD Drawing {i+1}:", styles["Normal"]))
                elements.append(Spacer(1, 3))
                elements.append(Image(cad_path, width=300, height=200))
                elements.append(Spacer(1, 10))

    # Multiple NX Screenshots Section with Grid Layout
    if nx_model_groups:
        elements.append(Paragraph("3D NX File Screenshots:", styles["Heading2"]))
        elements.append(Spacer(1, 6))
        
        for group_idx, model_group in enumerate(nx_model_groups):
            elements.append(Paragraph(f"NX 3D Model {group_idx + 1}:", styles["Normal"]))
            elements.append(Spacer(1, 6))
            
            # Create 2x2 grid for each model group (4 views)
            if len(model_group) >= 4:
                # Create table data for 2x2 grid
                grid_data = []
                
                # Row 1: First two images
                row1 = []
                for i in range(2):
                    if i < len(model_group) and os.path.exists(model_group[i]):
                        img = Image(model_group[i], width=140, height=100)
                        row1.append(img)
                    else:
                        row1.append("")
                grid_data.append(row1)
                
                # Row 2: Next two images
                row2 = []
                for i in range(2, 4):
                    if i < len(model_group) and os.path.exists(model_group[i]):
                        img = Image(model_group[i], width=140, height=100)
                        row2.append(img)
                    else:
                        row2.append("")
                grid_data.append(row2)
                
                # Create table with 2x2 grid
                grid_table = Table(grid_data, colWidths=[150, 150])
                grid_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
                    ('LEFTPADDING', (0, 0), (-1, -1), 5),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                    ('TOPPADDING', (0, 0), (-1, -1), 5),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                ]))
                
                elements.append(grid_table)
            else:
                # Fallback for less than 4 images - display them normally
                for img_path in model_group:
                    if os.path.exists(img_path):
                        elements.append(Image(img_path, width=200, height=150))
                        elements.append(Spacer(1, 5))
            
            elements.append(Spacer(1, 15))

    doc.build(elements)

# ---------- STREAMLIT GUI ----------
def main():
    """
    Main Streamlit application with multiple file support
    """
    # Add logo at the top center
    add_logo_to_streamlit()

    st.title("Enhanced Tyre Report Generator")

    # Initialize session state
    initialize_session_state()

    # Create tabs
    tab1, tab2, tab3 = st.tabs([
        "📸 Capture CAD Drawing", 
        "📐 Capture NX 3D Models", 
        "📄 Generate Reports"
    ])

    with tab1:
        updated_tab1_section()

    with tab2:
        updated_tab2_section()

    with tab3:
        updated_tab3_section()

# ---------- TAB SECTIONS ----------
def updated_tab1_section():
    """
    Updated CAD section with multiple file upload support
    """
    st.header("📸 AutoCAD Drawing Capture")
    cad_files = st.file_uploader("Upload 2D CAD Files (.dwg)", 
                               type=["dwg"], 
                               accept_multiple_files=True,
                               key="cad_uploader_multiple")
    
    if cad_files:
        st.info(f"📁 {len(cad_files)} CAD file(s) uploaded")
        
        if st.button("🚀 Process All CAD Files"):
            cad_screenshot_paths = []
            
            with st.spinner('Processing CAD files... Please wait...'):
                for i, cad_file in enumerate(cad_files):
                    st.info(f"Processing file {i+1}/{len(cad_files)}: {cad_file.name}")
                    
                    cad_temp_dir = tempfile.mkdtemp()
                    cad_path = os.path.join(cad_temp_dir, cad_file.name)
                    
                    with open(cad_path, "wb") as f:
                        f.write(cad_file.read())
                    
                    cad_screenshot_path = open_autocad_and_capture_screenshot(cad_path)
                    
                    if cad_screenshot_path:
                        # Rename to include file index
                        new_path = os.path.join(tempfile.gettempdir(), f"cad_screenshot_{i+1}.png")
                        shutil.copy2(cad_screenshot_path, new_path)
                        cad_screenshot_paths.append(new_path)
                        st.success(f"✅ Captured drawing {i+1}")
                    else:
                        st.error(f"❌ Failed to capture drawing {i+1}")
            
            if cad_screenshot_paths:
                st.session_state['cad_screenshot_paths'] = cad_screenshot_paths
                st.session_state['cad_screenshots_captured'] = True
                st.success(f"🎉 Successfully captured {len(cad_screenshot_paths)} CAD drawings!")
                
                # Display all captured screenshots
                for i, path in enumerate(cad_screenshot_paths):
                    st.image(path, caption=f"CAD Drawing {i+1}")

def updated_tab2_section():
    """
    NX section for standard views capture only
    """
    st.header("📐 NX 3D Model Capture")
    
    # Check if NX models have been captured
    if st.session_state.get('nx_screenshots_captured') and st.session_state.get('nx_model_groups'):
        st.success(f"✅ Currently have {len(st.session_state['nx_model_groups'])} NX model(s) captured")
        
        with st.expander("View Captured NX Models"):
            for group_idx, model_group in enumerate(st.session_state['nx_model_groups']):
                st.subheader(f"NX Model {group_idx + 1}")
                for i, path in enumerate(model_group):
                    if os.path.exists(path):
                        view_names = ['Top View', 'Front View', 'Right View', 'Isometric View']
                        view_name = view_names[i] if i < len(view_names) else f"View {i+1}"
                        st.image(path, caption=view_name)
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🚀 Open NX and Capture Views", key="nx_capture_manual_btn"):
            # Clear previous captures
            if 'nx_model_groups' in st.session_state:
                for model_group in st.session_state['nx_model_groups']:
                    for path in model_group:
                        try:
                            if os.path.exists(path):
                                os.remove(path)
                        except:
                            pass
            
            st.session_state['nx_model_groups'] = []
            st.session_state['nx_screenshots_captured'] = False
            
            views = open_nx_and_capture_views_manual("", manual_file_open=True)

            if views:
                st.success("✅ All 3D views captured successfully from NX!")
                
                model_group = []
                for view_name, view_path in views.items():
                    if os.path.exists(view_path):
                        model_group.append(view_path)
                
                st.session_state['nx_model_groups'] = [model_group]
                st.session_state['nx_screenshots_captured'] = True
                
                with st.expander("View All Captured Screenshots"):
                    for name, path in views.items():
                        if os.path.exists(path):
                            st.image(path, caption=f"{name.capitalize()} View")
            else:
                st.error("Failed to capture 3D views from NX.")
    
    with col2:
        if st.button("🔄 Add Another NX Model", key="nx_add_another_btn"):
            views = open_nx_and_capture_views_manual("", manual_file_open=True)

            if views:
                st.success("✅ Additional 3D views captured successfully!")
                
                model_group = []
                for view_name, view_path in views.items():
                    if os.path.exists(view_path):
                        model_group.append(view_path)
                
                existing_groups = st.session_state.get('nx_model_groups', [])
                existing_groups.append(model_group)
                st.session_state['nx_model_groups'] = existing_groups
                st.session_state['nx_screenshots_captured'] = True
                
                with st.expander("View Newly Captured Screenshots"):
                    for name, path in views.items():
                        if os.path.exists(path):
                            st.image(path, caption=f"{name.capitalize()} View - Model {len(existing_groups)}")
            else:
                st.error("Failed to capture additional 3D views from NX.")
    
    if st.session_state.get('nx_screenshots_captured'):
        st.markdown("---")
        if st.button("🗑️ Clear All NX Captures", key="clear_nx_btn"):
            for model_group in st.session_state.get('nx_model_groups', []):
                for path in model_group:
                    try:
                        if os.path.exists(path):
                            os.remove(path)
                    except:
                        pass
            
            st.session_state['nx_model_groups'] = []
            st.session_state['nx_screenshots_captured'] = False
            st.success("✅ All NX captures cleared!")
            st.rerun()

def updated_tab3_section():
    """
    Updated report generation section with PowerPoint output
    """
    st.header("📄 Generate Enhanced Reports")
    
    pdf_files = st.file_uploader("Upload PDF Files (Specs)", 
                               type=["pdf"], 
                               accept_multiple_files=True,
                               key="pdf_uploader_multiple")
    excel_files = st.file_uploader("Upload Excel Files (Specifications)", 
                                 type=["xlsx"], 
                                 accept_multiple_files=True,
                                 key="excel_uploader_multiple")
    
    # Status indicators with count
    col1, col2 = st.columns(2)
    with col1:
        cad_count = len(st.session_state.get('cad_screenshot_paths', []))
        cad_status = f"✅ {cad_count} drawings ready" if st.session_state.get('cad_screenshots_captured') else "❌ Not captured"
        st.info(f"CAD Drawings: {cad_status}")
    
    with col2:
        nx_groups = st.session_state.get('nx_model_groups', [])
        nx_count = len(nx_groups)
        nx_status = f"✅ {nx_count} models ready" if st.session_state.get('nx_screenshots_captured') else "❌ Not captured"
        st.info(f"NX Models: {nx_status}")
        
    # Generate Report Button
    if st.button("🚀 Generate Enhanced PowerPoint Report"):
        missing_items = []
        if not pdf_files:
            missing_items.append("PDF files")
        if not excel_files:
            missing_items.append("Excel files")
        if not st.session_state.get('cad_screenshots_captured'):
            missing_items.append("CAD drawing captures")
        if not st.session_state.get('nx_screenshots_captured'):
            missing_items.append("NX model captures")
        
        if missing_items:
            st.error(f"Missing required items: {', '.join(missing_items)}")
            st.info("Please complete all sections before generating the report.")
        else:
            # Generate PowerPoint report
            with st.spinner("Generating enhanced PowerPoint report... Please wait..."):
                with tempfile.TemporaryDirectory() as tmpdir:
                    # Process multiple PDF files
                    all_pdf_info = []
                    for pdf_file in pdf_files:
                        pdf_path = os.path.join(tmpdir, pdf_file.name)
                        with open(pdf_path, "wb") as f:
                            f.write(pdf_file.read())
                        pdf_info = extract_pdf_info(pdf_path)
                        all_pdf_info.extend(pdf_info)
                    
                    # Process multiple Excel files (combine all data)
                    combined_excel_df = pd.DataFrame()
                    for excel_file in excel_files:
                        excel_path = os.path.join(tmpdir, excel_file.name)
                        with open(excel_path, "wb") as f:
                            f.write(excel_file.read())
                        excel_df = read_excel_data(excel_path)
                        combined_excel_df = pd.concat([combined_excel_df, excel_df], ignore_index=True)
                    
                    # Get image paths and model groups
                    cad_screenshot_paths = st.session_state.get('cad_screenshot_paths', [])
                    nx_model_groups = st.session_state.get('nx_model_groups', [])
                    
                    # Output path for PowerPoint
                    ppt_out_path = os.path.join(tmpdir, "Enhanced_Tyre_Report.pptx")
                    
                    # Generate PowerPoint report
                    generate_ppt(all_pdf_info, combined_excel_df, cad_screenshot_paths, nx_model_groups, ppt_out_path)
                    
                    # Download button for PowerPoint
                    with open(ppt_out_path, "rb") as f:
                        st.download_button(
                            "📊 Download Enhanced PowerPoint Report", 
                            f, 
                            file_name="Enhanced_Tyre_Report.pptx", 
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
            
            st.success("Enhanced PowerPoint report generated successfully!")

# ---------- RUN THE APPLICATION ----------
if __name__ == "__main__":
    main()