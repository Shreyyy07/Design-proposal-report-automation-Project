import streamlit as st
import pandas as pd
import fitz  # PyMuPDF for enhanced PDF processing
import io
from PIL import Image as PILImage
import tempfile
import os
import subprocess
import pyautogui
import pygetwindow as gw
import time
import shutil
import base64
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import textwrap
import glob
import winreg

# ---------- ASSET PATHS ----------
# Define paths for assets to make them easy to change and manage.
ASSETS_DIR = "assets"
LOGO_PATH = os.path.join(ASSETS_DIR, "apollo_logo.png")
SLIDE1_PATH = os.path.join(ASSETS_DIR, "slide1.png")
PD_PATH = os.path.join(ASSETS_DIR, "pd.png")
PD2_PATH = os.path.join(ASSETS_DIR, "pd2.png")
MOM_PATH = os.path.join(ASSETS_DIR, "mom.png")
LASTSLIDE_PATH = os.path.join(ASSETS_DIR, "lastslide.png")
HEADER_BANNER_PATH = os.path.join(ASSETS_DIR, "header_banner.png")

def get_apollo_logo_base64():
    """
    Reads the Apollo Tyres logo image and converts it to a base64 string.
    """
    try:
        with open(LOGO_PATH, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    except FileNotFoundError:
        st.warning(f"Logo not found at {LOGO_PATH}. Please ensure it exists in the 'assets' folder.")
        return None

def add_logo_to_streamlit():
    """
    Displays the Apollo Tyres logo at the top-center of the Streamlit application.
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
        st.markdown(
            """
            <div style="text-align: center; margin-bottom: 20px;">
                <h2 style="color: #8A2BE2; margin: 0;">🏎️ APOLLO TYRES LTD</h2>
            </div>
            """,
            unsafe_allow_html=True
        )

# ---------- SOFTWARE DETECTION FUNCTIONS ----------

def find_autocad_executable():
    """
    Automatically finds the AutoCAD executable ('acad.exe') on the system.
    """
    if 'custom_autocad_path' in st.session_state and st.session_state['custom_autocad_path']:
        custom_path = st.session_state['custom_autocad_path']
        if os.path.exists(custom_path):
            return custom_path
    
    possible_paths = [
        "C:/Program Files/Autodesk/AutoCAD 2025/acad.exe",
        "C:/Program Files/Autodesk/AutoCAD 2024/acad.exe",
        "C:/Program Files/Autodesk/AutoCAD 2023/acad.exe",
    ]
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    autocad_patterns = ["C:/Program Files/Autodesk/AutoCAD*/acad.exe"]
    for pattern in autocad_patterns:
        matches = glob.glob(pattern)
        if matches:
            matches.sort(reverse=True)
            return matches[0]
    
    return None

def find_nx_executable():
    """
    Automatically finds the Siemens NX executable ('ugraf.exe') on the system.
    """
    if 'custom_nx_path' in st.session_state and st.session_state['custom_nx_path']:
        custom_path = st.session_state['custom_nx_path']
        if os.path.exists(custom_path):
            return custom_path
    
    nx_patterns = [
        "C:/Program Files/Siemens/NX*/NXBIN/ugraf.exe",
        "D:/abcde/NXBIN/ugraf.exe",
    ]
    
    for pattern in nx_patterns:
        matches = glob.glob(pattern)
        if matches:
            matches.sort(reverse=True)
            return matches[0]
            
    return None

# ---------- SESSION STATE INITIALIZATION ----------

def initialize_session_state():
    """
    Initializes the Streamlit session state variables.
    """
    if 'cad_screenshots_captured' not in st.session_state:
        st.session_state['cad_screenshots_captured'] = False
    if 'cad_screenshot_paths' not in st.session_state:
        st.session_state['cad_screenshot_paths'] = []
    if 'nx_screenshots_captured' not in st.session_state:
        st.session_state['nx_screenshots_captured'] = False
    if 'nx_model_groups' not in st.session_state:
        st.session_state['nx_model_groups'] = []
    if 'custom_autocad_path' not in st.session_state:
        st.session_state['custom_autocad_path'] = ""
    if 'custom_nx_path' not in st.session_state:
        st.session_state['custom_nx_path'] = ""

# ---------- PDF EXTRACTION ----------
def extract_pdf_elements(pdf_path):
    """
    Extracts both text and images from a PDF, page by page, with their bounding boxes.
    """
    doc = fitz.open(pdf_path)
    pages_elements = []
    
    for page_num, page in enumerate(doc):
        page_elements = []
        blocks = page.get_text("dict", flags=11)["blocks"]
        for b in blocks:
            for l in b["lines"]:
                for s in l["spans"]:
                    page_elements.append({
                        "type": "text",
                        "bbox": s["bbox"],
                        "text": s["text"],
                        "size": s["size"] # Also extract font size
                    })

        image_list = page.get_images(full=True)
        for img_index, img in enumerate(image_list):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            
            img_rects = page.get_image_rects(img)
            if img_rects:
                bbox = img_rects[0]
                page_elements.append({
                    "type": "image",
                    "bbox": bbox,
                    "bytes": image_bytes
                })

        page_elements.sort(key=lambda el: (el["bbox"][1], el["bbox"][0]))
        pages_elements.append((page.rect, page_elements))
        
    return pages_elements

# ---------- EXCEL READING ----------
def read_excel_data(excel_path):
    """
    Reads data from an Excel file into a pandas DataFrame.
    """
    try:
        return pd.read_excel(excel_path)
    except Exception as e:
        st.error(f"Excel error: {e}")
        return pd.DataFrame()

# ---------- AUTOCAD FUNCTIONS (SMART CONTENT DETECTION) ----------
def open_autocad_and_capture_screenshot(dwg_path):
    """
    Automates opening AutoCAD and takes a precise screenshot with smart content detection
    to crop only the drawing area, removing all white space from sides.
    """
    autocad_process = None
    try:
        autocad_path = find_autocad_executable()
        if not autocad_path:
            st.error("❌ AutoCAD not found. Please configure the path in settings.")
            return None
        
        st.info(f"📍 Found AutoCAD at: {autocad_path}")
        autocad_process = subprocess.Popen([autocad_path, dwg_path])

        st.info("Opening AutoCAD... Please wait.")
        time.sleep(18)
        
        autocad_window = None
        for _ in range(10): 
            windows = gw.getWindowsWithTitle('AutoCAD')
            if windows:
                autocad_window = windows[0]
                break
            time.sleep(1)
        
        if not autocad_window:
            st.error("Could not find AutoCAD window. Automation failed.")
            if autocad_process: autocad_process.kill()
            return None

        autocad_window.activate()
        if not autocad_window.isMaximized:
            autocad_window.maximize()
        time.sleep(3)

        pyautogui.press('esc', presses=2, interval=0.3)
        pyautogui.typewrite('zoom\n')
        time.sleep(0.5)
        pyautogui.typewrite('e\n')
        time.sleep(5)  # Wait for zoom to complete

        # Additional wait time before taking screenshot
        st.info("Waiting for drawing to render properly...")
        time.sleep(3)

        win_left, win_top, win_width, win_height = autocad_window.left, autocad_window.top, autocad_window.width, autocad_window.height
        
        # Initial aggressive cropping to remove UI elements
        top_offset = 220
        bottom_offset = 140
        left_offset = 350
        right_offset = 120
        
        # Take initial screenshot of the viewport area
        initial_region = (
            win_left + left_offset,
            win_top + top_offset,
            win_width - left_offset - right_offset,
            win_height - top_offset - bottom_offset
        )
        
        initial_screenshot = pyautogui.screenshot(region=initial_region)
        
        # Convert to PIL Image for content detection
        img = initial_screenshot
        img_array = img.load()
        img_width, img_height = img.size
        
        # Find content boundaries by detecting non-white pixels
        # Threshold for "white" pixels (allowing slight variations)
        white_threshold = 240
        
        # Find leftmost non-white column
        left_content = 0
        for x in range(img_width):
            has_content = False
            for y in range(img_height):
                r, g, b = img_array[x, y][:3]
                if r < white_threshold or g < white_threshold or b < white_threshold:
                    has_content = True
                    break
            if has_content:
                left_content = max(0, x - 20)  # Add small margin
                break
        
        # Find rightmost non-white column
        right_content = img_width
        for x in range(img_width - 1, -1, -1):
            has_content = False
            for y in range(img_height):
                r, g, b = img_array[x, y][:3]
                if r < white_threshold or g < white_threshold or b < white_threshold:
                    has_content = True
                    break
            if has_content:
                right_content = min(img_width, x + 20)  # Add small margin
                break
        
        # Find topmost non-white row
        top_content = 0
        for y in range(img_height):
            has_content = False
            for x in range(img_width):
                r, g, b = img_array[x, y][:3]
                if r < white_threshold or g < white_threshold or b < white_threshold:
                    has_content = True
                    break
            if has_content:
                top_content = max(0, y - 10)  # Add small margin
                break
        
        # Find bottommost non-white row
        bottom_content = img_height
        for y in range(img_height - 1, -1, -1):
            has_content = False
            for x in range(img_width):
                r, g, b = img_array[x, y][:3]
                if r < white_threshold or g < white_threshold or b < white_threshold:
                    has_content = True
                    break
            if has_content:
                bottom_content = min(img_height, y + 10)  # Add small margin
                break
        
        # Crop to content boundaries
        if right_content > left_content and bottom_content > top_content:
            cropped_screenshot = img.crop((left_content, top_content, right_content, bottom_content))
        else:
            # Fallback: use original image if content detection fails
            cropped_screenshot = img
        
        output_image_path = os.path.join(tempfile.gettempdir(), "cad_screenshot.png")
        cropped_screenshot.save(output_image_path)

        return output_image_path

    except Exception as e:
        st.error(f"Error with AutoCAD: {e}")
        return None
    finally:
        if autocad_process:
            try:
                st.info("Terminating AutoCAD process...")
                autocad_process.terminate()
                autocad_process.wait(timeout=5)
            except Exception:
                autocad_process.kill()
    
# ---------- NX FUNCTIONS (INCREASED OPENING TIME) ----------
def open_nx_and_capture_views_manual():
    """
    Guides the user through capturing the three required 3D views from Siemens NX.
    Updated with more time for NX to open and load properly.
    """
    try:
        nx_path = find_nx_executable()
        if not nx_path:
            st.error("❌ Siemens NX not found. Please configure the path in settings.")
            return {}
        st.info(f"📍 Found NX at: {nx_path}")
        subprocess.Popen([nx_path], shell=True)
        
        # Increased opening time from 20 to 30 seconds
        st.info("Opening NX... Please wait (this may take up to 30 seconds).")
        time.sleep(30)

        nx_window = gw.getWindowsWithTitle('NX')[0]
        nx_window.activate()
        if not nx_window.isMaximized:
            nx_window.maximize()
        time.sleep(2)

        st.warning("Please manually open your 3D file in NX now.")
        st.info("After opening, follow the instructions for each view.")
        time.sleep(10)
        
        views_to_capture = ['Top View', 'Front View', 'Isometric View']
        screenshots = {}
        status_placeholder = st.empty()
        progress_placeholder = st.empty()
        timer_placeholder = st.empty()

        for i, view in enumerate(views_to_capture):
            status_placeholder.info(f"📸 **Step {i+1}/{len(views_to_capture)}**: Please set the **{view}** in NX")
            progress_placeholder.progress((i) / len(views_to_capture))
            
            # Reduced timer to 5 seconds as requested
            for t in range(5, 0, -1):
                timer_placeholder.info(f"Adjusting to {view}. Screenshot in {t} seconds...")
                time.sleep(1)
            
            timer_placeholder.empty()

            nx_window.activate()
            pyautogui.press('f')
            time.sleep(2)

            screenshot = pyautogui.screenshot()
            screen_width, screen_height = pyautogui.size()
            
            # Updated cropping - remove filename from top (more aggressive top crop)
            left = int(screen_width * 0.28)    
            top = int(screen_height * 0.20)    # Increased from 0.15 to remove filename
            right = int(screen_width * 0.75)   
            bottom = int(screen_height * 0.90)

            cropped_screenshot = screenshot.crop((left, top, right, bottom))
            img_path = os.path.join(tempfile.gettempdir(), f"nxview_{view.replace(' ', '_').lower()}.png")
            cropped_screenshot.save(img_path)
            screenshots[view] = img_path

            st.success(f"{view} captured!")

        status_placeholder.success("🎉 **ALL VIEWS CAPTURED!** 🎉")
        progress_placeholder.progress(1.0)
        st.info("Closing NX...")
        nx_window.close()
        time.sleep(2)
        try:
            pyautogui.press('n')
        except:
            pass
        
        return screenshots

    except Exception as e:
        st.error(f"Error in NX automation: {e}")
        return {}

# ---------- POWERPOINT GENERATION FUNCTIONS (UPDATED) ----------

def add_slide_banner(slide, title_text):
    """
    Adds a custom purple banner with a title to the top of a slide.
    """
    try:
        if os.path.exists(HEADER_BANNER_PATH):
             slide.shapes.add_picture(HEADER_BANNER_PATH, 0, 0, width=prs.slide_width, height=Inches(0.8))
        else:
            banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.8))
            banner.fill.solid()
            banner.fill.fore_color.rgb = RGBColor(138, 43, 226)
            banner.line.fill.background()
    except:
        pass

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.15), Inches(12.33), Inches(0.5))
    title_frame = title_box.text_frame
    p = title_frame.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.LEFT
    title_frame.word_wrap = True

def create_outline_slide(prs, topics):
    """
    Creates the 'Outline' slide with a DYNAMIC bulleted list of the report sections.
    UPDATED: Now adds Apollo logo at bottom left corner.
    """
    blank_layout = prs.slide_layouts[6] 
    slide = prs.slides.add_slide(blank_layout)

    title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(12), Inches(1.0))
    title_tf = title_shape.text_frame
    p_title = title_tf.paragraphs[0]
    p_title.text = "Outline:-"
    p_title.font.size = Pt(44)
    p_title.font.bold = True

    body_shape = slide.shapes.add_textbox(Inches(1.0), Inches(1.5), Inches(11), Inches(5.0))
    tf = body_shape.text_frame
    tf.clear()

    for topic in topics:
        p = tf.add_paragraph()
        p.text = topic
        p.font.size = Pt(24)
        p.level = 0
        p.space_after = Pt(12)

    # Add Apollo logo at bottom left corner
    try:
        if os.path.exists(LOGO_PATH):
            logo_width = Inches(1.5)
            logo_height = Inches(0.75)
            logo_left = Inches(0.3)
            logo_top = prs.slide_height - logo_height - Inches(0.3)
            slide.shapes.add_picture(LOGO_PATH, logo_left, logo_top, logo_width, logo_height)
    except Exception as e:
        pass  # Logo addition failed, continue without it


def generate_ppt(pdf_pages, excel_df, cad_image_paths, nx_model_groups, output_path):
    """
    Main function to generate the entire PowerPoint presentation with the corrected slide order
    and dynamic outline.
    """
    global prs
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- 1. Front Page Slide ---
    slide1 = prs.slides.add_slide(blank_layout)
    try:
        slide1.shapes.add_picture(SLIDE1_PATH, 0, 0, prs.slide_width, prs.slide_height)
    except Exception:
        st.warning(f"Could not find {SLIDE1_PATH}. Slide 1 will be blank.")

    # --- 2. Dynamic Outline Slide ---
    outline_topics = []
    if pdf_pages:
        outline_topics.append("Project brief details")
    outline_topics.append("Minutes of the Meeting")
    if not excel_df.empty:
        outline_topics.append("Design input sheet")
    if cad_image_paths:
        outline_topics.append("Cavity and thread profile details")
    if nx_model_groups:
        outline_topics.append("3D model images")
        outline_topics.append("Front & isometric close-up view of tyre images")
    create_outline_slide(prs, outline_topics)

    # --- 3. pd.png Slide (RESTORED) ---
    slide2 = prs.slides.add_slide(blank_layout)
    try:
        slide2.shapes.add_picture(PD_PATH, 0, 0, prs.slide_width, prs.slide_height)
    except Exception:
        st.warning(f"Could not find {PD_PATH}. Slide 3 will be blank.")

    # --- 4. pd2.png Slide ---
    slide3 = prs.slides.add_slide(blank_layout)
    try:
        slide3.shapes.add_picture(PD2_PATH, 0, 0, prs.slide_width, prs.slide_height)
    except Exception:
        st.warning(f"Could not find {PD2_PATH}. Slide 4 will be blank.")
        
    # --- 5. Minutes of the Meeting Slide ---
    mom_slide = prs.slides.add_slide(blank_layout)
    add_slide_banner(mom_slide, "Minutes of the Meeting")
    try:
        img_width = Inches(12) 
        img_height = Inches(6)
        left = (prs.slide_width - img_width) / 2
        top = (prs.slide_height - img_height) / 2 + Inches(0.2)
        mom_slide.shapes.add_picture(MOM_PATH, left, top, width=img_width, height=img_height)
    except Exception:
        st.warning(f"Could not find {MOM_PATH}. MOM slide will be blank.")

    # --- Content Sections ---
    if pdf_pages:
        create_pdf_content_slides(prs, pdf_pages, section_title="Project brief details")
    if not excel_df.empty:
        create_excel_slides(prs, excel_df, section_title="Design input sheet:")
    if cad_image_paths:
        create_cad_slides(prs, cad_image_paths, section_title="Cavity and thread profile details")
    if nx_model_groups:
        create_nx_model_slides(prs, nx_model_groups, section_title="3D model images")
        # Add close-up view slide right after 3D model slides
        create_nx_closeup_slide(prs, nx_model_groups, section_title="Front & isometric close-up view of tyre images")

    # --- Final Thank You Slide ---
    thank_slide = prs.slides.add_slide(blank_layout)
    try:
        thank_slide.shapes.add_picture(LASTSLIDE_PATH, 0, 0, prs.slide_width, prs.slide_height)
    except Exception:
        st.warning(f"Could not find {LASTSLIDE_PATH}. Final slide will be blank.")

    prs.save(output_path)

def create_pdf_content_slides(prs, pdf_pages, section_title):
    """
    FIXED: Creates multiple slides for PDF content with proper text distribution using WORD COUNT method.
    """
    blank_layout = prs.slide_layouts[6]
    
    for page_idx, (page_rect, page_elements) in enumerate(pdf_pages):
        # Extract all text content first
        all_text_content = []
        for element in page_elements:
            if element["type"] == "text" and element["text"].strip():
                all_text_content.append(element["text"].strip())
        
        # Join all text
        full_text = " ".join(all_text_content)
        
        # Split text by WORD COUNT - much more reliable
        words = full_text.split()
        words_per_slide = 80  # Conservative word count per slide
        
        # Create word chunks
        word_chunks = []
        for i in range(0, len(words), words_per_slide):
            chunk = words[i:i + words_per_slide]
            word_chunks.append(" ".join(chunk))
        
        # Ensure we have at least one chunk
        if not word_chunks:
            word_chunks = [full_text] if full_text.strip() else ["No content available"]
        
        # Create slides for each word chunk
        for chunk_idx, chunk_text in enumerate(word_chunks):
            slide = prs.slides.add_slide(blank_layout)
            
            # Add slide title
            if len(word_chunks) > 1:
                title = f"{section_title} - Page {page_idx + 1} (Part {chunk_idx + 1}/{len(word_chunks)})"
            else:
                title = f"{section_title} - Page {page_idx + 1}"
            
            add_slide_banner(slide, title)
            
            # Add text content with FIXED height
            content_top = Inches(1.0)
            content_height = Inches(5.5)  # Fixed height to prevent overflow
            
            text_box = slide.shapes.add_textbox(
                Inches(0.5),
                content_top,
                Inches(12.33),
                content_height
            )
            
            text_frame = text_box.text_frame
            text_frame.clear()
            text_frame.word_wrap = True
            text_frame.auto_size = MSO_AUTO_SIZE.NONE  # Prevent auto-sizing
            
            # Add text as one paragraph to prevent overflow
            p = text_frame.paragraphs[0]
            p.text = chunk_text
            p.font.size = Pt(14)
            p.space_after = Pt(6)
            p.line_spacing = 1.2
            p.alignment = PP_ALIGN.LEFT
        
        # Handle images on separate slides
        image_count = 0
        for element in page_elements:
            if element["type"] == "image":
                image_count += 1
                slide = prs.slides.add_slide(blank_layout)
                add_slide_banner(slide, f"{section_title} - Page {page_idx + 1} - Image {image_count}")
                
                try:
                    img_width = Inches(10)
                    img_height = Inches(5)
                    img_left = (prs.slide_width - img_width) / 2
                    img_top = Inches(1.5)
                    
                    image_stream = io.BytesIO(element["bytes"])
                    slide.shapes.add_picture(image_stream, img_left, img_top, img_width, img_height)
                except Exception as e:
                    print(f"Could not add PDF image to slide: {e}")

def create_excel_slides(prs, excel_df, section_title):
    """
    Creates slides for Excel data with pagination.
    """
    if excel_df.empty:
        return
    
    blank_layout = prs.slide_layouts[6]
    max_rows_per_slide = 20
    total_rows = len(excel_df)
    num_slides = (total_rows + max_rows_per_slide - 1) // max_rows_per_slide

    for slide_idx in range(num_slides):
        slide = prs.slides.add_slide(blank_layout)
        title_text = section_title
        if num_slides > 1:
            title_text += f" (Page {slide_idx + 1}/{num_slides})"
        add_slide_banner(slide, title_text)

        start_row = slide_idx * max_rows_per_slide
        end_row = min(start_row + max_rows_per_slide, total_rows)
        df_slice = excel_df.iloc[start_row:end_row]

        rows, cols = df_slice.shape
        rows += 1

        table_left = Inches(0.5)
        table_top = Inches(1.2)
        table_width = prs.slide_width - Inches(1.0)
        table_height = prs.slide_height - Inches(1.5)

        table_shape = slide.shapes.add_table(rows, cols, table_left, table_top, table_width, table_height)
        table = table_shape.table

        for c in range(cols):
            table.columns[c].width = int(table_width / cols)

        for c, col_name in enumerate(df_slice.columns):
            cell = table.cell(0, c)
            p = cell.text_frame.paragraphs[0]
            p.text = str(col_name)
            p.font.bold = True
            p.font.size = Pt(10)
            p.alignment = PP_ALIGN.CENTER
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(138, 43, 226)
            p.font.color.rgb = RGBColor(255, 255, 255)

        for r, row_data in enumerate(df_slice.itertuples(index=False)):
            for c, value in enumerate(row_data):
                cell = table.cell(r + 1, c)
                p = cell.text_frame.paragraphs[0]
                p.text = str(value) if pd.notna(value) else ""
                p.font.size = Pt(9)
                p.alignment = PP_ALIGN.CENTER


def create_cad_slides(prs, cad_image_paths, section_title):
    """
    Creates a slide for each CAD screenshot, ensuring it uses the full content width.
    """
    slide_layout = prs.slide_layouts[6]
    for i, cad_path in enumerate(cad_image_paths):
        if cad_path and os.path.exists(cad_path):
            slide = prs.slides.add_slide(slide_layout)
            add_slide_banner(slide, f"{section_title} - Drawing {i+1}")
            
            content_area_top = Inches(0.8)
            content_area_height = prs.slide_height - content_area_top
            
            try:
                display_w = prs.slide_width
                
                img = PILImage.open(cad_path)
                aspect = img.height / img.width
                display_h = display_w * aspect

                left = 0
                top = content_area_top + (content_area_height - display_h) / 2
                
                if top < content_area_top:
                    top = content_area_top
                    display_h = content_area_height
                    display_w = display_h / aspect
                    left = (prs.slide_width - display_w) / 2

                slide.shapes.add_picture(cad_path, left, top, width=display_w)
            except Exception as e:
                st.error(f"Error adding CAD image {i+1}: {e}")

def create_nx_model_slides(prs, nx_model_groups, section_title):
    """
    Creates a slide for each NX model, arranging the three views horizontally
    with section name on top and proper view labels: Front View, Back View, Isometric View.
    """
    slide_layout = prs.slide_layouts[6]
    for group_idx, model_group in enumerate(nx_model_groups):
        if len(model_group) < 3: continue

        slide = prs.slides.add_slide(slide_layout)
        add_slide_banner(slide, f"{section_title} - Model {group_idx + 1}")
        
        total_content_width = prs.slide_width - Inches(1.0)
        gap = Inches(0.15)
        
        image_width = (total_content_width - (2 * gap)) / 3
        
        view_paths = model_group[:3]
        # Updated view names as requested: Front, Back, Isometric
        view_names = ['Front View', 'Back View', 'Isometric View']
        
        current_left = Inches(0.5)
        
        for i, img_path in enumerate(view_paths):
            try:
                img = PILImage.open(img_path)
                aspect_ratio = img.height / img.width
                image_height = image_width * aspect_ratio

                top = Inches(1.0) + ((prs.slide_height - Inches(1.0) - image_height) / 2)

                slide.shapes.add_picture(img_path, current_left, top, width=image_width, height=image_height)
                
                # Add view name label below each image
                label_top = top + image_height + Inches(0.1)
                label_box = slide.shapes.add_textbox(current_left, label_top, image_width, Inches(0.3))
                label_frame = label_box.text_frame
                p = label_frame.paragraphs[0]
                p.text = view_names[i] if i < len(view_names) else f"View {i+1}"
                p.font.size = Pt(14)
                p.font.bold = True
                p.alignment = PP_ALIGN.CENTER
                
                current_left += image_width + gap

            except Exception as e:
                st.error(f"Could not add NX image {i+1} to slide: {e}")

def create_nx_closeup_slide(prs, nx_model_groups, section_title):
    """
    Creates a close-up slide showing Front and Isometric views side by side
    with IMPROVED FRONT VIEW cropping to show full tire width.
    """
    if not nx_model_groups or len(nx_model_groups) == 0:
        return
    
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    
    # Add header banner
    add_slide_banner(slide, section_title)
    
    # Get first model group
    model_group = nx_model_groups[0]
    
    if len(model_group) >= 3:
        # Front view (index 1) and Isometric view (index 2)
        front_view_path = model_group[1]  # Front View
        isometric_view_path = model_group[2]  # Isometric View
        
        # Calculate positions for side-by-side layout
        gap = Inches(0.5)
        image_width = (prs.slide_width - Inches(1.0) - gap) / 2
        
        # Left image (Front view)
        left_img_left = Inches(0.5)
        # Right image (Isometric view)  
        right_img_left = left_img_left + image_width + gap
        
        content_top = Inches(1.2)
        available_height = prs.slide_height - content_top - Inches(1.0)  # Leave space for labels
        
        def create_front_view_crop(img_path, zoom_factor=1.8):
            """Create a VERTICAL STRIP crop for front view to show full tire width"""
            try:
                img = PILImage.open(img_path)
                width, height = img.size
                
                # For front view, we want a vertical strip showing the full tire width
                # Take a vertical slice from the center
                slice_width = int(width / zoom_factor)
                center_x = width // 2
                
                # Calculate vertical strip boundaries
                left = max(0, center_x - slice_width // 2)
                right = min(width, center_x + slice_width // 2)
                top = 0  # Keep full height
                bottom = height  # Keep full height
                
                # Crop to vertical strip and resize for clarity
                cropped = img.crop((left, top, right, bottom))
                # Resize to emphasize the tire pattern
                resized = cropped.resize((width, height), PILImage.Resampling.LANCZOS)
                
                # Save cropped image
                cropped_path = img_path.replace('.png', '_front_strip.png')
                resized.save(cropped_path)
                return cropped_path
            except Exception as e:
                print(f"Error creating front view crop: {e}")
                return img_path
        
        def create_center_cropped_zoomed_image(img_path, zoom_factor=2.0):
            """Create a center-cropped zoomed version for isometric view"""
            try:
                img = PILImage.open(img_path)
                width, height = img.size
                
                # Find center of image
                center_x = width // 2
                center_y = height // 2
                
                # Calculate crop dimensions for zoom
                crop_width = int(width / zoom_factor)
                crop_height = int(height / zoom_factor)
                
                # Ensure center crop
                left = max(0, center_x - crop_width // 2)
                top = max(0, center_y - crop_height // 2)
                right = min(width, left + crop_width)
                bottom = min(height, top + crop_height)
                
                # Adjust if crop goes beyond image boundaries
                if right > width:
                    right = width
                    left = right - crop_width
                if bottom > height:
                    bottom = height
                    top = bottom - crop_height
                if left < 0:
                    left = 0
                    right = left + crop_width
                if top < 0:
                    top = 0
                    bottom = top + crop_height
                
                # Crop from center and resize for zoom effect
                cropped = img.crop((left, top, right, bottom))
                zoomed = cropped.resize((width, height), PILImage.Resampling.LANCZOS)
                
                # Save zoomed image
                zoomed_path = img_path.replace('.png', '_center_zoomed.png')
                zoomed.save(zoomed_path)
                return zoomed_path
            except Exception as e:
                print(f"Error creating center zoomed image: {e}")
                return img_path
        
        # Add Front view (vertical strip crop to show full tire width)
        try:
            if os.path.exists(front_view_path):
                cropped_front_path = create_front_view_crop(front_view_path, zoom_factor=1.8)
                
                img = PILImage.open(cropped_front_path)
                aspect_ratio = img.height / img.width
                image_height = min(image_width * aspect_ratio, available_height - Inches(0.5))
                
                top_pos = content_top + (available_height - image_height - Inches(0.5)) / 2
                
                slide.shapes.add_picture(cropped_front_path, left_img_left, top_pos, 
                                       width=image_width, height=image_height)
                
                # Add "Front View" label below the image
                label_top = top_pos + image_height + Inches(0.1)
                label_box = slide.shapes.add_textbox(left_img_left, label_top, image_width, Inches(0.4))
                label_frame = label_box.text_frame
                p = label_frame.paragraphs[0]
                p.text = "Front View"
                p.font.size = Pt(16)
                p.font.bold = True
                p.alignment = PP_ALIGN.CENTER
                
        except Exception as e:
            st.error(f"Error adding improved front view: {e}")
        
        # Add Isometric view (center-cropped and zoomed)
        try:
            if os.path.exists(isometric_view_path):
                zoomed_iso_path = create_center_cropped_zoomed_image(isometric_view_path, zoom_factor=2.5)
                
                img = PILImage.open(zoomed_iso_path)
                aspect_ratio = img.height / img.width
                image_height = min(image_width * aspect_ratio, available_height - Inches(0.5))
                
                top_pos = content_top + (available_height - image_height - Inches(0.5)) / 2
                
                slide.shapes.add_picture(zoomed_iso_path, right_img_left, top_pos, 
                                       width=image_width, height=image_height)
                
                # Add "Isometric View" label below the image
                label_top = top_pos + image_height + Inches(0.1)
                label_box = slide.shapes.add_textbox(right_img_left, label_top, image_width, Inches(0.4))
                label_frame = label_box.text_frame
                p = label_frame.paragraphs[0]
                p.text = "Isometric View"
                p.font.size = Pt(16)
                p.font.bold = True
                p.alignment = PP_ALIGN.CENTER
                
        except Exception as e:
            st.error(f"Error adding zoomed isometric view: {e}")


# ---------- STREAMLIT GUI ----------
def main():
    """
    The main function that runs the Streamlit application.
    """
    add_logo_to_streamlit()
    st.title("Enhanced Tyre Report Generator")
    initialize_session_state()

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

def updated_tab1_section():
    """
    UI and logic for the 'Capture CAD Drawing' tab.
    """
    st.header("📸 AutoCAD Drawing Capture")
    cad_files = st.file_uploader("Upload 2D CAD Files (.dwg)", type=["dwg"], accept_multiple_files=True)
    
    if cad_files:
        if st.button("🚀 Process All CAD Files"):
            paths = []
            with st.spinner('Processing CAD files... This is now faster.'):
                for i, cad_file in enumerate(cad_files):
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".dwg") as tmp:
                        tmp.write(cad_file.getvalue())
                        cad_path = tmp.name
                    
                    ss_path = open_autocad_and_capture_screenshot(cad_path)
                    
                    try:
                        os.remove(cad_path)
                    except OSError as e:
                        st.warning(f"Could not remove temp file {cad_path}: {e}")

                    if ss_path:
                        final_path = os.path.join(tempfile.gettempdir(), f"cad_ss_{i}.png")
                        shutil.copy(ss_path, final_path)
                        paths.append(final_path)
                        st.success(f"✅ Captured {cad_file.name}")
                    else:
                        st.error(f"❌ Failed to capture {cad_file.name}")

            if paths:
                st.session_state['cad_screenshot_paths'] = paths
                st.session_state['cad_screenshots_captured'] = True
                st.success("🎉 All CAD drawings captured!")
                for path in paths:
                    st.image(path)

def updated_tab2_section():
    """
    UI and logic for the 'Capture NX 3D Models' tab.
    """
    st.header("📐 NX 3D Model Capture")
    
    if st.session_state.get('nx_model_groups'):
        st.success(f"✅ {len(st.session_state.get('nx_model_groups', []))} NX model(s) captured.")
        with st.expander("View Captured Models"):
            for i, model_group in enumerate(st.session_state.get('nx_model_groups', [])):
                st.subheader(f"Model {i+1}")
                cols = st.columns(len(model_group))
                for col, path in zip(cols, model_group):
                    col.image(path)

    if st.session_state.get('nx_screenshots_captured'):
        if st.button("🗑️ Clear All NX Captures"):
            st.session_state['nx_model_groups'] = []
            st.session_state['nx_screenshots_captured'] = False
            st.rerun()

    if st.button("🚀 Capture New NX Model"):
        with st.spinner("Waiting for NX automation..."):
            views = open_nx_and_capture_views_manual()
            if views and len(views) == 3:
                model_group = [views[v] for v in ['Top View', 'Front View', 'Isometric View']]
                st.session_state.setdefault('nx_model_groups', []).append(model_group)
                st.session_state['nx_screenshots_captured'] = True
                st.success("✅ Capture complete! Please return to the Streamlit app.")
                st.rerun()
            else:
                st.error("❌ Failed to capture all 3 required views from NX.")

def updated_tab3_section():
    """
    UI and logic for the 'Generate Reports' tab.
    """
    st.header("📄 Generate Enhanced Reports")
    
    with st.expander("🔧 Software Configuration (Optional)"):
        st.info("The app tries to auto-detect software. Use these fields to override.")
        custom_autocad = st.text_input("Custom AutoCAD Path:", st.session_state.get('custom_autocad_path', ''))
        st.session_state['custom_autocad_path'] = custom_autocad
        
        custom_nx = st.text_input("Custom NX Path:", st.session_state.get('custom_nx_path', ''))
        st.session_state['custom_nx_path'] = custom_nx

    pdf_files = st.file_uploader("Upload PDF Files", type=["pdf"], accept_multiple_files=True)
    excel_files = st.file_uploader("Upload Excel Files", type=["xlsx"], accept_multiple_files=True)
    
    st.info(f"CAD Drawings: {'✅ Ready' if st.session_state.get('cad_screenshots_captured') else '❌ Not captured'}")
    st.info(f"NX Models: {'✅ Ready' if st.session_state.get('nx_screenshots_captured') else '❌ Not captured'}")
        
    if st.button("🚀 Generate Enhanced PowerPoint Report"):
        if not all([pdf_files, excel_files, st.session_state.get('cad_screenshots_captured'), st.session_state.get('nx_screenshots_captured')]):
            st.error("Missing one or more inputs. Please provide all files and captures.")
        else:
            with st.spinner("Generating report... This may take a moment."):
                with tempfile.TemporaryDirectory() as tmpdir:
                    all_pdf_pages = []
                    for pdf_file in pdf_files:
                        pdf_path = os.path.join(tmpdir, pdf_file.name)
                        with open(pdf_path, "wb") as f:
                            f.write(pdf_file.getvalue())
                        all_pdf_pages.extend(extract_pdf_elements(pdf_path))
                    
                    all_excel_df = pd.concat(
                        [read_excel_data(excel_file) for excel_file in excel_files], 
                        ignore_index=True
                    )
                    
                    ppt_out_path = os.path.join(tmpdir, "Enhanced_Tyre_Report.pptx")
                    
                    generate_ppt(
                        all_pdf_pages, 
                        all_excel_df, 
                        st.session_state.get('cad_screenshot_paths', []), 
                        st.session_state.get('nx_model_groups', []), 
                        ppt_out_path
                    )
                    
                    with open(ppt_out_path, "rb") as f:
                        st.download_button(
                            "📊 Download Enhanced PowerPoint Report", 
                            f, 
                            file_name="Enhanced_Tyre_Report.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
            st.success("Report generated!")

if __name__ == "__main__":
    main()