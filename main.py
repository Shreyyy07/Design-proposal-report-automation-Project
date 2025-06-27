import streamlit as st
import pandas as pd
import pdfplumber
import cv2
import tempfile
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
# from pptx import Presentation
# from pptx.util import Inches
import os
import subprocess
import pyautogui
import pygetwindow as gw
import time
import shutil
import traceback
import numpy as np
import matplotlib.pyplot as plt
from skimage import measure, morphology, filters
from scipy import ndimage
import base64
from io import BytesIO

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
    # NEW: Add wear analysis session state variables
    if 'wear_analysis_completed' not in st.session_state:
        st.session_state['wear_analysis_completed'] = False
    if 'wear_analysis_results' not in st.session_state:
        st.session_state['wear_analysis_results'] = {}
    if 'original_tyre_image' not in st.session_state:
        st.session_state['original_tyre_image'] = None

# ---------- PDF Extraction (UNCHANGED) ----------
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

# ---------- Excel Reading (UNCHANGED) ----------
def read_excel_data(excel_path):
    try:
        return pd.read_excel(excel_path)
    except Exception as e:
        st.error(f"Excel error: {e}")
        return pd.DataFrame()

# ---------- AutoCAD Functions (UNCHANGED) ----------
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
        
        # More precise cropping for drawing area only
        left = screen_width * 0.05   # Remove left panel/toolbars
        top = screen_height * 0.15   # Remove ribbon and title bar
        right = screen_width * 0.95  # Remove right panels
        bottom = screen_height * 0.85  # Remove command line and status bar

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
    
    # ---------- NEW: Tyre Wear Analysis Functions ----------
def analyze_tyre_wear(image_path, wear_percentage):
    """
    Enhanced tyre wear analysis with realistic wear visualization
    Creates visual representations of different wear stages with detailed effects
    """
    try:
        import cv2
        import numpy as np
        from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageEnhance
        
        # Load the original image
        img = cv2.imread(image_path)
        if img is None:
            return None
            
        # Convert to PIL for easier manipulation
        pil_img = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
        
        # Create enhanced wear simulation
        enhanced_img = pil_img.copy()
        
        # Apply wear effects based on percentage
        if wear_percentage > 0:
            # Convert to numpy for advanced processing
            img_array = np.array(enhanced_img)
            
            # Create wear pattern overlay
            height, width = img_array.shape[:2]
            
            # Generate wear texture based on percentage
            wear_intensity = wear_percentage / 100.0
            
            # Create random wear pattern
            np.random.seed(42)  # For consistent results
            wear_pattern = np.random.random((height, width)) * wear_intensity
            
            # Apply different wear effects based on percentage
            if wear_percentage >= 25:
                # Add surface roughness
                noise = np.random.normal(0, 20 * wear_intensity, img_array.shape)
                img_array = np.clip(img_array + noise, 0, 255)
                
            if wear_percentage >= 50:
                # Add darker patches (worn areas)
                dark_patches = wear_pattern > 0.3
                img_array[dark_patches] = img_array[dark_patches] * (1 - wear_intensity * 0.4)
                
                # Add tread wear lines
                for i in range(0, height, max(1, int(30 * (1 - wear_intensity)))):
                    start_y = max(0, i - 2)
                    end_y = min(height, i + 2)
                    img_array[start_y:end_y, :] = img_array[start_y:end_y, :] * 0.7
                    
            if wear_percentage >= 75:
                # Severe wear - add more dramatic effects
                severe_wear = wear_pattern > 0.2
                img_array[severe_wear] = img_array[severe_wear] * 0.5
                
                # Add worn spots
                for _ in range(int(wear_percentage / 10)):
                    spot_x = np.random.randint(0, width - 50)
                    spot_y = np.random.randint(0, height - 50)
                    img_array[spot_y:spot_y+50, spot_x:spot_x+50] = img_array[spot_y:spot_y+50, spot_x:spot_x+50] * 0.4
            
            # Convert back to PIL
            enhanced_img = Image.fromarray(np.uint8(np.clip(img_array, 0, 255)))
            
            # Apply additional PIL effects
            if wear_percentage > 0:
                # Reduce contrast for worn look
                enhancer = ImageEnhance.Contrast(enhanced_img)
                enhanced_img = enhancer.enhance(1 - wear_intensity * 0.3)
                
                # Add slight blur for worn texture
                if wear_percentage > 50:
                    enhanced_img = enhanced_img.filter(ImageFilter.GaussianBlur(radius=wear_intensity))
        
        # Add wear information overlay
        draw = ImageDraw.Draw(enhanced_img)
        
        # Determine text color and background based on wear percentage
        if wear_percentage == 0:
            text_color = (0, 255, 0)  # Green for new
            bg_color = (0, 0, 0, 180)  # Semi-transparent black
            status_text = "NEW TYRE"
        elif wear_percentage <= 25:
            text_color = (255, 255, 0)  # Yellow for light wear
            bg_color = (0, 0, 0, 180)
            status_text = "LIGHT WEAR"
        elif wear_percentage <= 50:
            text_color = (255, 165, 0)  # Orange for moderate wear
            bg_color = (0, 0, 0, 180)
            status_text = "MODERATE WEAR"
        else:
            text_color = (255, 0, 0)  # Red for heavy wear
            bg_color = (0, 0, 0, 180)
            status_text = "HEAVY WEAR"
        
        # Load font
        try:
            title_font = ImageFont.truetype("arial.ttf", 36)
            percent_font = ImageFont.truetype("arial.ttf", 48)
            status_font = ImageFont.truetype("arial.ttf", 24)
        except:
            title_font = ImageFont.load_default()
            percent_font = ImageFont.load_default()
            status_font = ImageFont.load_default()
        
        # Create overlay for text background
        overlay = Image.new('RGBA', enhanced_img.size, (0, 0, 0, 0))
        overlay_draw = ImageDraw.Draw(overlay)
        
        # Main percentage text
        percent_text = f"{wear_percentage}%"
        percent_bbox = draw.textbbox((0, 0), percent_text, font=percent_font)
        percent_width = percent_bbox[2] - percent_bbox[0]
        percent_height = percent_bbox[3] - percent_bbox[1]
        
        # Position text at top right
        percent_x = enhanced_img.width - percent_width - 30
        percent_y = 20
        
        # Draw background rectangle for percentage
        overlay_draw.rectangle([
            percent_x - 15, percent_y - 10,
            percent_x + percent_width + 15, percent_y + percent_height + 10
        ], fill=bg_color)
        
        # Status text
        status_bbox = draw.textbbox((0, 0), status_text, font=status_font)
        status_width = status_bbox[2] - status_bbox[0]
        status_height = status_bbox[3] - status_bbox[1]
        
        status_x = percent_x + (percent_width - status_width) // 2
        status_y = percent_y + percent_height + 15
        
        # Draw background rectangle for status
        overlay_draw.rectangle([
            status_x - 10, status_y - 5,
            status_x + status_width + 10, status_y + status_height + 5
        ], fill=bg_color)
        
        # Composite the overlay
        enhanced_img = Image.alpha_composite(enhanced_img.convert('RGBA'), overlay)
        
        # Draw the text
        final_draw = ImageDraw.Draw(enhanced_img)
        final_draw.text((percent_x, percent_y), percent_text, fill=text_color, font=percent_font)
        final_draw.text((status_x, status_y), status_text, fill=text_color, font=status_font)
        
        # Add wear indicator bar at bottom
        bar_height = 20
        bar_y = enhanced_img.height - bar_height - 20
        bar_x = 30
        bar_width = enhanced_img.width - 60
        
        # Background bar
        final_draw.rectangle([bar_x, bar_y, bar_x + bar_width, bar_y + bar_height], 
                           fill=(100, 100, 100, 200))
        
        # Wear progress bar
        progress_width = int(bar_width * (wear_percentage / 100))
        if progress_width > 0:
            final_draw.rectangle([bar_x, bar_y, bar_x + progress_width, bar_y + bar_height], 
                               fill=text_color)
        
        # Add "WEAR LEVEL" text above bar
        wear_label = "WEAR LEVEL"
        label_bbox = draw.textbbox((0, 0), wear_label, font=status_font)
        label_width = label_bbox[2] - label_bbox[0]
        label_x = bar_x + (bar_width - label_width) // 2
        label_y = bar_y - 30
        
        final_draw.rectangle([label_x - 5, label_y - 3, label_x + label_width + 5, label_y + 20], 
                           fill=(0, 0, 0, 150))
        final_draw.text((label_x, label_y), wear_label, fill=(255, 255, 255), font=status_font)
        
        # Save the result
        output_path = image_path.replace('.png', f'_wear_{wear_percentage}%.png')
        enhanced_img.convert('RGB').save(output_path, quality=95)
        
        return output_path
        
    except Exception as e:
        st.error(f"Error in tyre wear analysis: {e}")
        return None

def perform_tyre_wear_analysis(original_image_path):
    """
    Generate all four wear analysis images (0%, 25%, 50%, 75%)
    """
    wear_percentages = [0, 25, 50, 75]
    analysis_results = {}
    
    for percentage in wear_percentages:
        analyzed_path = analyze_tyre_wear(original_image_path, percentage)
        if analyzed_path:
            analysis_results[f"{percentage}%"] = analyzed_path
    
    return analysis_results

def open_nx_for_wear_analysis():
    """
    Open NX specifically for tyre wear analysis - captures single isometric view
    """
    try:
        # Launch NX
        nx_path = r"D:\abcde\NXBIN\ugraf.exe"
        nx_process = subprocess.Popen([nx_path], shell=True)
        st.info("Opening NX for tyre wear analysis... Please wait.")
        time.sleep(20)
        
        # Find and activate NX window
        nx_window = None
        for w in gw.getWindowsWithTitle('NX'):
            if 'NX' in w.title:
                if not w.isMaximized:
                    w.maximize()
                w.activate()
                nx_window = w
                break
        
        if not nx_window:
            st.error("Could not find NX window for wear analysis.")
            return None
        
        # Instructions for user
        st.warning("🔧 Please manually open your 3D tyre model in NX for wear analysis.")
        st.info("📋 Set up an isometric view showing the tyre tread clearly.")
        
        # Wait for manual setup
        time.sleep(15)
        
        # Countdown for capture
        countdown_placeholder = st.empty()
        for countdown in range(8, 0, -1):
            countdown_placeholder.warning(f"⏰ Capturing tyre model for wear analysis in {countdown} seconds...")
            time.sleep(1)
        countdown_placeholder.empty()
        
        # Ensure NX is active and fit view
        nx_window.activate()
        time.sleep(1)
        pyautogui.press('f')  # Fit view
        time.sleep(2)
        
        # Capture screenshot
        st.success("📸 Capturing tyre model for wear analysis...")
        screenshot = pyautogui.screenshot()
        
        # Crop to viewport
        screen_width, screen_height = pyautogui.size()
        left = screen_width * 0.12
        top = screen_height * 0.12
        right = screen_width * 0.88
        bottom = screen_height * 0.88
        
        cropped_screenshot = screenshot.crop((left, top, right, bottom))
        
        # Save original capture
        output_path = os.path.join(tempfile.gettempdir(), "tyre_wear_original.png")
        cropped_screenshot.save(output_path)
        
        # Close NX
        st.info("Closing NX application...")
        nx_window.activate()
        pyautogui.hotkey('alt', 'f4')
        time.sleep(2)
        pyautogui.press('n')  # Don't save
        
        try:
            nx_process.terminate()
        except:
            pass
        
        return output_path
        
    except Exception as e:
        st.error(f"Error in NX wear analysis capture: {e}")
        return None
    
# ---------- NX Functions - Updated for Multiple Files ----------
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

# def open_nx_and_capture_screenshot(prt_path):
#     """
#     Open NX with a specific file and capture a screenshot
#     """
#     try:
#         # Launch NX with the specific file
#         nx_path = r"D:\abcde\NXBIN\ugraf.exe"  # Adjust this to your NX installation path
#         nx_process = subprocess.Popen([nx_path, prt_path], shell=True)
        
#         st.info("Opening NX... Waiting for file to load.")
#         time.sleep(20)  # Wait for NX to open and load the file
        
#         # Find and activate the NX window
#         nx_window = None
#         attempts = 0
#         while attempts < 3 and not nx_window:
#             for w in gw.getWindowsWithTitle('NX'):
#                 if 'NX' in w.title:
#                     if not w.isMaximized:
#                         w.maximize()
#                     w.activate()
#                     nx_window = w
#                     break
#             if not nx_window:
#                 time.sleep(3)
#                 attempts += 1
                
#         if not nx_window:
#             st.error("Could not find NX window. Please check if NX launched properly.")
#             return None

#         # Dismiss any startup dialogs
#         pyautogui.press('escape')
#         time.sleep(1)
        
#         # Click on empty area and ensure NX is active
#         screen_width, screen_height = pyautogui.size()
#         pyautogui.click(screen_width//2, screen_height//2)
#         nx_window.activate()
#         time.sleep(2)

#         # Fit view to screen for better capture
#         pyautogui.press('f')
#         time.sleep(3)
        
#         # Take full screen screenshot first
#         screenshot = pyautogui.screenshot()
        
#         # Crop to show only the 3D viewport area
#         left = screen_width * 0.12   # Remove left toolbar area
#         top = screen_height * 0.12   # Remove ribbon and title bar
#         right = screen_width * 0.88  # Remove right panels
#         bottom = screen_height * 0.88  # Remove bottom status/command area
        
#         cropped_screenshot = screenshot.crop((left, top, right, bottom))
        
#         # Save cropped screenshot
#         output_image_path = os.path.join(tempfile.gettempdir(), "nx_screenshot.png")
#         cropped_screenshot.save(output_image_path)
        
#         # Close NX after capturing
#         if nx_window:
#             try:
#                 # Send ALT+F4 to close NX
#                 nx_window.activate()
#                 pyautogui.hotkey('alt', 'f4')
#                 time.sleep(1)
#                 # Handle any save dialog by pressing 'n' for No
#                 pyautogui.press('n')
#             except Exception as e:
#                 st.warning(f"Could not gracefully close NX: {e}")
#                 # Force kill the process if graceful close failed
#                 try:
#                     nx_process.kill()
#                 except:
#                     pass

#         return output_image_path

#     except Exception as e:
#         st.error(f"Error opening NX or capturing screenshot: {e}")
#         return None
    
# ---------- UPDATED PDF Report Generator ----------
def generate_pdf(pdf_info, excel_df, cad_image_paths, nx_model_groups, output_path):
    """
    UPDATED: Enhanced PDF generation with grid layout for NX screenshots
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

            # NEW: Tyre Wear Analysis Section
    if hasattr(st.session_state, 'wear_analysis_results') and st.session_state.get('wear_analysis_completed'):
        elements.append(Paragraph("Tyre Wear Analysis:", styles["Heading2"]))
        elements.append(Spacer(1, 6))
        
        # Add original tyre image
        if hasattr(st.session_state, 'original_tyre_image') and st.session_state['original_tyre_image']:
            if os.path.exists(st.session_state['original_tyre_image']):
                elements.append(Paragraph("Original Tyre Analysis:", styles["Normal"]))
                elements.append(Spacer(1, 3))
                elements.append(Image(st.session_state['original_tyre_image'], width=200, height=150))
                elements.append(Spacer(1, 10))
        
        # Add wear progression images in 2x2 grid
        wear_results = st.session_state.get('wear_analysis_results', {})
        if wear_results:
            elements.append(Paragraph("Wear Progression Analysis:", styles["Normal"]))
            elements.append(Spacer(1, 6))
            
            # Create 2x2 grid for wear analysis
            wear_stages = ['0%', '25%', '50%', '75%']
            grid_data = []
            
            # Row 1: 0% and 25%
            row1 = []
            for stage in ['0%', '25%']:
                if stage in wear_results and os.path.exists(wear_results[stage]):
                    img = Image(wear_results[stage], width=140, height=100)
                    row1.append(img)
                else:
                    row1.append("")
            grid_data.append(row1)
            
            # Row 2: 50% and 75%
            row2 = []
            for stage in ['50%', '75%']:
                if stage in wear_results and os.path.exists(wear_results[stage]):
                    img = Image(wear_results[stage], width=140, height=100)
                    row2.append(img)
                else:
                    row2.append("")
            grid_data.append(row2)
            
            # Create wear analysis table
            wear_table = Table(grid_data, colWidths=[150, 150])
            wear_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
                ('LEFTPADDING', (0, 0), (-1, -1), 5),
                ('RIGHTPADDING', (0, 0), (-1, -1), 5),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ]))
            
            elements.append(wear_table)
            elements.append(Spacer(1, 10))
            
            # Add wear analysis summary table
            wear_summary_data = [
                ['Wear Stage', 'Condition', 'Recommendation'],
                ['0% (New)', 'Excellent', 'Continue normal use'],
                ['25% (Light)', 'Good', 'Monitor regularly'],
                ['50% (Moderate)', 'Fair', 'Plan replacement'],
                ['75% (Heavy)', 'Poor', 'Replace immediately']
            ]
            
            summary_table = Table(wear_summary_data)
            summary_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                # Color code the rows
                ("BACKGROUND", (0, 1), (-1, 1), colors.lightgreen),
                ("BACKGROUND", (0, 2), (-1, 2), colors.lightyellow),
                ("BACKGROUND", (0, 3), (-1, 3), colors.orange),
                ("BACKGROUND", (0, 4), (-1, 4), colors.lightcoral),
            ]))
            
            elements.append(Paragraph("Wear Analysis Summary:", styles["Normal"]))
            elements.append(Spacer(1, 3))
            elements.append(summary_table)
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

    doc.build(elements)


# ---------- ENHANCED STREAMLIT GUI ----------
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
        updated_tab1_section()  # Call the updated function

    with tab2:
        updated_tab2_section()  # Call the updated function

    with tab3:
        updated_tab3_section()  # Call the updated function

# ---------- UPDATED TAB SECTIONS ----------
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
    Enhanced NX section with tyre wear analysis
    """
    st.header("📐 NX 3D Model Capture & Analysis")
    
    # Create sub-tabs for different functionalities
    subtab1, subtab2 = st.tabs(["🔧 Standard Views Capture", "🔍 Tyre Wear Analysis"])
    
    with subtab1:
        # Original functionality for standard views
        if st.session_state.get('nx_screenshots_captured') and st.session_state.get('nx_screenshot_paths'):
            st.success(f"✅ Currently have {len(st.session_state['nx_screenshot_paths'])} NX model(s) captured")
            
            with st.expander("View Captured NX Models"):
                for i, path in enumerate(st.session_state['nx_screenshot_paths']):
                    if os.path.exists(path):
                        st.image(path, caption=f"NX 3D Model {i+1}")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("🚀 Open NX and Capture Views", key="nx_capture_manual_btn"):
                # [Previous NX capture code remains the same]
                if 'nx_screenshot_paths' in st.session_state:
                    for old_path in st.session_state['nx_screenshot_paths']:
                        try:
                            if os.path.exists(old_path):
                                os.remove(old_path)
                        except:
                            pass
                
                st.session_state['nx_screenshot_paths'] = []
                st.session_state['nx_screenshots_captured'] = False
                st.session_state['nx_model_groups'] = []
                
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
    
    with subtab2:
        # NEW: Tyre Wear Analysis Section
        st.subheader("🔍 Tyre Wear Analysis")
        st.info("📋 This section analyzes tyre wear patterns and generates wear percentage visualizations (0%, 25%, 50%, 75%)")
        
        # Check if wear analysis has been performed
        if st.session_state.get('wear_analysis_completed'):
            st.success("✅ Tyre wear analysis completed!")
            
            # Display wear analysis results
            # Check if wear analysis has been performed
        if st.session_state.get('wear_analysis_completed'):
            st.success("✅ Tyre wear analysis completed!")
            
            # Display wear analysis results in enhanced layout
            if 'wear_analysis_results' in st.session_state and 'original_tyre_image' in st.session_state:
                st.subheader("📊 Complete Tyre Wear Analysis")
                
                # First show the original tyre with current wear percentage indicator
                st.markdown("### 🔍 Original Tyre Analysis")
                original_img = st.session_state['original_tyre_image']
                if os.path.exists(original_img):
                    col_orig1, col_orig2, col_orig3 = st.columns([1, 2, 1])
                    with col_orig2:
                        st.image(original_img, caption="📊 Current Tyre Condition Assessment", use_column_width=True)
                        st.info("👆 This shows the current tyre with wear analysis capabilities")
                
                st.markdown("---")
                
                # Then show the 4 wear progression stages
                st.markdown("### 📈 Tyre Wear Progression Analysis")
                st.info("👇 This shows how the tyre will look at different wear stages")
                
                wear_results = st.session_state['wear_analysis_results']
                
                # Create 2x2 grid for the 4 wear stages
                col1, col2 = st.columns(2)
                
                with col1:
                    # 0% Wear - New Tyre
                    if "0%" in wear_results and os.path.exists(wear_results["0%"]):
                        st.image(wear_results["0%"], caption="🟢 0% Wear - Brand New Tyre", use_column_width=True)
                        st.success("Perfect condition - Maximum grip and safety")
                    
                    # 50% Wear - Moderate
                    if "50%" in wear_results and os.path.exists(wear_results["50%"]):
                        st.image(wear_results["50%"], caption="🟠 50% Wear - Moderate Usage", use_column_width=True)
                        st.warning("Moderate wear - Consider replacement planning")
                
                with col2:
                    # 25% Wear - Light
                    if "25%" in wear_results and os.path.exists(wear_results["25%"]):
                        st.image(wear_results["25%"], caption="🟡 25% Wear - Light Usage", use_column_width=True)
                        st.info("Light wear - Good condition, regular monitoring needed")
                    
                    # 75% Wear - Heavy
                    if "75%" in wear_results and os.path.exists(wear_results["75%"]):
                        st.image(wear_results["75%"], caption="🔴 75% Wear - Heavy Usage", use_column_width=True)
                        st.error("Heavy wear - Immediate replacement recommended!")
                
                # Add wear analysis summary
                st.markdown("---")
                st.markdown("### 📋 Wear Analysis Summary")
                
                summary_col1, summary_col2 = st.columns(2)
                with summary_col1:
                    st.markdown("""
                    **🟢 0% Wear (New)**
                    - Full tread depth
                    - Maximum safety
                    - Optimal performance
                    
                    **🟡 25% Wear (Light)**
                    - Good tread remaining
                    - Safe for continued use
                    - Monitor regularly
                    """)
                
                with summary_col2:
                    st.markdown("""
                    **🟠 50% Wear (Moderate)**
                    - Half tread depth remaining
                    - Plan for replacement
                    - Reduced wet grip
                    
                    **🔴 75% Wear (Heavy)**
                    - Critical tread depth
                    - Replace immediately
                    - Safety risk
                    """)
        
        # Action buttons
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("🔬 Start Tyre Wear Analysis", key="start_wear_analysis"):
                with st.spinner("Starting tyre wear analysis..."):
                    # Capture tyre model from NX
                    original_image = open_nx_for_wear_analysis()
                    
                    if original_image:
                        st.success("✅ Tyre model captured successfully!")
                        st.image(original_image, caption="Original Tyre Model", width=300)
                        
                        # Perform wear analysis
                        with st.spinner("Analyzing tyre wear patterns..."):
                            wear_results = perform_tyre_wear_analysis(original_image)
                            
                            if wear_results:
                                # Store results in session state
                                st.session_state['wear_analysis_results'] = wear_results
                                st.session_state['wear_analysis_completed'] = True
                                st.session_state['original_tyre_image'] = original_image
                                
                                st.success("🎉 Tyre wear analysis completed successfully!")
                                st.balloons()
                                st.rerun()
                            else:
                                st.error("❌ Failed to complete wear analysis")
                    else:
                        st.error("❌ Failed to capture tyre model from NX")
        
        with col2:
            if st.button("🗑️ Clear Wear Analysis", key="clear_wear_analysis"):
                # Clean up wear analysis files
                if 'wear_analysis_results' in st.session_state:
                    for wear_path in st.session_state['wear_analysis_results'].values():
                        try:
                            if os.path.exists(wear_path):
                                os.remove(wear_path)
                        except:
                            pass
                
                if 'original_tyre_image' in st.session_state:
                    try:
                        if os.path.exists(st.session_state['original_tyre_image']):
                            os.remove(st.session_state['original_tyre_image'])
                    except:
                        pass
                
                # Clear session state
                st.session_state['wear_analysis_results'] = {}
                st.session_state['wear_analysis_completed'] = False
                st.session_state['original_tyre_image'] = None
                
                st.success("✅ Wear analysis data cleared!")
                st.rerun()

def updated_tab3_section():
    """
    Updated report generation section with model groups support
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
    if st.button("🚀 Generate Enhanced Report"):
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
            # Generate reports with multiple files
            with st.spinner("Generating enhanced PDF report... Please wait..."):
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
                    
                    # Output path for PDF only
                    pdf_out_path = os.path.join(tmpdir, "Enhanced_Tyre_Report.pdf")
                    
                    # Generate PDF report with model groups
                    generate_pdf(all_pdf_info, combined_excel_df, cad_screenshot_paths, nx_model_groups, pdf_out_path)
                    
                    # Download button for PDF only
                    with open(pdf_out_path, "rb") as f:
                        st.download_button(
                            "📄 Download Enhanced PDF Report", 
                            f, 
                            file_name="Enhanced_Tyre_Report.pdf", 
                            mime="application/pdf"
                        )
            
            st.success("Enhanced PDF report generated successfully!")
            # ---------- RUN THE APPLICATION ----------
if __name__ == "__main__":
    main()