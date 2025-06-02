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
import numpy as np
import matplotlib.pyplot as plt
from skimage import measure, morphology, filters
from scipy import ndimage

# ---------- ENHANCED TYRE WEAR ANALYSIS FUNCTIONS (NEW) ----------

def is_tyre_image(image_path):
    """
    Enhanced validation to ensure uploaded image is actually a tyre image
    Returns True if it's likely a tyre image, False otherwise
    """
    try:
        # Load the image
        img = cv2.imread(image_path)
        if img is None:
            return False
        
        # Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        height, width = gray.shape
        
        # 1. Check for predominantly dark/rubber-like colors
        mean_intensity = np.mean(gray)
        # Tyres are generally darker (rubber is dark)
        is_dark_enough = mean_intensity < 180  # Tyres should be darker than flowcharts/documents
        
        # 2. Check for circular/curved patterns (tyres have curved edges)
        circles = cv2.HoughCircles(gray, cv2.HOUGH_GRADIENT, 1, min(height, width)//4,
                                 param1=50, param2=30, minRadius=min(height, width)//8, 
                                 maxRadius=min(height, width)//2)
        has_curves = circles is not None and len(circles[0]) > 0 if circles is not None else False
        
        # 3. Check edge characteristics (tyres have many irregular edges from tread)
        edges = cv2.Canny(gray, 50, 150)
        edge_density = np.sum(edges) / edges.size
        has_good_edges = edge_density > 0.15  # Tyres have dense edge patterns
        
        # 4. Check for regular geometric shapes (flowcharts have rectangles/boxes)
        # Detect rectangles/squares which are common in flowcharts but not in tyres
        contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        rectangular_shapes = 0
        for cnt in contours:
            approx = cv2.approxPolyDP(cnt, 0.02 * cv2.arcLength(cnt, True), True)
            if len(approx) == 4:  # Rectangle/square
                rectangular_shapes += 1
        
        # Flowcharts/diagrams have many rectangular shapes, tyres don't
        has_too_many_rectangles = rectangular_shapes > 10
        
        # 5. Color distribution check (tyres are mostly black/dark gray)
        hist = cv2.calcHist([gray], [0], None, [256], [0, 256])
        # Check if most pixels are in the darker range (0-100)
        dark_pixels = np.sum(hist[0:100])
        total_pixels = height * width
        dark_ratio = dark_pixels / total_pixels
        is_predominantly_dark = dark_ratio > 0.3  # At least 30% dark pixels for tyres
        
        # 6. Text detection (flowcharts often contain text, tyres usually don't have readable text)
        # Simple text detection using high contrast regions
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        text_like_regions = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, 
                                           cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3)))
        white_pixels = np.sum(text_like_regions == 255)
        white_ratio = white_pixels / total_pixels
        has_too_much_text = white_ratio > 0.6  # Documents/flowcharts have lots of white/text areas
        
        # Scoring system - tyres should meet multiple criteria
        tyre_indicators = [
            is_dark_enough,
            has_curves,
            has_good_edges,
            is_predominantly_dark,
            not has_too_many_rectangles,  # Tyres shouldn't have many rectangles
            not has_too_much_text  # Tyres shouldn't have lots of text/white areas
        ]
        
        tyre_score = sum(tyre_indicators)
        
        # Debug info (you can remove this later)
        print(f"Tyre validation scores:")
        print(f"- Dark enough: {is_dark_enough} (mean: {mean_intensity:.1f})")
        print(f"- Has curves: {has_curves}")
        print(f"- Good edges: {has_good_edges} (density: {edge_density:.3f})")
        print(f"- Predominantly dark: {is_predominantly_dark} (ratio: {dark_ratio:.3f})")
        print(f"- Rectangles: {rectangular_shapes} (too many: {has_too_many_rectangles})")
        print(f"- White ratio: {white_ratio:.3f} (too much: {has_too_much_text})")
        print(f"- Total score: {tyre_score}/6")
        
        # Tyres should score at least 4 out of 6 criteria
        return tyre_score >= 4
        
    except Exception as e:
        print(f"Tyre validation error: {e}")
        return False


def analyze_comprehensive_wear(image_path):
    """
    Comprehensive tyre wear analysis using Computer Vision
    Returns detailed wear information including percentage, condition, and recommendations
    """
    try:
        # Load and preprocess the image
        img = cv2.imread(image_path)
        if img is None:
            return {"error": "Could not load image"}
        
        # Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # Get image dimensions
        height, width = gray.shape
        
        # Step 1: Preprocessing and noise reduction
        # Apply Gaussian blur to reduce noise
        blurred = cv2.GaussianBlur(gray, (5, 5), 0)
        
        # Step 2: Tread pattern detection using multiple techniques
        wear_analysis = analyze_tread_patterns(blurred)
        
        # Step 3: Calculate overall wear percentage
        wear_percentage = calculate_wear_percentage(wear_analysis)
        
        # Step 4: Determine condition and remaining life
        condition_info = determine_condition(wear_percentage)
        
        # Step 5: Generate wear visualization
        wear_map_path = generate_wear_visualization(image_path, wear_analysis, wear_percentage)
        
        return {
            'wear_percentage': round(wear_percentage, 1),
            'condition': condition_info['condition'],
            'remaining_life_months': condition_info['remaining_months'],
            'safety_status': condition_info['safety_status'],
            'recommendations': condition_info['recommendations'],
            'tread_depth_score': wear_analysis['tread_depth_score'],
            'pattern_integrity': wear_analysis['pattern_integrity'],
            'wear_map_path': wear_map_path,
            'analysis_details': {
                'groove_count': wear_analysis['groove_count'],
                'average_groove_width': wear_analysis['avg_groove_width'],
                'pattern_regularity': wear_analysis['pattern_regularity']
            }
        }
        
    except Exception as e:
        return {"error": f"Analysis failed: {str(e)}"}

def analyze_tread_patterns(gray_img):
    """
    Analyze tread patterns to determine wear level
    """
    # Edge detection to find groove patterns
    edges = cv2.Canny(gray_img, 50, 150)
    
    # Morphological operations to enhance groove patterns
    kernel = np.ones((3, 3), np.uint8)
    edges_enhanced = cv2.morphologyEx(edges, cv2.MORPH_CLOSE, kernel)
    
    # Detect horizontal lines (typical tread patterns)
    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (25, 1))
    horizontal_lines = cv2.morphologyEx(edges_enhanced, cv2.MORPH_OPEN, horizontal_kernel)
    
    # Detect vertical grooves
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 25))
    vertical_lines = cv2.morphologyEx(edges_enhanced, cv2.MORPH_OPEN, vertical_kernel)
    
    # Combine patterns
    combined_patterns = cv2.add(horizontal_lines, vertical_lines)
    
    # Analyze groove characteristics
    contours, _ = cv2.findContours(combined_patterns, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # Calculate groove metrics
    groove_count = len(contours)
    groove_areas = [cv2.contourArea(cnt) for cnt in contours if cv2.contourArea(cnt) > 50]
    avg_groove_width = np.mean(groove_areas) if groove_areas else 0
    
    # Tread depth analysis using intensity gradients
    tread_depth_score = analyze_tread_depth(gray_img)
    
    # Pattern integrity (how regular/uniform the pattern is)
    pattern_integrity = calculate_pattern_integrity(combined_patterns)
    
    # Pattern regularity (consistency across the tyre)
    pattern_regularity = calculate_pattern_regularity(combined_patterns)
    
    return {
        'groove_count': groove_count,
        'avg_groove_width': avg_groove_width,
        'tread_depth_score': tread_depth_score,
        'pattern_integrity': pattern_integrity,
        'pattern_regularity': pattern_regularity,
        'edge_density': np.sum(edges_enhanced) / edges_enhanced.size
    }

def analyze_tread_depth(gray_img):
    """
    Analyze tread depth using intensity gradients and shadows
    """
    # Calculate gradients to find depth changes
    grad_x = cv2.Sobel(gray_img, cv2.CV_64F, 1, 0, ksize=3)
    grad_y = cv2.Sobel(gray_img, cv2.CV_64F, 0, 1, ksize=3)
    
    # Calculate gradient magnitude
    gradient_magnitude = np.sqrt(grad_x**2 + grad_y**2)
    
    # Normalize and calculate depth score
    depth_score = np.mean(gradient_magnitude) / 255.0 * 100
    
    return depth_score

def calculate_pattern_integrity(pattern_img):
    """
    Calculate how intact the tread pattern is
    """
    # Count non-zero pixels (pattern areas)
    pattern_pixels = np.count_nonzero(pattern_img)
    total_pixels = pattern_img.size
    
    integrity_score = (pattern_pixels / total_pixels) * 100
    return integrity_score

def calculate_pattern_regularity(pattern_img):
    """
    Calculate pattern regularity across different sections
    """
    height, width = pattern_img.shape
    
    # Divide image into sections and analyze consistency
    sections = []
    section_height = height // 4
    section_width = width // 4
    
    for i in range(4):
        for j in range(4):
            section = pattern_img[i*section_height:(i+1)*section_height, 
                                j*section_width:(j+1)*section_width]
            section_density = np.count_nonzero(section) / section.size
            sections.append(section_density)
    
    # Calculate standard deviation (lower = more regular)
    regularity = 100 - (np.std(sections) * 1000)  # Scale appropriately
    return max(0, min(100, regularity))

def calculate_wear_percentage(analysis):
    """
    Calculate overall wear percentage based on multiple factors
    """
    # Weight different factors
    weights = {
        'tread_depth': 0.4,
        'pattern_integrity': 0.3,
        'groove_density': 0.2,
        'regularity': 0.1
    }
    
    # Normalize scores (higher scores = less wear)
    tread_score = min(100, analysis['tread_depth_score'])
    integrity_score = analysis['pattern_integrity']
    groove_score = min(100, analysis['groove_count'] * 2)  # Adjust multiplier as needed
    regularity_score = analysis['pattern_regularity']
    
    # Calculate weighted average of "good condition" scores
    good_condition_score = (
        weights['tread_depth'] * tread_score +
        weights['pattern_integrity'] * integrity_score +
        weights['groove_density'] * groove_score +
        weights['regularity'] * regularity_score
    )
    
    # Convert to wear percentage (inverse of good condition)
    wear_percentage = 100 - good_condition_score
    
    # Ensure reasonable bounds
    return max(0, min(100, wear_percentage))

def determine_condition(wear_percentage):
    """
    Determine tyre condition and provide recommendations
    """
    if wear_percentage <= 25:
        condition = "Excellent"
        safety_status = "Safe"
        remaining_months = 24
        recommendations = "Tyre is in excellent condition. Continue regular maintenance and rotation."
        
    elif wear_percentage <= 50:
        condition = "Good"
        safety_status = "Safe"
        remaining_months = 12
        recommendations = "Tyre is in good condition. Monitor wear patterns and consider rotation."
        
    elif wear_percentage <= 75:
        condition = "Fair"
        safety_status = "Caution"
        remaining_months = 6
        recommendations = "Tyre shows moderate wear. Plan for replacement within 6 months. Check alignment."
        
    else:
        condition = "Poor"
        safety_status = "Replace Soon"
        remaining_months = 1
        recommendations = "Tyre requires immediate replacement. Unsafe for continued use."
    
    return {
        'condition': condition,
        'safety_status': safety_status,
        'remaining_months': remaining_months,
        'recommendations': recommendations
    }

def generate_wear_visualization(original_image_path, analysis, wear_percentage):
    """
    Generate a visual wear map similar to the reference image
    """
    try:
        # Create visualization
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(12, 8))
        
        # Original image
        original_img = cv2.imread(original_image_path)
        original_rgb = cv2.cvtColor(original_img, cv2.COLOR_BGR2RGB)
        
        # Create wear level representations
        wear_levels = [0, 25, 50, 75]
        colors = ['green', 'yellow', 'orange', 'red']
        
        # Determine current wear level
        current_level = min(3, int(wear_percentage // 25))
        
        axes = [ax1, ax2, ax3, ax4]
        
        for i, (level, color, ax) in enumerate(zip(wear_levels, colors, axes)):
            # Create a representation for each wear level
            wear_img = create_wear_level_image(original_rgb, level, color)
            ax.imshow(wear_img)
            
            # Highlight current level
            if i == current_level:
                ax.set_title(f"{level}% - CURRENT ({wear_percentage:.1f}%)", 
                           fontweight='bold', color='red', fontsize=12)
                ax.add_patch(plt.Rectangle((0, 0), wear_img.shape[1], wear_img.shape[0], 
                                         fill=False, edgecolor='red', linewidth=3))
            else:
                ax.set_title(f"{level}%", fontsize=10)
            
            ax.axis('off')
        
        plt.tight_layout()
        plt.suptitle(f"Tyre Wear Analysis - {wear_percentage:.1f}% Worn", 
                    fontsize=14, fontweight='bold', y=0.98)
        
        # Save visualization
        output_path = os.path.join(tempfile.gettempdir(), "wear_analysis_visualization.png")
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        
        return output_path
        
    except Exception as e:
        print(f"Visualization error: {e}")
        return None

def create_wear_level_image(original_img, wear_level, color):
    """
    Create a visual representation of specific wear level
    """
    # Create a copy of the original image
    wear_img = original_img.copy()
    
    # Apply wear effect based on level
    if wear_level > 0:
        # Create wear mask
        gray = cv2.cvtColor(wear_img, cv2.COLOR_RGB2GRAY)
        
        # Simulate wear by reducing contrast and adding noise
        wear_factor = wear_level / 100.0
        
        # Reduce tread definition
        blurred = cv2.GaussianBlur(wear_img, (int(wear_factor * 10) + 1, int(wear_factor * 10) + 1), 0)
        wear_img = cv2.addWeighted(wear_img, 1 - wear_factor * 0.7, blurred, wear_factor * 0.7, 0)
        
        # Add color tint based on wear level
        color_map = {'green': [0, 255, 0], 'yellow': [255, 255, 0], 
                    'orange': [255, 165, 0], 'red': [255, 0, 0]}
        
        if color in color_map:
            tint = np.full_like(wear_img, color_map[color], dtype=np.uint8)
            wear_img = cv2.addWeighted(wear_img, 0.8, tint, 0.2, 0)
    
    return wear_img

def display_wear_analysis_results(result):
    """
    Display comprehensive wear analysis results in Streamlit
    """
    st.subheader("🔍 Comprehensive Tyre Wear Analysis")
    
    # Create columns for better layout
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            label="Wear Percentage", 
            value=f"{result['wear_percentage']}%",
            delta=f"{100 - result['wear_percentage']:.1f}% remaining"
        )
    
    with col2:
        st.metric(
            label="Condition", 
            value=result['condition']
        )
    
    with col3:
        st.metric(
            label="Est. Life Remaining", 
            value=f"{result['remaining_life_months']} months"
        )
    
    # Safety status with color coding
    safety_color = {
        "Safe": "🟢",
        "Caution": "🟡", 
        "Replace Soon": "🔴"
    }
    
    st.info(f"{safety_color.get(result['safety_status'], '⚪')} **Safety Status:** {result['safety_status']}")
    
    # Detailed analysis in expandable section
    with st.expander("📊 Detailed Analysis"):
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**Tread Analysis:**")
            st.write(f"- Tread Depth Score: {result['tread_depth_score']:.1f}")
            st.write(f"- Pattern Integrity: {result['pattern_integrity']:.1f}%")
            st.write(f"- Groove Count: {result['analysis_details']['groove_count']}")
        
        with col2:
            st.write("**Pattern Analysis:**")
            st.write(f"- Pattern Regularity: {result['analysis_details']['pattern_regularity']:.1f}%")
            st.write(f"- Average Groove Width: {result['analysis_details']['average_groove_width']:.1f}")
    
    # Recommendations
    st.warning(f"**Recommendations:** {result['recommendations']}")
    
    # Display wear visualization if available
    if result.get('wear_map_path') and os.path.exists(result['wear_map_path']):
        st.subheader("🗺️ Wear Pattern Visualization")
        st.image(result['wear_map_path'], caption="Tyre Wear Level Comparison", use_column_width=True)

# ---------- UPDATED WEAR ANALYSIS FUNCTION (REPLACES ORIGINAL) ----------
def analyze_wear_image(image_path):
    """
    UPDATED: Store original image path for report generation
    """
    try:
        # First validate if it's a tyre image
        if not is_tyre_image(image_path):
            st.error("🚫 **Image Not Supported**: This doesn't appear to be a tyre image. Please upload a clear photograph of a tyre showing the tread pattern and rubber surface.")
            st.info("💡 **Tip**: Make sure your image shows a real tyre with visible tread patterns, not diagrams, flowcharts, or other objects.")
            return 0
        
        result = analyze_comprehensive_wear(image_path)
        
        if "error" in result:
            st.error(f"Wear analysis error: {result['error']}")
            return 0
        
        # Store original image path in results
        result['original_image_path'] = image_path
        
        # Store detailed results in session state for report generation
        if 'wear_analysis_results' not in st.session_state:
            st.session_state['wear_analysis_results'] = {}
        
        st.session_state['wear_analysis_results'] = result
        
        # Display results in Streamlit
        display_wear_analysis_results(result)
        
        return result['wear_percentage']
        
    except Exception as e:
        st.error(f"Wear analysis failed: {e}")
        return 0
    
# ---------- SESSION STATE INITIALIZATION (NEW) ----------

def initialize_session_state():
    if 'cad_screenshots_captured' not in st.session_state:
        st.session_state['cad_screenshots_captured'] = False
    if 'cad_screenshot_paths' not in st.session_state:
        st.session_state['cad_screenshot_paths'] = []
    if 'multiple_wear_results' not in st.session_state:
        st.session_state['multiple_wear_results'] = []
    if 'wear_image_paths' not in st.session_state:
        st.session_state['wear_image_paths'] = []

# def initialize_session_state():
#     """
#     Initialize all session state variables
#     """
#     session_vars = {
#         'cad_screenshot_captured': False,
#         'cad_screenshot_path': None,
#         'cad_file_uploaded': False,
#         'processing_cad': False,
#         'nx_views': {},
#         'wear_analysis_results': {}
#     }
    
#     for var, default_value in session_vars.items():
#         if var not in st.session_state:
#             st.session_state[var] = default_value

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

        # UPDATED: Crop only the drawing area (excluding toolbars, ribbons, command line)
        # These coordinates target the main drawing viewport area in AutoCAD
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
    
# ---------- NX Functions (UNCHANGED) ----------
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
            
            # Crop to NX viewport only (excluding menus, toolbars, etc.)
            # These coordinates target the main 3D viewport in NX
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
    
# ---------- UPDATED PDF Report Generator ----------
def generate_pdf(pdf_info, excel_df, image_paths, cad_image_paths, output_path):
    """
    UPDATED: Enhanced PDF generation with multiple file support
    """
    doc = SimpleDocTemplate(output_path, pagesize=A4, 
                          topMargin=72, bottomMargin=72, leftMargin=72, rightMargin=72)
    styles = getSampleStyleSheet()
    elements = []
    
    # Title
    elements.append(Paragraph("<b>Enhanced Tyre Design Proposal Report</b>", styles["Title"]))
    elements.append(Spacer(1, 20))
    
    # PDF Content Section
    elements.append(Paragraph("Extracted PDF Content:", styles["Heading2"]))
    elements.append(Spacer(1, 6))
    for i, pg in enumerate(pdf_info):
        if not pg.startswith("Comprehensive Wear Analysis:"):
            elements.append(Paragraph(pg.replace('\n', '<br/>'), styles["Normal"]))
            if i < len(pdf_info) - 1:
                elements.append(Spacer(1, 6))

    elements.append(Spacer(1, 15))
    
    # Enhanced Wear Analysis Section
    if 'wear_analysis_results' in st.session_state and st.session_state['wear_analysis_results']:
        wear_results = st.session_state['wear_analysis_results']
        
        elements.append(Paragraph("Comprehensive Wear Analysis:", styles["Heading2"]))
        elements.append(Spacer(1, 6))
        
        # Add multiple wear images
        elements.append(Paragraph("Tyre Images:", styles["Heading3"]))
        elements.append(Spacer(1, 3))
        
        for i, image_path in enumerate(image_paths):
            if image_path and os.path.exists(image_path):
                elements.append(Paragraph(f"Tyre Image {i+1}:", styles["Normal"]))
                elements.append(Spacer(1, 3))
                elements.append(Image(image_path, width=250, height=180))
                elements.append(Spacer(1, 10))
        
        # Wear analysis results table
        elements.append(Paragraph("Analysis Results:", styles["Heading3"]))
        elements.append(Spacer(1, 3))
        
        wear_data = [
            ["Metric", "Value", "Status"],
            ["Wear Percentage", f"{wear_results['wear_percentage']}%", wear_results['condition']],
            ["Safety Status", wear_results['safety_status'], ""],
            ["Remaining Life", f"{wear_results['remaining_life_months']} months", ""],
            ["Tread Depth Score", f"{wear_results['tread_depth_score']:.1f}", ""],
            ["Pattern Integrity", f"{wear_results['pattern_integrity']:.1f}%", ""]
        ]
        
        wear_table = Table(wear_data)
        wear_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ]))
        elements.append(wear_table)
        elements.append(Spacer(1, 10))
        
        elements.append(Paragraph("Recommendations:", styles["Heading3"]))
        elements.append(Spacer(1, 3))
        elements.append(Paragraph(wear_results['recommendations'], styles["Normal"]))
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
    
    # NX Views Section
    if 'nx_views' in st.session_state and st.session_state['nx_views']:
        elements.append(Paragraph("3D Model Views (NX):", styles["Heading2"]))
        elements.append(Spacer(1, 6))
        for view, path in st.session_state['nx_views'].items():
            if os.path.exists(path):
                elements.append(Paragraph(f"{view.capitalize()} View:", styles["Heading3"]))
                elements.append(Spacer(1, 3))
                elements.append(Image(path, width=300, height=200))
                elements.append(Spacer(1, 10))

    doc.build(elements)

    # ---------- UPDATED PPTX Report Generator ----------
def generate_pptx(pdf_info, excel_df, image_paths, cad_image_paths, output_path):
    """
    UPDATED: Enhanced PPTX generation with proper titles and multiple file support
    """
    prs = Presentation()
    
    # Slide 1: Title slide
    slide1 = prs.slides.add_slide(prs.slide_layouts[0])
    slide1.shapes.title.text = "Enhanced Tyre Design Proposal Report"
    slide1.placeholders[1].text = "Generated using Streamlit + Advanced Computer Vision"

    # Slide 2: PDF Content
    slide2 = prs.slides.add_slide(prs.slide_layouts[1])
    slide2.shapes.title.text = "PDF Content (Specs)"
    # Filter out the wear analysis summary from PDF content
    filtered_pdf_info = [info for info in pdf_info if not info.startswith("Comprehensive Wear Analysis:")]
    tf = slide2.placeholders[1].text_frame
    tf.text = "\n--- PAGE SPLIT ---\n".join(filtered_pdf_info)

    # Slide 3: Enhanced Wear Analysis Summary with Original Images
    if 'wear_analysis_results' in st.session_state and st.session_state['wear_analysis_results']:
        wear_results = st.session_state['wear_analysis_results']
        
        # Create slide with blank layout and add title manually
        slide3 = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
        
        # Add title manually
        title_shape = slide3.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_shape.text_frame
        title_frame.text = "Comprehensive Wear Analysis"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Inches(0.4)
        title_para.font.bold = True
        
        # Add multiple wear analysis images if available
        current_y = Inches(1.5)
        image_count = 0
        
        for i, image_path in enumerate(image_paths):
            if image_path and os.path.exists(image_path) and image_count < 2:  # Limit to 2 images per slide
                slide3.shapes.add_picture(image_path, Inches(0.5 + (image_count * 4.5)), current_y, 
                                        width=Inches(4), height=Inches(3))
                image_count += 1
        
        # Add text box with analysis results
        textbox = slide3.shapes.add_textbox(Inches(1), Inches(5), Inches(8), Inches(2))
        text_frame = textbox.text_frame
        
        # Create wear analysis content
        wear_content = f"""Wear Analysis Results:
• Wear Percentage: {wear_results['wear_percentage']}%
• Condition: {wear_results['condition']}
• Safety Status: {wear_results['safety_status']}
• Remaining Life: {wear_results['remaining_life_months']} months

Recommendations: {wear_results['recommendations']}"""
        
        text_frame.text = wear_content

    # Slide 4: CAD Screenshots (Multiple)
    if cad_image_paths:
        slide4 = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
        
        # Add title manually
        title_shape = slide4.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_shape.text_frame
        title_frame.text = "2D CAD File Screenshots"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Inches(0.4)
        title_para.font.bold = True
        
        # Add CAD images in a grid
        images_per_row = 2
        current_row = 0
        current_col = 0
        
        for i, cad_path in enumerate(cad_image_paths):
            if cad_path and os.path.exists(cad_path):
                x_pos = Inches(0.5 + (current_col * 4.5))
                y_pos = Inches(1.5 + (current_row * 2.5))
                
                slide4.shapes.add_picture(cad_path, x_pos, y_pos, width=Inches(4), height=Inches(2))
                
                current_col += 1
                if current_col >= images_per_row:
                    current_col = 0
                    current_row += 1

    # Slide 5: Excel Specifications
    if not excel_df.empty:
        slide5 = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
        
        # Add title manually
        title_shape = slide5.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_shape.text_frame
        title_frame.text = "Tyre Specifications"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Inches(0.4)
        title_para.font.bold = True
        
        rows, cols = excel_df.shape
        table = slide5.shapes.add_table(rows+1, cols, Inches(0.5), Inches(1.5), 
                                       Inches(8.5), Inches(0.8 + rows * 0.3)).table
        for i, col in enumerate(excel_df.columns):
            table.cell(0, i).text = str(col)
        for r in range(rows):
            for c in range(cols):
                table.cell(r+1, c).text = str(excel_df.iloc[r, c])

    # Slide 6: NX 3D Model Views
    if 'nx_views' in st.session_state and st.session_state['nx_views']:
        slide6 = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
        
        # Add title manually
        title_shape = slide6.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        title_frame = title_shape.text_frame
        title_frame.text = "NX 3D Model Views"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Inches(0.4)
        title_para.font.bold = True
        
        left = Inches(0.5)
        top = Inches(1.5)
        for i, (view, path) in enumerate(st.session_state['nx_views'].items()):
            if os.path.exists(path):
                slide6.shapes.add_picture(path, left + Inches((i % 2) * 4.5), 
                                        top + Inches((i // 2) * 2.5), width=Inches(4))

    prs.save(output_path)


# ---------- ENHANCED STREAMLIT GUI ----------
def main():
    """
    Main Streamlit application with enhanced wear analysis and multiple file support
    """
    st.title("🛞 Enhanced Tyre Report Generator with Advanced Wear Analysis")
    st.markdown("*Features: Multiple PDF/Excel files, Multiple CAD captures, Multiple wear images, NX 3D views, and AI-powered analysis*")

    # Initialize session state
    initialize_session_state()

    # Create tabs with enhanced wear analysis
    tab1, tab2, tab3, tab4 = st.tabs([
        "📸 Capture CAD Drawing", 
        "🔍 Advanced Wear Analysis", 
        "📐 Capture NX 3D Views", 
        "📄 Generate Reports"
    ])

    with tab1:
        updated_tab1_section()  # Call the updated function

    with tab2:
        updated_tab2_section()  # Call the updated function

    with tab3:
        updated_tab3_section()  # Call the updated function

    with tab4:
        updated_tab4_section()  # Call the updated function

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
    Updated wear analysis section with multiple file upload support
    """
    st.header("🔍 Advanced Tyre Wear Analysis")
    st.markdown("Upload multiple tyre wear images for comprehensive AI-powered analysis")
    
    wear_image_files = st.file_uploader("Upload Tyre Wear Images", 
                                      type=["jpg", "png", "jpeg"], 
                                      accept_multiple_files=True,
                                      key="wear_image_uploader_multiple")
    
    if wear_image_files:
        st.info(f"📁 {len(wear_image_files)} wear image(s) uploaded")
        
        # Display uploaded images
        cols = st.columns(min(3, len(wear_image_files)))
        for i, wear_file in enumerate(wear_image_files):
            with cols[i % 3]:
                st.image(wear_file, caption=f"Wear Image {i+1}", width=200)
        
        if st.button("🔍 Analyze All Wear Patterns", key="analyze_all_wear_btn"):
            wear_results_list = []
            temp_image_paths = []
            
            with st.spinner("Analyzing multiple tyre wear patterns... Please wait..."):
                for i, wear_file in enumerate(wear_image_files):
                    st.info(f"Analyzing image {i+1}/{len(wear_image_files)}")
                    
                    # Save uploaded file temporarily
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_file:
                        tmp_file.write(wear_file.read())
                        temp_image_path = tmp_file.name
                        temp_image_paths.append(temp_image_path)
                    
                    # Perform comprehensive wear analysis
                    wear_percentage = analyze_wear_image(temp_image_path)
                    
                    if wear_percentage > 0:
                        wear_results_list.append(st.session_state.get('wear_analysis_results', {}))
                        st.success(f"✅ Analysis completed for image {i+1}")
                    else:
                        st.error(f"❌ Analysis failed for image {i+1}")
            
            if wear_results_list:
                st.session_state['multiple_wear_results'] = wear_results_list
                st.session_state['wear_image_paths'] = temp_image_paths
                st.success(f"🎉 Successfully analyzed {len(wear_results_list)} tyre images!")

def updated_tab3_section():
    """
    Updated NX section with enhanced manual control
    """
    st.header("📐 NX 3D View Capture")
    st.info("💡 Click the button below to open NX. You will get notifications for each view capture.")
    
    if st.button("🚀 Open NX and Capture Views with Manual Control", key="nx_capture_manual_btn"):
        views = open_nx_and_capture_views_manual("", manual_file_open=True)

        if views:
            st.success("✅ All 3D views captured successfully from NX!")
            st.session_state['nx_views'] = views
            
            with st.expander("View All Captured Screenshots"):
                for name, path in views.items():
                    st.image(path, caption=f"{name.capitalize()} View")
        else:
            st.error("Failed to capture 3D views from NX.")

def updated_tab4_section():
    """
    Updated report generation section with multiple file support
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
    
    # Status indicators
    col1, col2, col3 = st.columns(3)
    with col1:
        cad_status = "✅ Ready" if st.session_state.get('cad_screenshots_captured') else "❌ Not captured"
        st.info(f"CAD Drawings: {cad_status}")
    
    with col2:
        wear_status = "✅ Analyzed" if st.session_state.get('multiple_wear_results') else "❌ Not analyzed"
        st.info(f"Wear Analysis: {wear_status}")
    
    with col3:
        nx_status = "✅ Captured" if st.session_state.get('nx_views') else "❌ Not captured"
        st.info(f"NX Views: {nx_status}")
    
    # Generate Report Button
    if st.button("🚀 Generate Enhanced Report"):
        missing_items = []
        if not pdf_files:
            missing_items.append("PDF files")
        if not excel_files:
            missing_items.append("Excel files")
        if not st.session_state.get('cad_screenshots_captured'):
            missing_items.append("CAD drawing captures")
        if not st.session_state.get('multiple_wear_results'):
            missing_items.append("Wear analysis")
        
        if missing_items:
            st.error(f"Missing required items: {', '.join(missing_items)}")
            st.info("Please complete all sections before generating the report.")
        else:
            # Generate reports with multiple files
            with st.spinner("Generating enhanced reports with multiple files... Please wait..."):
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
                    
                    # Get image paths
                    wear_image_paths = st.session_state.get('wear_image_paths', [])
                    cad_screenshot_paths = st.session_state.get('cad_screenshot_paths', [])
                    
                    # Output paths
                    pptx_path = os.path.join(tmpdir, "Enhanced_Tyre_Report.pptx")
                    pdf_out_path = os.path.join(tmpdir, "Enhanced_Tyre_Report.pdf")
                    
                    # Generate reports
                    generate_pdf(all_pdf_info, combined_excel_df, wear_image_paths, cad_screenshot_paths, pdf_out_path)
                    generate_pptx(all_pdf_info, combined_excel_df, wear_image_paths, cad_screenshot_paths, pptx_path)
                    
                    # Download buttons
                    col1, col2 = st.columns(2)
                    with col1:
                        with open(pdf_out_path, "rb") as f:
                            st.download_button(
                                "📄 Download Enhanced PDF Report", 
                                f, 
                                file_name="Enhanced_Tyre_Report.pdf", 
                                mime="application/pdf"
                            )

                    with col2:
                        with open(pptx_path, "rb") as f:
                            st.download_button(
                                "📊 Download Enhanced PPT Report", 
                                f, 
                                file_name="Enhanced_Tyre_Report.pptx", 
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
            
            st.success("🎉 Enhanced reports with multiple files generated successfully!")

# ---------- RUN THE APPLICATION ----------
if __name__ == "__main__":
    main()