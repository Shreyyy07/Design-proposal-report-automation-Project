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
    """
    Initialize all session state variables
    """
    session_vars = {
        'cad_screenshot_captured': False,
        'cad_screenshot_path': None,
        'cad_file_uploaded': False,
        'processing_cad': False,
        'nx_views': {},
        'wear_analysis_results': {}
    }
    
    for var, default_value in session_vars.items():
        if var not in st.session_state:
            st.session_state[var] = default_value

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
    
# ---------- NX Functions (UNCHANGED) ----------
def open_nx_and_capture_views(prt_file_path="", manual_file_open=True):
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
        st.warning("Please manually open your 3D file in NX now.")
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
    
# ---------- UPDATED PDF Report Generator ----------
def generate_pdf(pdf_info, excel_df, image_path, cad_image_path, output_path):
    """
    UPDATED: Enhanced PDF generation with comprehensive wear analysis and better formatting
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
    elements.append(Spacer(1, 6))  # Small space after heading
    for i, pg in enumerate(pdf_info):
        # Skip the wear analysis summary that was added separately
        if not pg.startswith("Comprehensive Wear Analysis:"):
            elements.append(Paragraph(pg.replace('\n', '<br/>'), styles["Normal"]))
            if i < len(pdf_info) - 1:  # Don't add spacer after last item
                elements.append(Spacer(1, 6))

    elements.append(Spacer(1, 15))
    
    # Enhanced Wear Analysis Section - Keep everything together
    if 'wear_analysis_results' in st.session_state and st.session_state['wear_analysis_results']:
        wear_results = st.session_state['wear_analysis_results']
        
        # Add heading
        elements.append(Paragraph("Comprehensive Wear Analysis:", styles["Heading2"]))
        elements.append(Spacer(1, 6))
        
        # Add original wear image first (right after heading)
        if image_path and os.path.exists(image_path):
            elements.append(Paragraph("Original Tyre Image:", styles["Heading3"]))
            elements.append(Spacer(1, 3))
            elements.append(Image(image_path, width=250, height=180))
            elements.append(Spacer(1, 10))
        
        # Wear summary table
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
        
        # Recommendations
        elements.append(Paragraph("Recommendations:", styles["Heading3"]))
        elements.append(Spacer(1, 3))
        elements.append(Paragraph(wear_results['recommendations'], styles["Normal"]))
        elements.append(Spacer(1, 10))
        
        # Wear visualization
        if (wear_results.get('wear_map_path') and 
            os.path.exists(wear_results['wear_map_path'])):
            elements.append(Paragraph("Wear Pattern Analysis:", styles["Heading3"]))
            elements.append(Spacer(1, 3))
            elements.append(Image(wear_results['wear_map_path'], width=400, height=300))
            elements.append(Spacer(1, 15))

    # CAD Screenshot Section
    if cad_image_path:
        elements.append(Paragraph("2D CAD File Screenshot:", styles["Heading2"]))
        elements.append(Spacer(1, 6))
        elements.append(Image(cad_image_path, width=300, height=200))
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
def generate_pptx(pdf_info, excel_df, image_path, cad_image_path, output_path):
    """
    UPDATED: Enhanced PPTX generation with original tyre image in wear analysis slide
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

    # Slide 3: Enhanced Wear Analysis Summary with Original Image
    if 'wear_analysis_results' in st.session_state and st.session_state['wear_analysis_results']:
        wear_results = st.session_state['wear_analysis_results']
        
        # Use layout 5 (Title and Content) instead of layout 6 (blank)
        slide3 = prs.slides.add_slide(prs.slide_layouts[5])
        
        # Add title using text box since layout 5 might not have title placeholder
        title_box = slide3.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
        title_frame = title_box.text_frame
        title_frame.text = "Comprehensive Wear Analysis"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Inches(0.3)
        title_para.font.bold = True
        
        # Add original tyre image on the left
        if image_path and os.path.exists(image_path):
            slide3.shapes.add_picture(image_path, Inches(0.5), Inches(1.5), width=Inches(4), height=Inches(3))
        
        # Add text box with analysis results on the right
        textbox = slide3.shapes.add_textbox(Inches(5), Inches(1.5), Inches(4.5), Inches(4.5))
        text_frame = textbox.text_frame
        
        # Create wear analysis content
        wear_content = f"""Wear Analysis Results:
• Wear Percentage: {wear_results['wear_percentage']}%
• Condition: {wear_results['condition']}
• Safety Status: {wear_results['safety_status']}
• Remaining Life: {wear_results['remaining_life_months']} months
• Tread Depth Score: {wear_results['tread_depth_score']:.1f}
• Pattern Integrity: {wear_results['pattern_integrity']:.1f}%

Recommendations:
{wear_results['recommendations']}"""
        
        text_frame.text = wear_content

    # Slide 4: Wear Visualization (if available)
    if ('wear_analysis_results' in st.session_state and 
        st.session_state['wear_analysis_results'].get('wear_map_path') and
        os.path.exists(st.session_state['wear_analysis_results']['wear_map_path'])):
        
        slide4 = prs.slides.add_slide(prs.slide_layouts[5])
        
        # Add title using text box
        title_box = slide4.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
        title_frame = title_box.text_frame
        title_frame.text = "Wear Pattern Visualization"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Inches(0.3)
        title_para.font.bold = True
        
        slide4.shapes.add_picture(st.session_state['wear_analysis_results']['wear_map_path'], 
                                Inches(0.5), Inches(1.5), width=Inches(8))

    # Slide 5: CAD Screenshot
    if cad_image_path and os.path.exists(cad_image_path):
        slide5 = prs.slides.add_slide(prs.slide_layouts[5])
        
        # Add title using text box
        title_box = slide5.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
        title_frame = title_box.text_frame
        title_frame.text = "2D CAD File Screenshot"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Inches(0.3)
        title_para.font.bold = True
        
        slide5.shapes.add_picture(cad_image_path, Inches(1), Inches(1.5), width=Inches(6))

    # Slide 6: Excel Specifications
    if not excel_df.empty:
        slide6 = prs.slides.add_slide(prs.slide_layouts[5])
        
        # Add title using text box
        title_box = slide6.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
        title_frame = title_box.text_frame
        title_frame.text = "Tyre Specifications"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Inches(0.3)
        title_para.font.bold = True
        
        rows, cols = excel_df.shape
        table = slide6.shapes.add_table(rows+1, cols, Inches(0.5), Inches(1.5), 
                                       Inches(8.5), Inches(0.8 + rows * 0.3)).table
        for i, col in enumerate(excel_df.columns):
            table.cell(0, i).text = str(col)
        for r in range(rows):
            for c in range(cols):
                table.cell(r+1, c).text = str(excel_df.iloc[r, c])

    # Slide 7: NX 3D Model Views
    if 'nx_views' in st.session_state and st.session_state['nx_views']:
        slide7 = prs.slides.add_slide(prs.slide_layouts[5])
        
        # Add title using text box
        title_box = slide7.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
        title_frame = title_box.text_frame
        title_frame.text = "NX 3D Model Views"
        title_para = title_frame.paragraphs[0]
        title_para.font.size = Inches(0.3)
        title_para.font.bold = True
        
        left = Inches(0.5)
        top = Inches(1.5)
        for i, (view, path) in enumerate(st.session_state['nx_views'].items()):
            if os.path.exists(path):
                slide7.shapes.add_picture(path, left + Inches((i % 2) * 4.5), 
                                        top + Inches((i // 2) * 2.5), width=Inches(4))

    prs.save(output_path)


# ---------- ENHANCED STREAMLIT GUI ----------
def main():
    """
    Main Streamlit application with enhanced wear analysis
    """
    st.title("🛞 Enhanced Tyre Report Generator with Advanced Wear Analysis")
    st.markdown("*Features: PDF extraction, Excel data, CAD capture, NX 3D views, and AI-powered wear analysis*")

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
        st.header("📸 AutoCAD Drawing Capture")
        cad_file = st.file_uploader("Upload 2D CAD File (.dwg)", type=["dwg"], 
                                 key="cad_uploader",
                                 on_change=lambda: setattr(st.session_state, 'cad_file_uploaded', True))
        
        # Auto-process CAD file when uploaded
        if (st.session_state['cad_file_uploaded'] and 
            not st.session_state['processing_cad'] and 
            not st.session_state['cad_screenshot_captured']):
            
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
            st.rerun()
        
        # Show captured screenshot
        if st.session_state['cad_screenshot_captured']:
            st.success("✅ CAD drawing has been captured and is ready for report generation")
            st.image(st.session_state['cad_screenshot_path'], caption="Captured CAD Drawing")
            
            if st.button("Capture Again"):
                st.session_state['cad_screenshot_captured'] = False
                st.session_state['cad_screenshot_path'] = None
                st.rerun()

    with tab2:
        st.header("🔍 Advanced Tyre Wear Analysis")
        st.markdown("Upload a tyre wear image for comprehensive AI-powered analysis")
        
        wear_image_file = st.file_uploader("Upload Tyre Wear Image", 
                                         type=["jpg", "png", "jpeg"], 
                                         key="wear_image_uploader")
        
        if wear_image_file:
            # Display uploaded image
            st.image(wear_image_file, caption="Uploaded Wear Image", width=400)
            
            if st.button("🔍 Analyze Wear Pattern", key="analyze_wear_btn"):
                # Save uploaded file temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_file:
                    tmp_file.write(wear_image_file.read())
                    temp_image_path = tmp_file.name
                
                with st.spinner("Analyzing tyre wear pattern... Please wait..."):
                    # Perform comprehensive wear analysis
                    wear_percentage = analyze_wear_image(temp_image_path)
                
                if wear_percentage > 0:
                    st.success("✅ Wear analysis completed successfully!")
                else:
                    st.error("❌ Wear analysis failed. Please try with a different image.")

    with tab3:
      st.header("📐 NX 3D View Capture")
      st.info("💡 Click the button below to open NX. You can then manually open your 3D file from within NX.")
    
      if st.button("🚀 Open NX and Capture Views", key="nx_capture_btn"):
        with st.spinner("Opening NX and preparing for view capture... Please manually open your 3D file in NX."):
            views = open_nx_and_capture_views("", manual_file_open=True)

        if views:
            st.success("✅ Captured 3D views from NX successfully!")
            st.session_state['nx_views'] = views
            
            with st.expander("View Captured Screenshots"):
                for name, path in views.items():
                    st.image(path, caption=f"{name.capitalize()} View")
        else:
            st.error("Failed to capture 3D views from NX.")

    with tab4:
        st.header("📄 Generate Enhanced Reports")
        
        pdf_file = st.file_uploader("Upload PDF File (Specs)", type=["pdf"])
        excel_file = st.file_uploader("Upload Excel File (Specifications)", type=["xlsx"])
        
        # Status indicators
        col1, col2, col3 = st.columns(3)
        with col1:
            cad_status = "✅ Ready" if st.session_state['cad_screenshot_captured'] else "❌ Not captured"
            st.info(f"CAD Drawing: {cad_status}")
        
        with col2:
            wear_status = "✅ Analyzed" if st.session_state.get('wear_analysis_results') else "❌ Not analyzed"
            st.info(f"Wear Analysis: {wear_status}")
        
        with col3:
            nx_status = "✅ Captured" if st.session_state.get('nx_views') else "❌ Not captured"
            st.info(f"NX Views: {nx_status}")
        
        # Generate Report Button
        if st.button("🚀 Generate Enhanced Report"):
            # Check required files
            missing_items = []
            if not pdf_file:
                missing_items.append("PDF file")
            if not excel_file:
                missing_items.append("Excel file")
            if not st.session_state['cad_screenshot_captured']:
                missing_items.append("CAD drawing capture")
            if not st.session_state.get('wear_analysis_results'):
                missing_items.append("Wear analysis")
            
            if missing_items:
                st.error(f"Missing required items: {', '.join(missing_items)}")
                st.info("Please complete all sections before generating the report.")
            else:
                # Generate reports
                with st.spinner("Generating enhanced reports... Please wait..."):
                    with tempfile.TemporaryDirectory() as tmpdir:
                        # Save uploaded files
                        pdf_path = os.path.join(tmpdir, pdf_file.name)
                        excel_path = os.path.join(tmpdir, excel_file.name)
                        pptx_path = os.path.join(tmpdir, "Enhanced_Tyre_Report.pptx")
                        pdf_out_path = os.path.join(tmpdir, "Enhanced_Tyre_Report.pdf")

                        with open(pdf_path, "wb") as f:
                            f.write(pdf_file.read())
                        with open(excel_path, "wb") as f:
                            f.write(excel_file.read())

                        # Extract data
                        pdf_info = extract_pdf_info(pdf_path)
                        excel_df = read_excel_data(excel_path)
                        
                        # Add wear analysis summary to PDF info
                        if st.session_state.get('wear_analysis_results'):
                            wear_summary = f"Comprehensive Wear Analysis: {st.session_state['wear_analysis_results']['wear_percentage']}% worn, Condition: {st.session_state['wear_analysis_results']['condition']}"
                            pdf_info.append(wear_summary)

                        # Get paths for images
                        wear_image_path = None
                        if st.session_state.get('wear_analysis_results'):
                            # Use original image path if available
                            wear_image_path = st.session_state['wear_analysis_results'].get('original_image_path')
                        
                        cad_screenshot_path = st.session_state['cad_screenshot_path']

                        # Generate reports
                        generate_pdf(pdf_info, excel_df, wear_image_path, cad_screenshot_path, pdf_out_path)
                        generate_pptx(pdf_info, excel_df, wear_image_path, cad_screenshot_path, pptx_path)

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
                
                st.success("🎉 Enhanced reports generated successfully!")

# ---------- RUN THE APPLICATION ----------
if __name__ == "__main__":
    main()