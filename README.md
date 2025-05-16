<h1>Tyre Report Generator</h1>

<p>A powerful Python-based application that automates the generation of comprehensive tyre design reports by integrating data from multiple sources including CAD designs, 3D models, technical specifications, and wear analysis.</p>

<h2>🔍 Features</h2>
<ul>
  <li>📸 <strong>Automated CAD Drawing Capture:</strong> Automatically opens AutoCAD, captures drawings, and closes the application</li>
  <li>📐 <strong>3D Model View Extraction:</strong> Captures multiple view angles from NX 3D models</li>
  <li>📄 <strong>PDF Analysis:</strong> Extracts and processes content from specification PDFs</li>
  <li>📊 <strong>Excel Data Integration:</strong> Imports and formats data from Excel specification sheets</li>
  <li>🖼️ <strong>Wear Image Analysis:</strong> Analyzes wear patterns in tyre images using contour detection</li>
  <li>📑 <strong>Report Generation:</strong> Creates professional PDF and PowerPoint reports with integrated data</li>
  <li>🔄 <strong>Streamlined Workflow:</strong> Intuitive tabbed interface with process automation</li>
</ul>

<h2>📋 Requirements</h2>
<ul>
  <li>Python 3.7+</li>
  <li>Streamlit</li>
  <li>pandas</li>
  <li>pdfplumber</li>
  <li>OpenCV (cv2)</li>
  <li>ReportLab</li>
  <li>python-pptx</li>
  <li>PyAutoGUI</li>
  <li>PyGetWindow</li>
  <li>AutoCAD 2023 (for CAD features)</li>
  <li>Siemens NX (for 3D model features)</li>
</ul>

<h2>🚀 Installation</h2>
<p><strong>Clone the repository:</strong></p>
<pre><code>git clone https://github.com/yourusername/tyre-report-generator.git
cd tyre-report-generator</code></pre>

<p><strong>Install required packages:</strong></p>
<pre><code>pip install -r requirements.txt</code></pre>

<p>Ensure you have AutoCAD 2023 installed at:</p>
<code>C:/Program Files/Autodesk/AutoCAD 2023/acad.exe</code>

<p>Ensure you have Siemens NX installed at:</p>
<code>D:/abcde/NXBIN/ugraf.exe</code> (modify path in code if different)

<h2>📌 Usage</h2>
<p><strong>Run the application with Streamlit:</strong></p>
<pre><code>streamlit run main.py</code></pre>

<h3>Workflow</h3>
<ul>
  <li><strong>Capture CAD Drawing:</strong>
    <ul>
      <li>Upload a DWG file</li>
      <li>The system will automatically open AutoCAD, capture the drawing, and close AutoCAD</li>
    </ul>
  </li>
  <li><strong>Capture NX 3D Views:</strong>
    <ul>
      <li>Upload a PRT or STEP file</li>
      <li>Click "Capture 3D Views"</li>
      <li>Follow on-screen instructions for manual file opening</li>
      <li>The system will capture top, front, right, and isometric views</li>
    </ul>
  </li>
  <li><strong>Generate Reports:</strong>
    <ul>
      <li>Upload a PDF specification file</li>
      <li>Upload an Excel file with tyre data</li>
      <li>Upload a wear image for analysis</li>
      <li>Click "Generate Report" to create both PDF and PowerPoint reports</li>
      <li>Download the generated reports</li>
    </ul>
  </li>
</ul>

<h2>🔧 Technical Details</h2>
<p><strong>Key Components:</strong></p>
<ul>
  <li><strong>PDF Extraction:</strong> Extracts text from the first three pages of PDF files</li>
  <li><strong>Excel Integration:</strong> Loads and formats tabular data from Excel files</li>
  <li><strong>Wear Analysis:</strong> Uses OpenCV to detect contours in wear images</li>
  <li><strong>AutoCAD Integration:</strong> Automates the process of opening files and capturing views</li>
  <li><strong>NX Integration:</strong> Semi-automated process for capturing standard 3D views</li>
  <li><strong>Report Generation:</strong> Creates professional reports with integrated data from all sources</li>
</ul>

<p><strong>File Structure:</strong></p>
<ul>
  <li><code>main.py</code>: Main application file</li>
  <li><code>requirements.txt</code>: Package dependencies</li>
</ul>

<h2>⚠️ Important Notes</h2>
<ul>
  <li>The application requires AutoCAD and NX to be installed at the specified paths.</li>
  <li>During CAD and NX operations, do not move your mouse or interact with the computer.</li>
  <li>The NX integration requires manual file opening as noted in the interface.</li>
  <li>Ensure all required files are uploaded before generating reports.</li>
</ul>

<h2>🛠️ Troubleshooting</h2>
<ul>
  <li>If AutoCAD or NX fails to launch, verify the path in the code matches your installation.</li>
  <li>If screenshots appear incorrect, adjust the waiting times in the code to account for slower systems.</li>
  <li>If the application crashes during CAD operations, ensure no other applications are competing for system resources.</li>
</ul>


