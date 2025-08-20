# Color-page-splitter
splits pdf/word document pages according to presence of color images/text in pages. suitable for separating jobs for color and B/W printer

=============================

This folder contains the fully portable version of the PrintSplitter tool.

🖥️ What It Does:
-----------------
- Analyzes a PDF or Word document
- Identifies which pages contain color vs. black & white content
- Splits those into two separate PDF files
- Optionally generates a CSV with page-wise color statistics
- Allows previewing individual pages before export

🚀 How to Use:
--------------
1. Open the folder `PrintSplitter_Portable`.
2. Double-click `PrintSplitter.exe` to launch the app.
3. Load your PDF or DOCX file using the "Load File" button.
4. After analysis, color and black-and-white pages will be displayed in a table.
5. Click any row to preview that page on the right.

🎛️ Customization: Sliders, Save, Reset
--------------------------------------

You can fine-tune the analysis behavior using three sliders:

1. **Color Detection Sensitivity**
   - Controls how strictly color is distinguished from grayscale.
   - Higher = more sensitive to small color variations.

2. **Minimum Colored Pixel %**
   - The minimum percentage of pixels on a page that must be colored to count as a color page.
   - Default is 0% (any amount of color counts as color page).

3. **Pixel Sampling Step**
   - Controls how frequently pixels are sampled for speed vs. accuracy.
   - Higher step = faster but less accurate.

✅ These sliders apply immediately to the next analysis.

💾 Save Settings:
-----------------
- Use the **"Save Settings"** button to store your current slider values as default.
- These will be remembered the next time you launch the app.

🔄 Reset Settings:
------------------
- The **"Reset"** button reverts all sliders to the original default values.
- It also removes any saved configuration files from disk.

📂 Technical Notes:
-------------------
- Settings are saved in a config file within the same folder as the EXE.
- The app creates no registry entries and leaves no trace after deletion.

📦 Folder Contents:
-------------------
- PrintSplitter.exe       ← the main app
- *.dll / *.pyd files     ← required support libraries
- README.txt              ← this file

🧰 Troubleshooting:
-------------------
If the app fails to launch:
- Make sure you're launching it from within the folder
- Try running `PrintSplitter.exe` from a command prompt to see any errors
- Ensure your antivirus is not blocking the EXE

Enjoy!

— Developed by Nazmus Sakib, PhD, PE
  Associate Professor, Department of Civil and Environmental Engineering
  Islamic University of Technology (IUT), Bangladesh
