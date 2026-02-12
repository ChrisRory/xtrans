import streamlit as st
from pdf2image import convert_from_bytes
from pptx import Presentation
from pptx.util import Inches
from PIL import Image, ImageDraw
import io
import random

# --- Configuration ---
st.set_page_config(page_title="PDF to PPTX & Watermark Remover", layout="wide")

def remove_watermark_and_convert(file_bytes, dpi=100):
    """
    Processes the PDF bytes, removes the watermark, and returns a PPTX byte stream.
    """
    
    # Create a placeholder for status updates
    status_text = st.empty()
    progress_bar = st.progress(0)
    
    status_text.text(f"Step 1: Converting PDF to images ({dpi} DPI)...")
    
    # Convert PDF bytes directly to images (no temp file needed)
    try:
        images = convert_from_bytes(file_bytes.read(), dpi=dpi)
    except Exception as e:
        st.error("Error reading PDF. Please ensure poppler is installed on the server.")
        st.error(f"Details: {e}")
        return None

    total_pages = len(images)
    processed_images = []

    # Step 2: Process each page (Watermark Removal)
    for i, image in enumerate(images):
        page_num = i + 1
        progress = int((i / total_pages) * 50) # First 50% of progress bar
        progress_bar.progress(progress)
        status_text.text(f"Step 2: Removing watermark on page {page_num}/{total_pages}...")

        width, height = image.size
        
        # Define watermark area (Bottom Right 150x35)
        watermark_width = 150
        watermark_height = 35
        
        x1 = width - watermark_width
        y1 = height - watermark_height
        x2 = width
        y2 = height

        # Get reference pixels
        pixel_bottom_right = image.getpixel((width - 1, height - 1))
        pixel_top = image.getpixel((width - 1, y1 - 1)) if y1 > 0 else pixel_bottom_right
        pixel_left = image.getpixel((x1 - 1, height - 1)) if x1 > 0 else pixel_bottom_right

        draw = ImageDraw.Draw(image)

        # Pixel manipulation loop (Your original logic)
        # Note: This is CPU intensive, but preserves your exact logic
        for x in range(x1, x2):
            for y in range(y1, y2):
                choice = random.randint(0, 2)
                if choice == 0:
                    color = pixel_bottom_right
                elif choice == 1:
                    color = pixel_top
                else:
                    color = pixel_left

                # Add slight noise
                if isinstance(color, tuple) and len(color) >= 3:
                    r = max(0, min(255, color[0] + random.randint(-2, 2)))
                    g = max(0, min(255, color[1] + random.randint(-2, 2)))
                    b = max(0, min(255, color[2] + random.randint(-2, 2)))
                    color = (r, g, b) if len(color) == 3 else (r, g, b, color[3])

                draw.point((x, y), fill=color)

        if image.mode != 'RGB':
            image = image.convert('RGB')
        
        processed_images.append(image)

    # Step 3: Create PPTX
    status_text.text("Step 3: Creating PowerPoint presentation...")
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625) 

    for i, image in enumerate(processed_images):
        # Update progress for the second half
        progress = 50 + int((i / total_pages) * 50)
        progress_bar.progress(progress)
        
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)

        # Save image to memory buffer
        img_stream = io.BytesIO()
        image.save(img_stream, format='PNG')
        img_stream.seek(0)

        slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)

    # Save PPTX to memory buffer
    pptx_out = io.BytesIO()
    prs.save(pptx_out)
    pptx_out.seek(0)
    
    progress_bar.progress(100)
    status_text.text("✓ Processing Complete!")
    
    return pptx_out

# --- Main UI ---
st.title("📄 PDF to PPTX Cleaner")
st.markdown("Upload a PDF to remove the bottom-right watermark and convert it to PowerPoint.")

# Sidebar controls
with st.sidebar:
    st.header("Settings")
    dpi_input = st.slider("Quality (DPI)", min_value=72, max_value=200, value=100, step=10, help="Higher DPI = Better quality but larger file size.")

uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    # Display file info
    st.info(f"Filename: {uploaded_file.name} | Size: {uploaded_file.size / 1024 / 1024:.2f} MB")
    
    if st.button("Start Conversion", type="primary"):
        with st.spinner("Processing... Do not close this tab."):
            # Run the conversion
            pptx_result = remove_watermark_and_convert(uploaded_file, dpi=dpi_input)
            
            if pptx_result:
                # Create download button
                st.success("Conversion successful!")
                st.download_button(
                    label="📥 Download PowerPoint (.pptx)",
                    data=pptx_result,
                    file_name=f"{uploaded_file.name.rsplit('.', 1)[0]}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
