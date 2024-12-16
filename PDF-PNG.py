from pdf2image import convert_from_path
import os

# Define file paths
pdf_path = r"C:\Users\jamai\OneDrive\Desktop\FOR CODING\BIR FORMS\form.pdf"
output_dir = r"C:\Users\jamai\OneDrive\Desktop\FOR CODING\BIR FORMS\output_images"

# Ensure output directory exists
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Set the Poppler path directly in your script
poppler_path = r'C:\Users\jamai\Downloads\Release-24.08.0-0\poppler-24.08.0\Library\bin'

# Ensure that pdf2image knows where to find Poppler
os.environ["PATH"] += os.pathsep + poppler_path

# Check if the PDF file exists
if os.path.exists(pdf_path):
    print(f"PDF file found: {pdf_path}")
else:
    print(f"PDF file not found: {pdf_path}")

try:
    # Convert PDF to images
    pages = convert_from_path(pdf_path, 300)
    print(f"Converted {len(pages)} pages.")

    # Save images to the output directory
    for i, page in enumerate(pages):
        image_path = os.path.join(output_dir, f'page_{i + 1}.png')
        page.save(image_path, 'PNG')
        print(f"Saved page {i + 1} as image: {image_path}")
except Exception as e:
    print(f"An error occurred: {e}")
