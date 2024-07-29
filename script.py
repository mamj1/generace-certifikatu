import fitz  # PyMuPDF
import os

# Function to extract and format names from the attendance PDF
def extract_and_format_names(pdf_path):
    doc = fitz.open(pdf_path)
    names = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        lines = text.split('\n')
        for line in lines:
            if ',' in line:
                parts = line.split(',')
                first_name = parts[1].strip()
                last_name = parts[0].strip()
                full_name = f"{first_name} {last_name}"
                names.append(full_name)
    return names

# Function to find the position and size of the text "Brian Lantz" in the template
def get_text_position_and_size(pdf_path, text_to_find):
    doc = fitz.open(pdf_path)
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text_instances = page.search_for(text_to_find)
        if text_instances:
            inst = text_instances[0]
            x0, y0, x1, y1 = inst
            font_size = y1 - y0
            position = fitz.Point(x0, y1)
            return position, font_size
    return None, None

# Function to create a certificate for each name
def create_certificates(names, template_path, output_dir, font_path, position, font_size):
    for name in names:
        # Open the template
        doc = fitz.open(template_path)
        page = doc.load_page(0)

        # Load the font
        font = fitz.Font(fontfile=font_path)

        # Remove the existing "Brian Lantz" text (optional step, depends on template)
        text_instances = page.search_for("Brian Lantz")
        for inst in text_instances:
            page.add_redact_annot(inst)
        page.apply_redactions()

        # Insert the new name
        page.insert_text(position, name, fontsize=font_size, fontfile=font_path, color=(0, 0, 0), fontname="Arial-BoldMT")

        # Save the new certificate
        output_path = os.path.join(output_dir, f"{name.replace(' ', '_')}_certificate.pdf")
        doc.save(output_path)
        doc.close()

# Paths to the files
attendance_pdf_path = "20240626_Lasvit Attendance (1).pdf"
template_pdf_path = "LASVIT AIA Template.pdf"
output_directory = "certificates"
font_path = "C:/Users/KubaMami/Documents/ondra/Arial-BoldMT.otf"  # Update this path to the actual location of the font file

# Ensure the output directory exists
os.makedirs(output_directory, exist_ok=True)

# Extract and format names from the attendance PDF
names = extract_and_format_names(attendance_pdf_path)

# Get the position and font size of the text "Brian Lantz" in the template
position, font_size = get_text_position_and_size(template_pdf_path, "Brian Lantz")

if position and font_size:
    # Create certificates
    create_certificates(names, template_pdf_path, output_directory, font_path, position, font_size)
    print(f"Certificates created for {len(names)} attendees.")
else:
    print("Failed to find the text 'Brian Lantz' in the template.")
