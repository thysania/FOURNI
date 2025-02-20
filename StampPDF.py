import fitz  # PyMuPDF
import os
from tkinter import Tk, filedialog

def add_stamp_to_pdf(input_pdf_path, output_pdf_path, stamp_image_path, a4_position, a5_landscape_position, a5_portrait_position):
    """
    Adds a PNG stamp to the last page of a PDF invoice, adjusting the position based on the page size.

    :param input_pdf_path: Path to the input PDF file.
    :param output_pdf_path: Path to save the output PDF file.
    :param stamp_image_path: Path to the PNG stamp image.
    :param a4_position: Tuple (x, y) representing the position for A4 portrait pages.
    :param a5_landscape_position: Tuple (x, y) representing the position for A5 landscape pages.
    """
    # Open the PDF file
    pdf_document = fitz.open(input_pdf_path)
    
    # Select the last page
    page = pdf_document[-1]
    
    # Get the page size
    page_rect = page.rect
    page_width = page_rect.width
    page_height = page_rect.height
    
    # Determine the page type based on dimensions
    # A4 portrait: 595 x 842 points (21.0 x 29.7 cm)
    # A5 portrait: 420 x 595 points (14.8 x 21.0 cm)
    if page_width > page_height:  # Landscape orientation
        if abs(page_width - 595) < 10 and abs(page_height - 420) < 10:  # A5 landscape
            position = a5_landscape_position
        else:
            raise ValueError("Unsupported page size or orientation.")
    else:  # Portrait orientation
        if abs(page_width - 595) < 10 and abs(page_height - 842) < 10:  # A4 portrait
            position = a4_position
        else:
        	if abs(page_width - 420) < 10 and abs(page_heigt - 595) < 10: # A5 portrait
                position= a5_portrait_position
            else
                raise ValueError("Unsupported page size or orientation.")
    
    # Define the rectangle where the image will be placed
    rect = fitz.Rect(position[0], position[1], position[0] + 100, position[1] + 100)  # Adjust size as needed
    
    # Add the image to the page
    page.insert_image(rect, filename=stamp_image_path)
    
    # Save the modified PDF
    pdf_document.save(output_pdf_path)
    pdf_document.close()

def select_files_or_folder():
    """
    Opens a file dialog to select multiple PDF files or a folder containing PDF files.
    """
    root = Tk()
    root.withdraw()  # Hide the root window
    file_paths = filedialog.askopenfilenames(
        title="Select PDF files or a folder",
        filetypes=[("PDF files", "*.pdf")],
        multiple=True
    )
    if not file_paths:
        folder_path = filedialog.askdirectory(title="Select a folder containing PDF files")
        if folder_path:
            file_paths = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(".pdf")]
    return file_paths

# Main script
if __name__ == "__main__":
    # Define positions for A4 portrait and A5 landscape
    a4_position = (400, 50)  # Position for A4 portrait
    a5_landscape_position = (300, 20)  # Position for A5 landscape
    a5_portrait_position = (300, 50) # Position of A5 portrait 

    # Select the stamp image
    stamp_image = filedialog.askopenfilename(
        title="Select the stamp image",
        filetypes=[("PNG files", "*.png")]
    )
    if not stamp_image:
        print("No stamp image selected. Exiting.")
        exit()

    # Select PDF files or a folder
    pdf_files = select_files_or_folder()
    if not pdf_files:
        print("No PDF files selected. Exiting.")
        exit()

    # Process each PDF file
    for pdf_file in pdf_files:
        output_pdf = os.path.splitext(pdf_file)[0] + "_stamped.pdf"
        try:
            add_stamp_to_pdf(pdf_file, output_pdf, stamp_image, a4_position, a5_landscape_position, a5_portrait_position)
            print(f"Stamp added to {pdf_file}. Saved as {output_pdf}")
        except Exception as e:
            print(f"Failed to process {pdf_file}: {e}")