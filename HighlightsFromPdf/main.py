import fitz  # PyMuPDF
import sys
import os

# Function to extract highlighted text from a PDF
def extract_highlights(pdf_path):
    doc = fitz.open(pdf_path)
    highlights = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        for annot in page.annots():
            if annot.type[0] == 8:  # 8 represents text highlight
                text = annot.info["content"] if "content" in annot.info else ""
                if not text:  # Extract text from highlighted area
                    text = page.get_text("text", clip=annot.rect)
                highlights.append(f"Page {page_num + 1}: {text.strip()}")

    return highlights

# Function to display the extracted highlights
def save_highlights_to_file(highlights, output_file="Rich Dad Poor Dad - Brief.txt"):
    if highlights:
        with open(output_file, "w", encoding="utf-8") as f:
            f.write("\n".join(highlights))
        print(f"Highlights saved to {output_file}")
    else:
        print("No highlights found.")

# Function to open the PDF in a viewer
def open_pdf_viewer(pdf_path):
    if sys.platform == "win32":
        os.startfile(pdf_path)
    elif sys.platform == "darwin":  # macOS
        os.system(f"open {pdf_path}")
    else:  # Linux
        os.system(f"xdg-open {pdf_path}")

# Main function
if __name__ == "__main__":
    pdf_file = input("Enter the path to your PDF file: ").strip()

    if not os.path.exists(pdf_file):
        print("Error: File not found.")
        sys.exit(1)

    print("\nExtracting highlights...\n")
    highlights = extract_highlights(pdf_file)

    if highlights:
        print("\n".join(highlights))  # Display in terminal
        save_highlights_to_file(highlights)
    else:
        print("No highlights found in the document.")

    print("\nOpening the PDF viewer...\n")
    open_pdf_viewer(pdf_file)
