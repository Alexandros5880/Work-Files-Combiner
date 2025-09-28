import os
from docx import Document

# Φάκελος που περιέχει τα αρχεία (π.χ. "./word_files")
INPUT_FOLDER = "./word_files"
OUTPUT_FILE = "combined.docx"

def combine_word_files(folder_path, output_file):
    # Φτιάχνουμε το κεντρικό έγγραφο
    combined_doc = Document()

    # Παίρνουμε όλα τα .docx αρχεία και τα ταξινομούμε αλφαβητικά
    files = sorted([f for f in os.listdir(folder_path) if f.endswith(".docx")])

    for idx, file in enumerate(files):
        file_path = os.path.join(folder_path, file)
        print(f"➡️ Προσθήκη: {file_path}")
        
        doc = Document(file_path)

        # Προσθέτουμε το περιεχόμενο παράγραφο-παράγραφο
        for para in doc.paragraphs:
            combined_doc.add_paragraph(para.text)

        # Προσθήκη page break μετά από κάθε αρχείο (εκτός του τελευταίου)
        if idx < len(files) - 1:
            combined_doc.add_page_break()

    # Αποθήκευση τελικού εγγράφου
    combined_doc.save(output_file)
    print(f"\n✅ Ολοκληρώθηκε! Δημιουργήθηκε το αρχείο: {output_file}")

if __name__ == "__main__":
    combine_word_files(INPUT_FOLDER, OUTPUT_FILE)

