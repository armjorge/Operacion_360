import os
from fpdf import FPDF

class PDF(FPDF):
    def header(self):
        self.add_font('Courier New', '', 'cour.ttf', uni=True)
        self.set_font('Courier New', '', 8)
        self.set_y(26.00)

    def footer(self):
        self.set_y(-26.35)
        self.set_font('Courier New', '', 11)
        self.cell(0, 10, '%s' % self.page_no(), 0, 0, 'C')

def txt_to_pdf(directory):
    # Temporarily create .txt files from .xml
    temp_txt_files = []
    for filename in os.listdir(directory):
        if filename.endswith(".xml"):
            src = os.path.join(directory, filename)
            dst = os.path.join(directory, os.path.splitext(filename)[0] + ".txt")
            with open(src, 'r', encoding='utf-8') as f_src, open(dst, 'w', encoding='utf-8') as f_dst:
                f_dst.write(f_src.read())
            temp_txt_files.append(dst)

    # Convert .txt files to PDF
    for temp_txt in temp_txt_files:
        pdf = PDF(format='letter')
        pdf.add_page()
        pdf.add_font('Courier New', '', 'cour.ttf', uni=True)
        pdf.set_left_margin(20)
        pdf.set_right_margin(20)
        pdf.set_top_margin(20)
        pdf.set_auto_page_break(auto=True, margin=20)

        with open(temp_txt, 'r', encoding='utf-8') as f:
            text = f.read()

        pdf.set_font("Courier New", size=10)
        pdf.multi_cell(0, 4.5, text, align='L')

        new_filename = os.path.splitext(os.path.basename(temp_txt))[0] + "_TXT.pdf"
        output_path = os.path.join(directory, new_filename)
        pdf.output(output_path)

    # Delete the temporary .txt files
    for temp_txt in temp_txt_files:
        os.remove(temp_txt)

# Ask user for the directory
directory = input("Enter the folder name: ")
txt_to_pdf(directory)
