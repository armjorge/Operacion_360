import os
from fpdf import FPDF

def xml_to_wordpad(directory):
    """
    Convert all .xml files in `directory` to PDFs using Courier New (cour.ttf).
    Assumes cour.ttf is in the current working directory.
    """
    class PDF(FPDF):
        def header(self):
            # Register and set header font
            self.add_font('Courier New', '', 'cour.ttf', uni=True)
            self.set_font('Courier New', '', 8)
            self.set_y(26.00)

        def footer(self):
            # Set footer font & page number
            self.set_y(-26.35)
            self.set_font('Courier New', '', 11)
            self.cell(0, 10, str(self.page_no()), 0, 0, 'C')

    # 1) Convert each .xml → .txt
    temp_txt_files = []
    for fname in os.listdir(directory):
        if fname.lower().endswith('.xml'):
            xml_path = os.path.join(directory, fname)
            txt_name = os.path.splitext(fname)[0] + '.txt'
            txt_path = os.path.join(directory, txt_name)
            with open(xml_path, 'r', encoding='utf-8') as fr, \
                 open(txt_path, 'w', encoding='utf-8') as fw:
                fw.write(fr.read())
            temp_txt_files.append(txt_path)

    # 2) Convert each .txt → PDF
    for txt_path in temp_txt_files:
        pdf = PDF(format='letter')
        pdf.add_page()
        pdf.add_font('Courier New', '', 'cour.ttf', uni=True)

        pdf.set_left_margin(20)
        pdf.set_right_margin(20)
        pdf.set_top_margin(20)
        pdf.set_auto_page_break(auto=True, margin=20)

        with open(txt_path, 'r', encoding='utf-8') as f:
            content = f.read()

        pdf.set_font('Courier New', '', 10)
        pdf.multi_cell(0, 4.5, content, align='L')

        out_name = os.path.splitext(os.path.basename(txt_path))[0] + '_TXT.pdf'
        pdf.output(os.path.join(directory, out_name))

    # 3) Clean up temporary .txt files
    for txt_path in temp_txt_files:
        os.remove(txt_path)
    print("\n********\nTodos los XML's fueron transformados a wordpad\n********\n")