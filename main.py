import os
import pandas as pd
from docx import Document
from docx.shared import Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tkinter import Tk, Label, Button, filedialog, messagebox

def generate_doc(ref, name, old_salary, new_salary, old_position, new_position):
    doc = Document()
    doc.styles['Normal'].paragraph_format.line_spacing = 1.3
    doc.add_paragraph(f"Date: April 10th, 2024").alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    doc.add_paragraph(f" REF:{ref} ").alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    paragraph = doc.add_paragraph(" \n Subject: Review and Appraisal")
    run = paragraph.runs[0]
    run.bold = True
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    doc.add_paragraph(f"Dear {name},").alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    offer_letter_paragraph = doc.add_paragraph()

    def print_line_if_new_position(new_position):
        if pd.notnull(new_position) and new_position.strip():
            run = offer_letter_paragraph.add_run("""We are pleased to formally notify you of changes to your employment contract with Techkraft Inc Pvt. Ltd., effective from January 1st, 2024. As per the recent review and assessment of your contribution, we are pleased to inform you that the following adjustments have been made to your salary and designation to reflect your valuable contributions and dedication to our organization. """)
            run.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            table = doc.add_table(rows=3, cols=3)
            table.style = 'Table Grid'
            for i, header_text in enumerate(["Particular", "Current Details", "New Details"]):
                cell = table.rows[0].cells[i]
                run = cell.paragraphs[0].add_run(header_text)
                run.bold = True
            table.rows[1].cells[0].text = "Basic Salary per month "
            table.rows[1].cells[1].text = "NRs. " + str(old_salary)
            table.rows[1].cells[2].text = "NRs. " + str(new_salary)
            table.rows[2].cells[0].text = "Designation"
            table.rows[2].cells[1].text = old_position
            table.rows[2].cells[2].text = new_position
        else:
            run = offer_letter_paragraph.add_run("""We are pleased to formally notify you of changes to your employment contract with Techkraft Inc Pvt. Ltd., effective from January 1st, 2024. As per the recent review and assessment of your contribution, we are pleased to inform you that the following adjustment have been made to your salary to reflect your valuable contributions and dedication to our organization. \n""")
            run.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            table = doc.add_table(rows=2, cols=3)
            table.style = 'Table Grid'
            for i, header_text in enumerate(["Particular", "Current Details", "New Details"]):
                cell = table.rows[0].cells[i]
                run = cell.paragraphs[0].add_run(header_text)
                run.bold = True
            table.rows[1].cells[0].text = "Salary"
            table.rows[1].cells[1].text = "NRs. " + str(old_salary)
            table.rows[1].cells[2].text = "NRs. " + str(new_salary)

    print_line_if_new_position(new_position)
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(5)
    conclusion = ("\n It is important to note that apart from these modifications, all other terms and conditions of your "
                  "original employment contract will remain unchanged and in full effect. This encompasses your benefits, "
                  "working hours, and leave entitlements, among other provisions.")
    doc.add_paragraph(conclusion).alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    doc.add_paragraph("""\nShould you require any clarification or have any queries regarding these adjustments, please do not hesitate to reach out to the People and Culture Department. """)
    doc.add_paragraph("""\nThank you for your ongoing dedication and contributions to Techkraft Inc Pvt. Ltd. We eagerly anticipate your continued success within the team. \n""")
    for paragraph in doc.paragraphs:
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    doc.add_paragraph("Yours sincerely, \n\n\n _______________ \n Santosh Koirala \n Executive Director \n Signed Date: ")
    return doc

def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            employee_data = pd.read_excel(file_path)
            employee_data.columns = employee_data.columns.str.strip()
            if not os.path.exists("docs"):
                os.makedirs("docs")
            for index, row in employee_data.iterrows():
                ref = row['ref']
                name = row['Name']
                old_salary = row['Old Salary']
                new_salary = row['New Salary']
                old_position = row['Old Position']
                new_position = row['New Position']
                doc = generate_doc(ref, name, old_salary, new_salary, old_position, new_position)
                doc.save(f"docs/{name.replace(' ', '_')}_Offer_Letter.docx")
            messagebox.showinfo("Success", "Offer letters generated successfully and saved in the 'docs' folder.")
        except Exception as e:
            messagebox.showerror("Error", f"Error occurred: {e}")

def main():
    root = Tk()
    root.title("Offer Letter Generator")
    root.geometry("400x200")
    Label(root, text="Welcome to the Offer Letter Generator", font=("Helvetica", 14)).pack(pady=20)
    Button(root, text="Load Employee Data", command=load_file, font=("Helvetica", 12)).pack(pady=10)
    root.mainloop()

if __name__ == "__main__":
    main()
