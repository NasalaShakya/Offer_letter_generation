import os
import pandas as pd
from docx import Document
from docx.shared import Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tkinter import Tk, Label, Button, filedialog, messagebox, StringVar, Entry, Frame, Listbox
import json
from fpdf import FPDF

# Function to generate Word document
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

# Function to generate PDF document
def generate_pdf(ref, name, old_salary, new_salary, old_position, new_position):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Date: April 10th, 2024", ln=True, align='L')
    pdf.cell(200, 10, txt=f" REF:{ref} ", ln=True, align='L')
    pdf.set_font("Arial", 'B', size=12)
    pdf.cell(200, 10, txt="Subject: Review and Appraisal", ln=True, align='L')
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt=f"Dear {name},", ln=True, align='L')

    def print_line_if_new_position(new_position):
        if pd.notnull(new_position) and new_position.strip():
            pdf.multi_cell(0, 10, """We are pleased to formally notify you of changes to your employment contract with Techkraft Inc Pvt. Ltd., effective from January 1st, 2024. As per the recent review and assessment of your contribution, we are pleased to inform you that the following adjustments have been made to your salary and designation to reflect your valuable contributions and dedication to our organization.""")
            pdf.ln()
            pdf.set_font("Arial", 'B', size=12)
            pdf.cell(63, 10, "Particular", 1)
            pdf.cell(63, 10, "Current Details", 1)
            pdf.cell(63, 10, "New Details", 1)
            pdf.ln()
            pdf.set_font("Arial", size=12)
            pdf.cell(63, 10, "Basic Salary per month", 1)
            pdf.cell(63, 10, f"NRs. {old_salary}", 1)
            pdf.cell(63, 10, f"NRs. {new_salary}", 1)
            pdf.ln()
            pdf.cell(63, 10, "Designation", 1)
            pdf.cell(63, 10, old_position, 1)
            pdf.cell(63, 10, new_position, 1)
        else:
            pdf.multi_cell(0, 10, """We are pleased to formally notify you of changes to your employment contract with Techkraft Inc Pvt. Ltd., effective from January 1st, 2024. As per the recent review and assessment of your contribution, we are pleased to inform you that the following adjustments have been made to your salary to reflect your valuable contributions and dedication to our organization.\n""")
            pdf.ln()
            pdf.set_font("Arial", 'B', size=12)
            pdf.cell(63, 10, "Particular", 1)
            pdf.cell(63, 10, "Current Details", 1)
            pdf.cell(63, 10, "New Details", 1)
            pdf.ln()
            pdf.set_font("Arial", size=12)
            pdf.cell(63, 10, "Salary", 1)
            pdf.cell(63, 10, f"NRs. {old_salary}", 1)
            pdf.cell(63, 10, f"NRs. {new_salary}", 1)
            pdf.ln()

    print_line_if_new_position(new_position)
    conclusion = ("\nIt is important to note that apart from these modifications, all other terms and conditions of your "
                  "original employment contract will remain unchanged and in full effect. This encompasses your benefits, "
                  "working hours, and leave entitlements, among other provisions.")
    pdf.multi_cell(0, 10, conclusion)
    pdf.ln()
    pdf.multi_cell(0, 10, """\nShould you require any clarification or have any queries regarding these adjustments, please do not hesitate to reach out to the People and Culture Department. """)
    pdf.ln()
    pdf.multi_cell(0, 10, """\nThank you for your ongoing dedication and contributions to Techkraft Inc Pvt. Ltd. We eagerly anticipate your continued success within the team. \n""")
    pdf.ln(20)
    pdf.cell(200, 10, txt="Yours sincerely,", ln=True, align='L')
    pdf.cell(200, 10, txt="_______________", ln=True, align='L')
    pdf.cell(200, 10, txt="Sam Browns", ln=True, align='L')
    pdf.cell(200, 10, txt="Executive Director", ln=True, align='L')
    pdf.cell(200, 10, txt="Signed Date:", ln=True, align='L')
    return pdf

# Function to load Excel file
def load_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    excel_file_path.set(file_path)

# Function to load template file
def load_template_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx"), ("PDF files", "*.pdf")])
    template_file_path.set(file_path)

# Function to save configurations
def save_config():
    config = {
        "excel_file": excel_file_path.get(),
        "template_file": template_file_path.get(),
        "output_option": output_option.get()
    }
    with open("config.json", "w") as config_file:
        json.dump(config, config_file)
    messagebox.showinfo("Success", "Configuration saved successfully.")

# Function to load configurations
def load_config():
    try:
        with open("config.json", "r") as config_file:
            config = json.load(config_file)
        excel_file_path.set(config.get("excel_file", ""))
        template_file_path.set(config.get("template_file", ""))
        output_option.set(config.get("output_option", "docx"))
        messagebox.showinfo("Success", "Configuration loaded successfully.")
    except Exception as e:
        messagebox.showerror("Error", f"Error loading configuration: {e}")

# Function to generate and preview letters
def generate_letters():
    try:
        employee_data = pd.read_excel(excel_file_path.get())
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
            if output_option.get() == "docx":
                doc = generate_doc(ref, name, old_salary, new_salary, old_position, new_position)
                doc.save(f"docs/{name.replace(' ', '_')}_Offer_Letter.docx")
            elif output_option.get() == "pdf":
                pdf = generate_pdf(ref, name, old_salary, new_salary, old_position, new_position)
                pdf.output(f"docs/{name.replace(' ', '_')}_Offer_Letter.pdf")
        messagebox.showinfo("Success", "Offer letters generated successfully and saved in the 'docs' folder.")
    except Exception as e:
        messagebox.showerror("Error", f"Error occurred: {e}")

def main():
    global excel_file_path, template_file_path, output_option

    root = Tk()
    root.title("Offer Letter Generator")
    root.geometry("500x400")

    excel_file_path = StringVar()
    template_file_path = StringVar()
    output_option = StringVar(value="docx")

    Label(root, text="Welcome to the Automated Letter Generation System", font=("Helvetica", 14)).pack(pady=10)

    frame = Frame(root)
    frame.pack(pady=10)

    Label(frame, text="Excel File:", font=("Helvetica", 12)).grid(row=0, column=0, sticky='e', padx=10)
    Entry(frame, textvariable=excel_file_path, width=40).grid(row=0, column=1)
    Button(frame, text="Browse", command=load_excel_file, font=("Helvetica", 10)).grid(row=0, column=2, padx=10)

    Label(frame, text="Template File:", font=("Helvetica", 12)).grid(row=1, column=0, sticky='e', padx=10)
    Entry(frame, textvariable=template_file_path, width=40).grid(row=1, column=1)
    Button(frame, text="Browse", command=load_template_file, font=("Helvetica", 10)).grid(row=1, column=2, padx=10)

    Label(root, text="Output Option:", font=("Helvetica", 12)).pack(pady=10)
    options = ["docx", "pdf"]
    for opt in options:
        Button(root, text=opt.upper(), command=lambda opt=opt: output_option.set(opt), font=("Helvetica", 10)).pack(pady=5, side="left", expand=True)

    Button(root, text="Generate and Preview Letters", command=generate_letters, font=("Helvetica", 12)).pack(pady=10)
    Button(root, text="Save Configuration", command=save_config, font=("Helvetica", 12)).pack(pady=5)
    Button(root, text="Load Configuration", command=load_config, font=("Helvetica", 12)).pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    main()
