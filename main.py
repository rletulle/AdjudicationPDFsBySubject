import csv
import os
import datetime
import tkinter as tk
from tkinter import filedialog
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import LETTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch, mm

STUDY_COL = 19
SUBJECT_COL = 20
SITE_COL = 25


# ----------------------------------------------------------------------
def add_page_number(canvas, doc):
    """
    Add the page number
    """
    page_num = canvas.getPageNumber()
    text = "Page #%s" % page_num
    canvas.drawRightString(200*mm, 20*mm, text)
# ----------------------------------------------------------------------


def create_pdf_dir(path, row, study, site, subject):
    this_dir = os.path.join(path, study, site)
    if not os.path.exists(this_dir):
        os.makedirs(this_dir)
    return os.path.join(this_dir, subject + ".pdf")


def add_fields(doc, header, row):
    field_count = 0
    for afield in header:
        line = afield + ": " + row[field_count]
        doc.append(Paragraph(line))
        doc.append(Spacer(1, 12))
        field_count += 1
    return doc


def render_pdf(pdf_file_name, header, row, study, site, subject):
    doc = []
    pdf_header = "Study: " + study + ", " + site + ", " + "Subject: " + subject
    pdf = SimpleDocTemplate(pdf_file_name, pagesize=LETTER,
                            rightMargin=0.5 * inch, leftMargin=0.5 * inch,
                            topMargin=1.0 * inch, bottomMargin=1.0 * inch)
    doc.append(Spacer(1, 12))
    doc.append(Paragraph(pdf_header, ParagraphStyle(name='Title', fontFamily='Helvetica-Bold', fontSize=14,
                                                    alignment=TA_CENTER)))
    doc.append(Spacer(1, 20))
    pdf.build(add_fields(doc, header, row), onFirstPage=add_page_number, onLaterPages=add_page_number)


def process_csv(path, csv_reader):
    header = []
    line_count = 0
    date_dir = os.path.join(path, datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S'))
    if not os.path.exists(date_dir):
        os.makedirs(date_dir)
    # after the header is extracted, each row gets rendered as a PDF for that Subject ID
    for row in csv_reader:
        if line_count == 0:
            header = row[:]  # get all of the columns in the header
            line_count += 1
        else:
            line_count += 1
            study = row[STUDY_COL]
            site = row[SITE_COL]
            subject = row[SUBJECT_COL]
            pdf_file_name = create_pdf_dir(date_dir, row, study, site, subject)  # create the PDF for this subject
            render_pdf(pdf_file_name, header, row, study, site, subject)  # render the row of data into that PDF


# Use TK to get a dialog to return the full path and file name of the CVS file to process
# the PDFs will be save in subdirectories under the directory with the CSV file
def get_csv_file():
    root = tk.Tk()
    root.withdraw()
    file_name_with_path = filedialog.askopenfilename()
    path = os.path.dirname(file_name_with_path)  # get only the path portion without the file name
    if file_name_with_path:
        with open(file_name_with_path, newline='') as csv_file:
            csv_reader = csv.reader(csv_file, delimiter=',', quotechar='"')
            process_csv(path, csv_reader)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    get_csv_file()
