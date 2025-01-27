import pandas as pd
from reportlab.platypus import SimpleDocTemplate, Table, Paragraph
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors


def read_excel(file_path):
    try:
        return pd.read_excel(file_path)
    except Exception as e:
        print("Error reading file:", e)
        return None


def generate_report_card(student_id, student_name, scores, total, average):
    pdf_file = f"report_card_{student_id}.pdf"
    doc = SimpleDocTemplate(pdf_file, pagesize=letter)

    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(
        f"Report Card for {student_name} (ID: {student_id})", styles['Title']))

    data = [["Subject Score"]] + [[score] for score in scores] + \
        [["Total", total], ["Average", round(average, 2)]]

    table = Table(data)
    table.setStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ])

    story.append(table)

    doc.build(story)
    print(f"Created {pdf_file}")


def main():
    df = read_excel("student_scores.xlsx")
    if df is None:
        return

    if 'Student ID' not in df or 'Name' not in df or 'Subject Score' not in df:
        print("Missing columns")
        return

    for (student_id, student_name), group in df.groupby(['Student ID', 'Name']):
        scores = group['Subject Score'].tolist()
        total = sum(scores)
        average = total / len(scores)
        generate_report_card(student_id, student_name, scores, total, average)


if __name__ == "__main__":
    main()



'''
Brief Explaination of Approach :

Libraries:pandas,reportlab

pandas used to read and manipulate file
report lab used to create pdf

The script checks if the required columns exist and then groups them by students.Then for each students their scores are totalled and averaged

After that a pdf file is created for each student with their id
'''