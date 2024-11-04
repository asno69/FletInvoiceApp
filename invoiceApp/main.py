import os
from datetime import datetime
import flet as ft
import win32com.client
import pythoncom  # Importiere pythoncom für CoInitialize
from docxtpl import DocxTemplate


def main(page: ft.Page):

    pythoncom.CoInitialize()

    page.title = "Rechnungserstellung"
    page.padding = 10
    page.spacing = 10

    invoice_number_input = ft.TextField(label='Rechnungsnummer')
    date_input = ft.TextField(label='Rechnungsdatum (TT.MM.JJJJ)')
    service_input = ft.TextField(label='Leistungszeitraum')
    salary1_input = ft.TextField(label='Commessa CM0221-003 ATTIVITA\' DI LOGISTICA COSTRUZIONE 706')
    salary2_input = ft.TextField(label='Commessa CM0231-003 ATTIVITA\' DI LOGISTICA COSTRUZIONE 721')
    salary3_input = ft.TextField(label='Commessa CM0255-0003 ATTIVITA\' DI LOGISTICA DISNEY ADVENTURE')

    def create_invoice(e):
        if (salary1_input.value == '' or salary2_input.value == '' or salary3_input.value == '' or
                invoice_number_input.value == '' or date_input.value == '' or service_input.value == ''):
            page.add(ft.Text("Alle Werte müssen gesetzt sein"))
            return

        date = date_input.value
        invoice_number = invoice_number_input.value
        service = service_input.value
        try:
            salary1 = float(salary1_input.value)
            salary2 = float(salary2_input.value)
            salary3 = float(salary3_input.value)
        except ValueError:
            page.add(ft.Text("Bitte geben Sie gültige Zahlen für die Gehälter ein"))
            return

        # Validate and format the date
        try:
            date = datetime.strptime(date, '%d.%m.%Y').strftime('%d.%m.%Y')
        except ValueError:
            page.add(ft.Text('Ungültiges Datum. Bitte im Format TT.MM.JJJJ eingeben.'))
            return

        # Template path and output paths
        template_path = os.path.join(os.path.dirname(__file__), 'Vorlage.docx')
        output_docx_path = os.path.join(os.path.dirname(__file__), f'Rechnung_Nr{invoice_number}.docx')
        output_pdf_path = output_docx_path.replace('.docx', '.pdf')

        # Load template
        doc = DocxTemplate(template_path)

        # Context data for template rendering
        context = {
            'DATE': date,
            'INVOICE_NUMBER': invoice_number,
            'SERVICE': service,
            'SALARY1': f'{salary1:.2f}',
            'SALARY2': f'{salary2:.2f}',
            'SALARY3': f'{salary3:.2f}',
            'BRUTTO': f'{(salary1 + salary2 + salary3) * 1.19:.2f}',
            'STEUER': f'{(salary1 + salary2 + salary3) * 0.19:.2f}',
            'NETTO': f'{(salary1 + salary2 + salary3):.2f}'
        }

        # Render template with context
        doc.render(context)
        doc.save(output_docx_path)

        # Convert DOCX to PDF using Microsoft Word
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(output_docx_path)
        doc.SaveAs(output_pdf_path, FileFormat=17)
        doc.Close()
        word.Quit()

        page.add(ft.Text(f'Rechnung wurde erfolgreich erstellt und unter {output_docx_path} und {output_pdf_path} gespeichert!'))

    submit_button = ft.ElevatedButton(text='Rechnung erstellen', on_click=create_invoice)

    page.add(invoice_number_input, date_input, service_input, salary1_input, salary2_input, salary3_input, submit_button)


# Start der Anwendung
ft.app(target=main)
