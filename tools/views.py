from django.shortcuts import render
import fitz  # PyMuPDF
from django.http import HttpResponse
from PyPDF2 import PdfMerger
from pdf2docx import Converter
import os
from django.core.files.storage import default_storage
from django.conf import settings
import pdfplumber
import pandas as pd
from django.core.files.storage import FileSystemStorage

def dashboard(request):
    return render(request, "dashboard.html")

def pdf_text_extractor(request):
    text = None
    if request.method == 'POST' and request.FILES.get('pdf_file'):
        pdf_file = request.FILES['pdf_file']
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        text = "\n".join([page.get_text() for page in doc])
        doc.close()
    return render(request, "pdf_text_extractor.html", {'extracted_text': text})


def pdf_merger(request):
    if request.method == 'POST' and request.FILES.getlist('pdf_files'):
        pdf_files = request.FILES.getlist('pdf_files')
        merger = PdfMerger()

        for pdf in pdf_files:
            merger.append(pdf)

        response = HttpResponse(content_type='application/pdf')
        response['Content-Disposition'] = 'attachment; filename="merged.pdf"'
        merger.write(response)
        merger.close()
        return response

    return render(request, 'pdf_merger.html')


def pdf_to_word(request):
    converted = False
    docx_file_url = None

    if request.method == 'POST' and request.FILES.get('pdf_file'):
        pdf_file = request.FILES['pdf_file']
        pdf_path = default_storage.save('temp/' + pdf_file.name, pdf_file)
        pdf_full_path = os.path.join(settings.MEDIA_ROOT, pdf_path)

        docx_output = pdf_full_path.replace('.pdf', '.docx')
        cv = Converter(pdf_full_path)
        cv.convert(docx_output)
        cv.close()

        docx_file_url = settings.MEDIA_URL + pdf_path.replace('.pdf', '.docx')
        converted = True

    return render(request, 'pdf_to_word.html', {
        'converted': converted,
        'docx_file_url': docx_file_url
    })


def pdf_to_excel(request):
    converted = False
    excel_file_url = None

    if request.method == 'POST' and request.FILES.get('pdf_file'):
        pdf_file = request.FILES['pdf_file']

        # ✅ Save file inside media/temp
        fs = FileSystemStorage(location=settings.MEDIA_ROOT / 'temp', base_url=settings.MEDIA_URL + 'temp/')
        filename = fs.save(pdf_file.name, pdf_file)
        pdf_full_path = fs.path(filename)

        all_tables = []

        with pdfplumber.open(pdf_full_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    all_tables.append(df)

        if all_tables:
            combined_df = pd.concat(all_tables, ignore_index=True)
            excel_path = pdf_full_path.replace('.pdf', '.xlsx')
            combined_df.to_excel(excel_path, index=False)

            excel_file_url = fs.base_url + os.path.basename(excel_path)
            converted = True

    return render(request, 'pdf_to_excel.html', {
        'converted': converted,
        'excel_file_url': excel_file_url
    })
