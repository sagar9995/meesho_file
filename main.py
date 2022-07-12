from io import StringIO
import os
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from pdfrw import PdfReader, PdfWriter
import pandas as pd 
import re 
from PyPDF2 import PdfFileWriter, PdfFileReader
from tqdm import tqdm 
import warnings

warnings.filterwarnings("ignore")

os.makedirs('input', exist_ok=True)
os.makedirs('output', exist_ok=True)

courier_sort = False
if os.path.exists('true'):
    courier_sort = True 
else:
    courier_sort = False


all_pdf = ["input\\" + x for x in os.listdir('input')]
if len(all_pdf) == 0:
    print("No pdf files found in input folder")
    exit()

def pdf_merger(all_path):
    writer = PdfWriter()
    for path in all_path:
        reader = PdfReader(path)
        for page in reader.pages:
            writer.addpage(page)
    writer.write('output.pdf')

def convert_pdf_to_string(file_path):
    all_page = []
    with open(file_path, 'rb') as in_file:
        parser = PDFParser(in_file)
        doc = PDFDocument(parser)
        for page in PDFPage.create_pages(doc):
            output_string = StringIO()
            rsrcmgr = PDFResourceManager()
            device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            interpreter.process_page(page)
            all_page.append(output_string.getvalue())
    return all_page

#merge pdf and remove blank pages 
pdf_merger(all_pdf)
text = convert_pdf_to_string('output.pdf')
print_page = []
for page_no,txt in enumerate(text):
    all_word = txt.split('\n')
    for i,word in enumerate(all_word):
        if 'Destination Code' in word:
            print_page.append(page_no)

reader_input = PdfReader('output.pdf')
writer_output = PdfWriter()
for page in range(len(reader_input.pages)):
    if page in print_page:
        writer_output.addpage(reader_input.pages[page])
writer_output.write('output.pdf')

def sku_extract(page):
    page = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\xff]', '', page)
    page = page.split('\n')
    try:
        sku_idx = [x for x in range(len(page)) if 'SKU:' in page[x]][0]
        sku = page[sku_idx].replace('SKU: ','')
    except:
        order_num_idx = None
        for i in page:
            if len(i) == 14 and '_' in i:
                order_num_idx = i 
        order_idx = page.index(order_num_idx)
        one_check = False 
        sku_idx = None 
        for i in range(order_idx,100000):
            if len(page[i]) > 0 and one_check == False:
                one_check = True 
            else:
                if one_check == True and len(page[i]) != 0:
                    sku_idx = i
                    break 
        sku = ''
        for i in page[sku_idx:]:
            if len(i) != 0:
                sku = sku + ' ' +  i
            else:
                break 
        sku = sku.strip()
    return sku.strip()

def quantity_extract(page):
    page = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\xff]', '', page)
    page = page.split('\n')
    try:
        qty_idx = [x for x in range(len(page)) if 'Quantity:' in page[x]][0]
        qty = page[qty_idx].replace('Quantity: ','').strip()
    except:
        page = [x for x in page if len(x) != 0]
        qty = page[-1]
    return qty.strip()
        

def courier_extract(page):
    page = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\xff]', '', page)
    page = page.split('\n')
    format_type = True 
    for i in page:
        if 'SKU:' in i:
            format_type = False 
    if format_type:
        page = [x.strip() for x in page if len(x) != 0]
        order_idx = page.index('Order Num')
        courier = page[order_idx-1]
    else:
        for i in page:
            if 'Destination Code:' in i:
                dest_idx = page.index(i)
                break
        courier = page[dest_idx-1]
    return courier.strip()

text = convert_pdf_to_string('output.pdf')

df = pd.DataFrame()
for idx, page in enumerate(text):
    sku = sku_extract(page)
    qty = int(quantity_extract(page))
    courier = courier_extract(page)
    df = df.append({'page':idx ,'sku':sku,'qty':qty,'courier':courier},ignore_index=True)

def pdf_cropper(pdf_path):
    with open(pdf_path, 'rb') as f:
        output = PdfFileWriter()
        pdf = PdfFileReader(f)
        for page_no in range(len(pdf.pages)):
            page = pdf.getPage(page_no)
            text = page.extractText()
            text = [x for x in text.split('\n') if len(x) != 0 and 'SKU:' in x]
            width, height = page.mediaBox.upperRight
            if len(text) != 0:
                x, y, w, h = (30, 20, width-50, height - 220)
            else:
                x, y, w, h = (30, 20, width-50, height - 360)
            page_x, page_y = pdf.getPage(page_no).cropBox.getUpperLeft()
            upperLeft = [page_x.as_numeric(), page_y.as_numeric()] 
            new_upperLeft  = (upperLeft[0] + x, upperLeft[1] - y)
            new_lowerRight = (new_upperLeft[0] + w, new_upperLeft[1] - h)
            page.cropBox.upperLeft  = new_upperLeft
            page.cropBox.lowerRight = new_lowerRight
            output.addPage(page)
        path2 = pdf_path.replace('.pdf', '_cropped.pdf')
        with open(path2, 'wb') as f:
            output.write(f)

if courier_sort:
    all_courier = df.courier.unique()
    courier_pages = {}
    for courier in all_courier:
        courier_pages[courier] = []
        temp = df[df.courier == courier]
        temp_sku = temp.sku.unique()
        for sku in sorted(temp_sku, key=lambda v: v.upper()):
            temp_df = temp[temp.sku == sku]
            temp_df = temp_df.sort_values(by='qty')
            courier_pages[courier] = courier_pages[courier] + temp_df.page.tolist()
    reader_input = PdfReader('output.pdf')
    for courier, pages in courier_pages.items():
        writer_output = PdfWriter()
        for page in pages:
            writer_output.addpage(reader_input.pages[page])
        fname = "output/{}.pdf".format(courier)
        writer_output.write(fname)
else:
    all_sku = df.sku.unique()
    for sku in all_sku:
        temp_df = df[df.sku == sku]
        temp_df = temp_df.sort_values(by=['qty'],ascending=True)
        all_page = temp_df.page.values
        reader_input = PdfReader('output.pdf')
        writer_output = PdfWriter()
        for page in all_page:
            writer_output.addpage(reader_input.pages[page])
        fname = "output/{}.pdf".format(sku)
        writer_output.write(fname)

all_pdf = ["output\\" + x for x in os.listdir('output')]
print('cropping all the pdf please wait...')
for i in tqdm(all_pdf):
    pdf_cropper(i)
    os.remove(i)
print('done')
os.remove('output.pdf')