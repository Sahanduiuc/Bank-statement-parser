import tabula  
import pandas as pd
from openpyxl import load_workbook
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO


file_path = "/Users/Hatim/Downloads/test.pdf"

"""
file = open(file_path, 'rb')
fileReader = PyPDF2.PdfFileReader(file)
pgObj = fileReader.getPage(0)
st = pgObj.extractText()


def get_pdf_content_lines(pdf_file_path):
    with open(pdf_file_path) as f:
        pdf_reader = PyPDF2.PdfFileReader(f)
        for page in pdf_reader.pages: 
            for line in page.extractText().splitlines():
                yield line

for line in get_pdf_content_lines(file_path):
    print(line)
"""


def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text

st = convert_pdf_to_txt(file_path)
    
rows = tabula.read_pdf(file_path,
                       pages='all',
                       silent=True,
                       pandas_options={
                           'header': None,
                           'error_bad_lines': False,
                           'warn_bad_lines': False
                       })
# converting to list
rows = rows.values.tolist()

# ----- get header row ----- #
def isNaN(num):
    return num != num
# list to contain all headers
headers = []

# iterate over all rows
for row in rows:
    if sum(isNaN(x) for x in row) >=4:
        headers.append(row)

# ----- remove headers ----- #

valid_transactions = []  # list to contain valid transactions

# iterate over all rows
for row in rows:
    if row not in headers:
        valid_transactions.append([x for x in row if str(x) != 'nan'])
        
date = []
desc = []
amt  = []
bal  = []

for trans in valid_transactions:
    date.append(trans[0].split(" ")[0])
    desc.append(trans[0].replace(trans[0].split(" ")[0],"").lstrip())
    amt.append(float(trans[1]))
    if type(trans[2]) == str:
        bal.append(float(trans[2].replace(',', '')))
    else:
        bal.append(trans[2])


statement = {"Date": date, "Description" : desc,
             "Amount" : amt, "Balance" : bal}
c = ["Date", "Description", "Amount", "Balance"]
statement = pd.DataFrame(statement, columns = c)
statement.set_index("Date", inplace = True)
            
beginning_balance = statement.iloc[0]["Balance"] - statement.iloc[0]["Amount"]
ending_balance = statement.iloc[-1]["Balance"]
deposits = statement["Amount"][statement["Amount"] > 0].sum()
withdrawals = statement["Amount"][statement["Amount"] < 0].sum()
account = st.split("Account Number:")[1].split()[0]
summary = pd.DataFrame.from_dict({"Account #": account, 
                                  "Beginning Balance": beginning_balance, 
                                  "Ending Balance" : ending_balance,
                                  "Deposits": deposits,
                                  "Withdrawals": withdrawals}, 
                                    orient = "index", columns = ["Details"])

summary.to_excel("chase.xlsx")
book = load_workbook("chase.xlsx")
writer = pd.ExcelWriter('chase.xlsx', engine='openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
statement.to_excel(writer,startrow=len(summary)+2)
writer.save()
