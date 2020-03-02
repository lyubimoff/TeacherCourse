import pandas as pd
from transliterate import translit
from docxtpl import DocxTemplate
import os
import comtypes.client

file = r'D:\Dev\TeacherCourse\mdl_course.csv'
doc = DocxTemplate(r'D:\Dev\TeacherCourse\Crt_tmplt_pythn.docx')
headers = ['fullname','lastname','firstname','middlename','grade','sex','email']
myCSV = pd.read_csv(file, sep=';')
df = pd.DataFrame(index=range(0, len(myCSV)), columns=headers)
contexts = []

for row in myCSV.values:
    
    if row[5] == "M":
        create = "разработал"
    else:
        create = "разработала"
    contexts.append({'author': row[1]+" "+row[2]+" "+row[3], 'course': row[0], 'grade': row[4], 'create': create, 'email': row[6]})

for context in contexts:
    print(context)
    doc.render(context)
    doc.save(r'D:\Dev\TeacherCourse\certificates\\'+context["author"]+".docx")
'''
wdFormatPDF = 17
in_folder = r'D:\Dev\TeacherCourse\Certificates\\'
out_folder = r'D:\Dev\TeacherCourse\PdfCertificates\\'

for in_file_name in os.listdir(in_folder):
    print(in_file_name)
    in_file = in_folder + in_file_name
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    print("\n" + in_file + " opened")

    outfile_name = in_file_name.replace("docx", "pdf")
    out_file = out_folder + outfile_name
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
    print("successfully converted " + outfile_name)
'''