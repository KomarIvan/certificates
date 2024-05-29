import csv
from docxtpl import DocxTemplate
from docx2pdf import convert
import os

csv_file_path = r"C:\...\1.csv"
output_folder = r"C:\..."  # Папка для сохранения файлов

surnames, names, patronymics = [], [], []

with open(csv_file_path, "r") as csvf:
    op = csvf.readlines()

    for i in op[1:]:
        surname = i.split(";")[0].strip()  
        name = i.split(";")[1].strip() 
        patronymic = i.split(";")[2].strip()  
        
        # Добавьте данные в соответствующие списки
        surnames.append(surname)
        names.append(name)
        patronymics.append(patronymic)


# Создайте новый список, чтобы хранить результаты обработки шаблона
result_files = []


for i in range(len(surnames)):
    doc = DocxTemplate(r"C:\....docx")
    context = {
        "surname": surnames[i],
        "name": names[i],
        "patronymic": patronymics[i],
    }

    doc.render(context)

    file_name = f"{surnames[i]}_{names[i]}_{patronymics[i]}"
    docx_file = os.path.join(output_folder, f"{file_name}.docx")
    pdf_file = os.path.join(output_folder, f"{file_name}.pdf")

    # Сохраняем в docx-файл
    doc.save(docx_file) 
    
    # Конвертируем в PDF
    convert(docx_file, pdf_file)
    
    # Добавляем имя файла в список результатов
    result_files.append(pdf_file)

# Выводим список созданных файлов
print("Созданные файлы:", result_files)
