from docxtpl import DocxTemplate
import pandas as pd
# Load the xlsx file
excel_data = pd.read_excel('data.xlsx')
# Read the values of the file in the dataframe
data = pd.DataFrame(excel_data, columns=['name', 'date', 'male', 'enter'])



for i in range(len(data)):
    if data['male'][i] == 'м':
        male = 'он'
        maleEnd = ''
    else:
        male = 'она'
        maleEnd = 'а'

    doc = DocxTemplate("шаблон.docx")


    context = {'name': data['name'][i],
               'date': data['date'][i],
               'male': male,
               'maleEnd': maleEnd,
               'enter': data['enter'][i]}
    doc.render(context)
    doc.save(f"res/Справка {context['name']}.docx")



