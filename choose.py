import openpyxl
import json

with open('color.json', 'r') as file:
    client_data = json.load(file)
    
row_need = ['green']

filter_data = []

for item in client_data:
    if item['color'] in row_need:
        filter_data.append({
            'fruit' : item['fruit'],
            'size'  : item['size'],
            'color' : item['color'],
        
        })
        
        
client_book = openpyxl.Workbook()

client_sheet = client_book.active



main= ['fruit', 'size', 'color']

client_sheet.append(main)


for item in filter_data:
    client_sheet.append([
        item['fruit'],
        item['size'],
        item['color'],
       
    ])

# Сохраняем в Excel файл
client_book.save('client_fruits.xlsx')

print("Успешно сохранено!")
         
