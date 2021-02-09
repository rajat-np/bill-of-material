def save_bom(item, raw_materials, workbook):
    worksheet = workbook.create_sheet(item)
    worksheet.cell(row=1, column=1, value='Finished Good List')
    worksheet.cell(row=2, column=1, value='#')
    worksheet.cell(row=2, column=2, value='Item Description')
    worksheet.cell(row=2, column=3, value='Quantity')
    worksheet.cell(row=2, column=4, value='Unit')
    worksheet.cell(row=3, column=1, value=1)
    worksheet.cell(row=3, column=2, value=item)
    worksheet.cell(row=3, column=3, value=1)
    worksheet.cell(row=3, column=4, value='Pc')
    worksheet.cell(row=4, column=1, value='End of FG')
    worksheet.cell(row=5, column=1, value='Raw Material List')
    worksheet.cell(row=6, column=1, value='#')
    worksheet.cell(row=6, column=2, value='Item Description')
    worksheet.cell(row=6, column=3, value='Quantity')
    worksheet.cell(row=6, column=4, value='Unit')

    for i in range(len(raw_materials.index)):
        worksheet.cell(row=i+7, column=1, value=i+1)
        worksheet.cell(row=i+7, column=2,
                       value=raw_materials.iloc[i]['Raw material'])
        worksheet.cell(row=i+7, column=3,
                       value=raw_materials.iloc[i]['Quantity'])
        worksheet.cell(row=i+7, column=4,
                       value=raw_materials.iloc[i]['Unit '])

    worksheet.cell(row=i+8, column=1, value='End of RM')
