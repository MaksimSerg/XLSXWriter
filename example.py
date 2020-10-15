import XLSXWriter

styles = {
    'header': {
        'format': '@',
        'font': 'Arial',  #
        'font-size': 10.0,  #
        'font-style': 'bold, italic',  #
        'color': '#ff5722',  #
        'fill': '#125678',  #
        'halign': 'center',  #
        'valign': 'center',  #
        'border': 'left,right,top,bottom',  #
        'border-color': '#4caf50',  #
        'border-style': 'thin',  #
        'wrap_text': True
    },
    'body': {
        'font-style': 'italic'
    },
    'body2': [
        {
            'format': 'date',
            'font': 'Arial',  #
            'font-size': 8.5
        },
        {
            'format': 'datetime',
            'font': 'Arial',  #
            'font-size': 8.5
        },
        {
            'format': '# ###,##0.00',
            'font': 'Arial',  #
            'font-size': 8.5
        },
        {
            'format': 'integer',
            'font': 'Arial',  #
            'font-size': 8.5
        },
        {
            'format': '#,##0',
            'font': 'Arial',  #
            'font-size': 8.5
        },
        {
            'format': 'price',
            'font': 'Arial',  #
            'font-size': 8.5,
            'border': 'top'
        },
        {
            'format': 'string',
            'font': 'Arial',  #
            'font-size': 8.5,
            'wrap_text': True
        },
        {
            'format': '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',
            'font': 'Arial',  #
            'font-size': 8.5,
            'border': 'bottom'
        },
    ]
}

writer = XLSXWriter.Writer()
writer.setStyles(styles)
writer.sheetAdd('Лист1', col_widths=(15, 25))

writer.writeSheetRow(
    ['2015-01-01', '2015-11-01 14:14:34', 122345.36365, 873, 1, '44.00', 'misc misc misc misc misc', '=E2*0.05'],
    styles='header', row_options={'height': 25}
)
writer.writeSheetRow(
    ['2015-01-01', '2015-11-01 14:14:34', 122345.36365, 873, 1, '44.00', 'misc misc misc misc misc', '=E2*0.05'],
    styles='body', row_options={'height': 25, 'collapsed': 1, 'hidden': True}
)
writer.writeSheetRow(
    ['2015-01-01', '2015-11-01 14:14:34', 122345.36365, 873, 1, '44.00', 'misc misc misc misc misc', '=E2*0.05'],
    styles='body2', row_options={'height': 25, 'collapsed': 1, 'hidden': True}
)

writer.sheetSetFiltr((0, 0), (3, 7))  # RC

writer.saveAs("test.xlsx")

print("--end--")
