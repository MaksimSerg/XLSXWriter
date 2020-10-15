import XLSXWriter

styles = {
    'header': {
        'format': '@',                      # '@'
        'font': 'Arial',                    # font name: string
        'font-size': 10,                    # font size: integer
        'font-style': 'bold, italic',       # font style: string: 'bold', 'italic'
        'color': '#409EFF',                 # string: font color
        'fill': '#b3d8ff',                  # string: background color
        'halign': 'center',                 # string: 'left', 'right', 'justify', 'center'
        'valign': 'center',                 # string: 'bottom', 'center', 'distributed', 'top'
        'border': 'left,right,top,bottom',  # string: 'left', 'right', 'top', 'bottom'
        'border-color': '#8cb8e6',          # string: border color
        'border-style': 'thin',             # string: 'thin', 'medium', 'thick', 'dashDot', 'dashDotDot', 'dashed', 'dotted', 'double', 'hair', 'mediumDashDot', 'mediumDashDotDot', 'mediumDashed', 'slantDashDot'
        'wrap_text': True                   # boolean: True, False
    },
    'body': {
        'font-size': 10,
        'color': '#5e6169',
        'wrap_text': True,
        'border': 'left,right,top,bottom',
        'border-color': '#5e6169',
        'border-style': 'thin',
    },
    'body2': [
        {
            'format': 'date'
        },
        {
            'format': 'datetime'
        },
        {
            'format': '# ###,##0.00'
        },
        {
            'format': 'integer'
        },
        {
            'format': '#,##0'
        },
        {
            'format': 'price'
        },
        {
            'format': 'string'
        },
        {
            'format': '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00'
        },
    ]
}

row_options = {
    'height': 12.1,
    'hidden': False,
    'collapsed': 0      # level
}

writer = XLSXWriter.Writer()
writer.setStyles(styles)
writer.sheetAdd('Sheet1', col_widths=(15, 25, 20))


writer.writeSheetRow(
    ['Column1', 'Column2', 'Column3', 'Column4', 'Column5', 'Column6', 'Column7', 'Column8'],
    styles='header', row_options={'height': 35}
)

writer.writeSheetRow(
    ['2020-01-01', '2020-11-01 14:14:34', 122345.36365, 873, 1, '44.00', 'text text text text text ', '10'],
    styles='body'
)

writer.writeSheetRow(
    ['2020-01-01', '2020-11-01 14:14:34', 122345.36365, 873, 1, '44.00', 'text text text text text', '=E2*(-0.05)'],
    styles='body2'
)

writer.writeSheetRow([])

writer.writeSheetRow(['Group 1'])
for r in range(5):
    writer.writeSheetRow(
        [None, 'hidden 1'], row_options = { 'height': 12.1, 'hidden': False, 'collapsed': 1 }
    )

writer.writeSheetRow(['Group 2'])
for r in range(5):
    writer.writeSheetRow(
        [None, 'hidden 2'], row_options = { 'height': 12.1, 'hidden': False, 'collapsed': 1 }
    )

writer.writeSheetRow(['Group 3'])
for r31 in range(3):
    writer.writeSheetRow(
        [None, 'hidden 31'], row_options = { 'height': 12.1, 'hidden': True, 'collapsed': 1 }
    )
    for r32 in range(2):
        writer.writeSheetRow(
            [None, None, 'hidden 32'], row_options={'height': 12.1, 'hidden': True, 'collapsed': 2}
        )
        for r33 in range(2):
            writer.writeSheetRow(
                [None, None, None, 'hidden 33'], row_options={'height': 12.1, 'hidden': True, 'collapsed': 3}
            )

# set autofilter
writer.sheetSetFiltr((0, 0), (3, 7))  # RC

writer.saveAs("test.xlsx")