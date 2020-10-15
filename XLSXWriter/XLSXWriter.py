import os
import time
import random
import re
import tempfile
import zipfile
import json
import copy
from functools import reduce
from collections import OrderedDict


def html_special_chars(text):
    return text \
        .replace("&", "&amp;") \
        .replace('"', "&quot;") \
        .replace("'", "&#039;") \
        .replace("<", "&lt;") \
        .replace(">", "&gt;") \
        .replace("\n", "&#10;")


class BuffererWriter:

    def __init__(self, filename, fd_fopen_flags='w', buffer_size=400):
        self.fd = open(filename, fd_fopen_flags)
        self.buffer_size = buffer_size
        self.buffer = ""
        if self.fd == False:
            raise Exception("Unable to open $filename for writing.")

    def __del__(self):
        self.close()

    def write(self, text):
        self.buffer += text
        if len(self.buffer) > self.buffer_size:
            self.purge()

    def purge(self):
        if self.fd:
            self.fd.write(self.buffer)
            self.buffer = ""

    def close(self):
        self.purge()
        if self.fd:
            self.fd.close()
            self.fd = None

    def ftell(self):
        if not self.fd:
            self.purge()
            return self.fd.tell()

        return 0

    def fseek(self, pos):
        if self.fd:
            self.purge()
            return self.fd.seek(pos)

        return -1


class Writer:
    EXCEL_2007_MAX_ROW = 1048570    # todo - не совсем так
    EXCEL_2007_MAX_COL = 256        # todo - не совсем так

    def __init__(self, buffer_size=1024):
        self._buffer_size = buffer_size
        self._title = ""
        self._subject = ""
        self._author = ""
        self._company = ""
        self._description = ""
        self._keywords = []
        self._current_sheet = ""
        self._sheets = OrderedDict({})
        self._tempdir = ""
        self._temp_files = []
        self._cell_styles = []
        self._number_formats = []
        self._styles = {}

        self.__addCellStyle(number_format='GENERAL', cell_style_string='')

    def __del__(self):
        for f in self._temp_files:
            if os.path.exists(f):
                os.unlink(f)

    def setTitle(self, title=''):
        self._title = title

    def setSubject(self, subject=''):
        self._subject = subject

    def setAuthor(self, author=''):
        self._author = author

    def setCompany(self, company=''):
        self._company = company

    def setKeywords(self, keywords=''):
        self._keywords = keywords

    def setDescription(self, description=''):
        self._description = description

    def setTempDir(self, tempdir=''):
        self._tempdir = tempdir

    def _tempFilename(self):
        tempdir = self._tempdir if self._tempdir else tempfile.gettempdir()
        filename = tempdir + "/xlsx_writer_" + str(random.randint(1, 1000))
        self._temp_files.append(filename)

        return filename

    def writeToStdOut(self):
        temp_file = self._tempFilename()
        self.writeToFile(temp_file)
        open(temp_file, 'r').read()

    def writeToString(self):
        temp_file = self._tempFilename()
        self.writeToFile(temp_file)
        return open(temp_file).read()

    def saveAs(self, filename):
        self.writeSheetRow([])
        self.writeToFile(filename)

    def sheetAdd(self, sheet_name, col_widths=(), freeze_rows=False, freeze_columns=False):
        self._initializeSheet(sheet_name, col_widths, freeze_rows, freeze_columns)
        self._current_sheet = sheet_name

    def setActiveSheet(self, sheet_name):
        self.sheetAdd(sheet_name)

    def writeToFile(self, filename):
        for sheet in self._sheets:
            self._finalizeSheet(sheet)

        if os.path.exists(filename):
            if os.access(filename, os.W_OK):
                os.unlink(filename)
            else:
                raise Exception("Error: " + "file is not writeable.")

        if len(self._sheets) == 0:
            raise Exception("Error: " + " no worksheets defined.")

        zip = zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED)
        zip.writestr("docProps/app.xml", self._buildAppXML())
        zip.writestr("docProps/core.xml", self._buildCoreXML())
        zip.writestr("_rels/.rels", self._buildRelationshipsXML())
        for sheet in self._sheets:
            zip.write(self._sheets[sheet]['filename'], "xl/worksheets/" + self._sheets[sheet]['xmlname'])

        zip.writestr("xl/workbook.xml", self._buildWorkbookXML())
        zip.write(self._writeStylesXML(), "xl/styles.xml")
        zip.writestr("[Content_Types].xml", self._buildContentTypesXML())
        zip.writestr("xl/_rels/workbook.xml.rels", self._buildWorkbookRelsXML())
        zip.close()

    def _initializeSheet(self, sheet_name, col_widths=(), freeze_rows=False, freeze_columns=False):

        sheet_filename = self._tempFilename()
        sheet_xmlname = 'sheet' + str(len(self._sheets) + 1) + ".xml"
        self._sheets[sheet_name] = {
            'filename': sheet_filename,
            'sheetname': sheet_name,
            'xmlname': sheet_xmlname,
            'row_count': 0,
            'file_writer': BuffererWriter(sheet_filename, buffer_size=self._buffer_size),
            'columns': [],
            'merge_cells': [],
            'max_cell_tag_start': 0,
            'max_cell_tag_end': 0,
            'auto_filter': None,
            'freeze_rows': freeze_rows,
            'freeze_columns': freeze_columns,
            'finalized': False
        }

        sheet = self._sheets[sheet_name]
        tabselected = 'true' if len(self._sheets) == 1 else 'false'
        max_cell = self.xlsCell(self.EXCEL_2007_MAX_ROW, self.EXCEL_2007_MAX_COL)
        sheet['file_writer'].write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "\n")
        sheet['file_writer'].write(
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        sheet['file_writer'].write('<sheetPr filterMode="false">')
        sheet['file_writer'].write('<pageSetUpPr fitToPage="false"/>')
        sheet['file_writer'].write('</sheetPr>')
        sheet['max_cell_tag_start'] = sheet['file_writer'].ftell()
        sheet['file_writer'].write('<dimension ref="A1:' + str(max_cell) + '"/>')
        sheet['max_cell_tag_end'] = sheet['file_writer'].ftell()
        sheet['file_writer'].write('<sheetViews>')
        sheet['file_writer'].write(
            '<sheetView colorId="64" defaultGridColor="true" rightToLeft="false" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" tabSelected="' + tabselected + '" topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">')

        if sheet['freeze_rows'] and sheet['freeze_columns']:
            sheet['file_writer'].write(
                '<pane ySplit="' + str(sheet['freeze_rows']) + '" xSplit="' + str(sheet['freeze_columns']) + '" topLeftCell="' + self.xlsCell(sheet['freeze_rows'], sheet[
                    'freeze_columns']) + '" activePane="bottomRight" state="frozen"/>')
            sheet['file_writer'].write(
                '<selection activeCell="' + self.xlsCell(sheet['freeze_rows'], 0) + '" activeCellId="0" pane="topRight" sqref="' + self.xlsCell(sheet['freeze_rows'], 0) + '"/>')
            sheet['file_writer'].write('<selection activeCell="' + self.xlsCell(0, sheet['freeze_columns']) + '" activeCellId="0" pane="bottomLeft" sqref="' + self.xlsCell(0,
                                                                                                                                                                            sheet[
                                                                                                                                                                                'freeze_columns']) + '"/>')
            sheet['file_writer'].write(
                '<selection activeCell="' + self.xlsCell(sheet['freeze_rows'], sheet['freeze_columns']) + '" activeCellId="0" pane="bottomRight" sqref="' + self.xlsCell(
                    sheet['freeze_rows'], sheet['freeze_columns']) + '"/>')
        elif sheet['freeze_rows']:
            sheet['file_writer'].write(
                '<pane ySplit="' + str(sheet['freeze_rows']) + '" topLeftCell="' + self.xlsCell(sheet['freeze_rows'], 0) + '" activePane="bottomLeft" state="frozen"/>')
            sheet['file_writer'].write(
                '<selection activeCell="' + self.xlsCell(sheet['freeze_rows'], 0) + '" activeCellId="0" pane="bottomLeft" sqref="' + self.xlsCell(sheet['freeze_rows'], 0) + '"/>')
        elif sheet['freeze_columns']:
            sheet['file_writer'].write(
                '<pane xSplit="' + str(sheet['freeze_columns']) + '" topLeftCell="' + self.xlsCell(0, sheet['freeze_columns']) + '" activePane="topRight" state="frozen"/>')
            sheet['file_writer'].write('<selection activeCell="' + self.xlsCell(0, sheet['freeze_columns']) + '" activeCellId="0" pane="topRight" sqref="' + self.xlsCell(0, sheet[
                'freeze_columns']) + '"/>')
        else:
            sheet['file_writer'].write('<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>')

        sheet['file_writer'].write('</sheetView>')
        sheet['file_writer'].write('</sheetViews>')
        sheet['file_writer'].write('<cols>')

        i = 0
        if len(col_widths) > 0:
            for column_width in col_widths:
                sheet['file_writer'].write(
                    '<col collapsed="false" hidden="false" max="' + str(i + 1) + '" min="' + str(i + 1) + '" style="0" customWidth="true" width="' + str(column_width) + '"/>')
                i += 1

        sheet['file_writer'].write('<col collapsed="false" hidden="false" max="1024" min="' + str(i + 1) + '" style="0" customWidth="false" width="11.5"/>')
        sheet['file_writer'].write('</cols>')
        sheet['file_writer'].write('<sheetData>')

    def __addCellStyle(self, number_format, cell_style_string):
        cell_style_string = cell_style_string if cell_style_string else ""
        number_format_idx = self.add_to_list_get_index(self._number_formats, number_format)
        lookup_string = str(number_format_idx) + ";" + cell_style_string
        cell_style_idx = self.add_to_list_get_index(self._cell_styles, lookup_string)

        return cell_style_idx

    def __initializeColumnTypes(self, header_types):
        column_types = []
        for v in header_types:
            cell_style_string = {}

            cellFormat = v.get('format', '@')
            if 'wrap' in v:
                if v['wrap']:
                    cell_style_string['wrap_text'] = True

            number_format = self.__numberFormatStandardized(cellFormat)
            number_format_type = self.__determineNumberFormatType(number_format)
            cell_style_idx = self.__addCellStyle(number_format, cell_style_string=json.dumps(cell_style_string))

            column_types.append({
                'number_format': number_format,
                'number_format_type': number_format_type,
                'default_cell_style': cell_style_idx
            })

        return column_types

    def sheetSetFiltr(self, cell1, cell2):
        if not self._current_sheet:
            return 0

        sheet = self._sheets[self._current_sheet]
        sheet['auto_filter'] = self.xlsCell(cell1[0], cell1[1], True) + ':' + self.xlsCell(cell2[0], cell2[1], True)

    def writeSheetHeader(self, header_types, col_options={}):
        if not self._current_sheet or len(header_types) == 0:
            return 0

        sheet_name = self._current_sheet

        suppress_row = int(col_options['suppress_row']) if 'suppress_row' in col_options else False
        if isinstance(col_options, bool):
            suppress_row = int(col_options)

        style = col_options

        col_widths = []
        for column in header_types:
            col_widths.append(column.get('width', 12))

        auto_filter = col_options['auto_filter'] if 'auto_filter' in col_options else False
        freeze_rows = col_options['freeze_rows'] if 'freeze_rows' in col_options else False
        freeze_columns = col_options['freeze_columns'] if 'freeze_columns' in col_options else False
        self._initializeSheet(sheet_name, col_widths, freeze_rows, freeze_columns)
        sheet = self._sheets[sheet_name]
        sheet['columns'] = self.__initializeColumnTypes(header_types)

        if not suppress_row:
            sheet['file_writer'].write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="1">')
            for c, v in enumerate(header_types):
                cell_style_idx = sheet['columns'][c]['default_cell_style'] if style is None else self.__addCellStyle('GENERAL', json.dumps(style[c] if '0' in style else style))
                self._writeCell(sheet['file_writer'], 0, c, v, 'n_string', cell_style_idx)
            sheet['file_writer'].write('</row>')
            sheet['row_count'] += 1

    def writeSheetRow(self, row, styles=None, row_options=None):
        if self._current_sheet == "":
            return 0

        sheet = self._sheets[self._current_sheet]

        if row_options is not None:
            ht = float(row_options['height']) if 'height' in row_options else 12.1
            customHt = 'true' if 'height' in row_options else 'false'
            hidden = 'true' if row_options.get('hidden', False) else 'false'
            collapsed = int(row_options.get('collapsed', 0))
            sheet['file_writer'].write(
                '<row collapsed="' + ('true' if collapsed > 0 else 'false') + '" customFormat="0" customHeight="' + customHt + '" hidden="' + hidden + '" ht="' + str(
                    ht) + '" outlineLevel="' + str(collapsed) + '" r="' + str(sheet['row_count'] + 1) + '">')
        else:
            sheet['file_writer'].write(
                '<row collapsed="false" customFormat="0" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' + str(sheet['row_count'] + 1) + '">')

        cell_count = len(row)
        cell_styles = [(0, 'n_auto')] * cell_count
        if styles:
            row_styles = self._styles.get(styles, (0, 'n_auto'))
            if isinstance(row_styles, tuple):
                cell_styles = [row_styles] * cell_count
            elif isinstance(row_styles, list):
                cell_styles = (row_styles + cell_styles)[:cell_count]

        for c, v in enumerate(row):
            self._writeCell(sheet['file_writer'], sheet['row_count'], c, v, cell_styles[c][1], cell_styles[c][0])

        sheet['file_writer'].write('</row>')
        sheet['row_count'] += 1

    def countSheetRows(self, sheet_name=''):
        sheet_name = sheet_name if sheet_name else self._current_sheet

        return self._sheets[sheet_name]['row_count'] if sheet_name in self._sheets else 0

    def _finalizeSheet(self, sheet_name):
        if not sheet_name or self._sheets[sheet_name]['finalized']:
            return 0

        sheet = self._sheets[sheet_name]
        sheet['file_writer'].write('</sheetData>')
        if len(sheet['merge_cells']) > 0:
            sheet['file_writer'].write('<mergeCells count="' + str(len(sheet['merge_cells'])) + '">')
            for _range in sheet['merge_cells']:
                sheet['file_writer'].write('<mergeCell ref="' + _range + '"/>')

            sheet['file_writer'].write('</mergeCells>')

        max_cell = self.xlsCell(sheet['row_count'] - 1, len(sheet['columns']) - 1)

        if sheet['auto_filter']:
            sheet['file_writer'].write('<autoFilter ref="' + sheet['auto_filter'] + '"></autoFilter>')

        sheet['file_writer'].write('<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>')
        sheet['file_writer'].write('<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>')
        sheet['file_writer'].write(
            '<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>')
        sheet['file_writer'].write('<headerFooter differentFirst="false" differentOddEven="false">')
        sheet['file_writer'].write('<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>')
        sheet['file_writer'].write('<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>')
        sheet['file_writer'].write('</headerFooter>')
        sheet['file_writer'].write('</worksheet>')

        # max_cell_tag = '<dimension ref="A1:' + str(max_cell) + '"/>'
        # padding_length = sheet['max_cell_tag_end'] - sheet['max_cell_tag_start'] - len(max_cell_tag)
        #
        # sheet['file_writer'].fseek(sheet['max_cell_tag_start'])
        # sheet['file_writer'].write(max_cell_tag * padding_length)

        sheet['file_writer'].close()

        sheet['finalized'] = True

    def markMergedCell(self, sheet_name, cell1, cell2):
        if not sheet_name or self._sheets[sheet_name]['finalized']:
            return 0

        sheet = self._sheets[sheet_name]
        startCell = self.xlsCell(cell1[0], cell1[1])
        endCell = self.xlsCell(cell2[0], cell2[1])
        sheet['merge_cells'].append(str(startCell) + ":" + str(endCell))

    def writeSheet(self, data, sheet_name='', header_types=[]):
        sheet_name = sheet_name if sheet_name else 'Sheet1'
        data = data if data else [['']]
        if header_types:
            self.writeSheetHeader(sheet_name, header_types)

        for row in data:
            self.writeSheetRow(sheet_name, row)

        self._finalizeSheet(sheet_name)

    def _writeCell(self, file, row_number, column_number, value, num_format_type, cell_style_idx):
        cell_name = self.xlsCell(row_number, column_number)

        if value == '' or value is None:
            file.write('<c r="' + cell_name + '" s="' + str(cell_style_idx) + '"/>')
        elif num_format_type == 'n_auto':
            if type(value) in (int, float):
                file.write('<c r="' + cell_name + '" s="' + str(cell_style_idx) + '" t="n"><v>' + self.xmlspecialchars(value) + '</v></c>')
            else:
                file.write('<c r="' + cell_name + '" s="' + str(cell_style_idx) + '" t="inlineStr"><is><t>' + self.xmlspecialchars(value) + '</t></is></c>')
        elif type(value) == str and value[:1] == '=':
            file.write('<c r="' + cell_name + '" s="' + str(cell_style_idx) + '" t="s"><f>' + self.xmlspecialchars(value) + '</f></c>')
        elif num_format_type == 'n_date':
            file.write('<c r="' + cell_name + '" s="' + str(cell_style_idx) + '" t="n"><v>' + str(self.convert_date_time(value)) + '</v></c>')
        elif num_format_type == 'n_datetime':
            file.write('<c r="' + cell_name + '" s="' + str(cell_style_idx) + '" t="n"><v>' + str(self.convert_date_time(value)) + '</v></c>')
        elif num_format_type == 'n_numeric':
            file.write('<c r="' + cell_name + '" s="' + str(cell_style_idx) + '" t="n"><v>' + self.xmlspecialchars(value) + '</v></c>')
        elif num_format_type == 'n_string':
            file.write('<c r="' + cell_name + '" s="' + str(cell_style_idx) + '" t="inlineStr"><is><t>' + self.xmlspecialchars(value) + '</t></is></c>')

    def _styleFontIndexes(self):
        border_allowed = ['left', 'right', 'top', 'bottom']
        border_style_allowed = ['thin', 'medium', 'thick', 'dashDot', 'dashDotDot', 'dashed', 'dotted', 'double', 'hair', 'mediumDashDot', 'mediumDashDotDot', 'mediumDashed',
                                'slantDashDot']
        horizontal_allowed = ['general', 'left', 'right', 'justify', 'center']
        vertical_allowed = ['bottom', 'center', 'distributed', 'top']
        default_font = {'size': '10', 'name': 'Arial', 'family': '2'}
        fills = ['', '']
        fonts = ['', '', '', '']
        borders = ['']
        style_indexes = [None] * len(self._cell_styles)

        for i, cell_style_string in enumerate(self._cell_styles):
            semi_colon_pos = cell_style_string.find(";")
            number_format_idx = int(cell_style_string[0:semi_colon_pos])
            style_json_string = cell_style_string[semi_colon_pos + 1:]
            style = json.loads(style_json_string) if style_json_string else {}
            style_indexes[i] = {'num_fmt_idx': number_format_idx}

            if 'border' in style and isinstance(style.get('border', 0), str):
                border_value = {}
                border_value['side'] = list(reduce(set.intersection, map(set, [style['border'].split(","), border_allowed])))
                if 'border-color' in style:
                    if style['border-style'] in border_style_allowed:
                        border_value['style'] = style['border-style']
                    if isinstance(style['border-color'], str):
                        if style['border-color'][:1] == '#':
                            v = style['border-color'][1:7]
                            v = v[0] + v[0] + v[1] + v[1] + v[2] + v[2] if len(v) == 3 else v
                            border_value['color'] = "FF" + v.upper()

                style_indexes[i]['border_idx'] = self.add_to_list_get_index(borders, json.dumps(border_value))

            if 'fill' in style:
                if isinstance(style.get('fill', 0), str):
                    if style['fill'][:1] == '#':
                        v = style['fill'][1:7]
                        v = v[0] + v[0] + v[1] + v[1] + v[2] + v[2] if len(v) == 3 else v
                        style_indexes[i]['fill_idx'] = self.add_to_list_get_index(fills, "FF" + v.upper())

            if 'halign' in style:
                if style['halign'] in horizontal_allowed:
                    style_indexes[i]['alignment'] = True
                    style_indexes[i]['halign'] = style['halign']

            if 'valign' in style:
                if style['valign'] in vertical_allowed:
                    style_indexes[i]['alignment'] = True
                    style_indexes[i]['valign'] = style['valign']

            if 'wrap_text' in style:
                style_indexes[i]['alignment'] = True
                style_indexes[i]['wrap_text'] = bool(style['wrap_text'])

            font = default_font.copy()
            is_add_font_idx = False
            if 'font-size' in style:
                font['size'] = float(style['font-size'])
                is_add_font_idx = True

            if 'font' in style:
                if isinstance(style.get('font', 0), str):
                    if style['font'] == 'Comic Sans MS': font['family'] = 4
                    if style['font'] == 'Times New Roman': font['family'] = 1
                    if style['font'] == 'Courier New': font['family'] = 3
                    font['name'] = str(style['font'])
                    is_add_font_idx = True

            if 'font-style' in style:
                if isinstance(style.get('font-style', 0), str):
                    if style['font-style'].find('bold') > -1: font['bold'] = True
                    if style['font-style'].find('italic') > -1: font['italic'] = True
                    if style['font-style'].find('strike') > -1: font['strike'] = True
                    if style['font-style'].find('underline') > -1: font['underline'] = True
                    is_add_font_idx = True

            if 'color' in style:
                if isinstance(style.get('color', 0), str):
                    if style['color'][:1] == '#':
                        v = style['color'][1:7]
                        v = v[0] + v[0] + v[1] + v[1] + v[2] + v[2] if len(v) == 3 else v
                        font['color'] = "FF" + v.upper()
                        is_add_font_idx = True

            if is_add_font_idx:
                style_indexes[i]['font_idx'] = self.add_to_list_get_index(fonts, json.dumps(font))

        return {'fills': fills, 'fonts': fonts, 'borders': borders, 'styles': style_indexes}

    def _writeStylesXML(self):
        r = self._styleFontIndexes()

        fills = r['fills']
        fonts = r['fonts']
        borders = r['borders']
        style_indexes = r['styles']
        temporary_filename = self._tempFilename()
        file = BuffererWriter(temporary_filename)

        file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "\n")
        file.write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        file.write('<numFmts count="' + str(len(self._number_formats)) + '">')

        for i, v in enumerate(self._number_formats):
            file.write('<numFmt numFmtId="' + str(164 + i) + '" formatCode="' + self.xmlspecialchars(v) + '" />')

        file.write('</numFmts>')
        file.write('<fonts count="' + str(len(fonts)) + '">')
        file.write('<font><name val="Arial"/><charset val="1"/><family val="2"/><sz val="10"/></font>')
        file.write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>')
        file.write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>')
        file.write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>')

        for font in fonts:
            if font:
                f = json.loads(font)
                file.write('<font>')
                file.write('<name val="' + html_special_chars(f['name']) + '"/><charset val="1"/><family val="' + str(int(f['family'])) + '"/>')
                file.write('<sz val="' + str(int(f['size'])) + '"/>')

                if f.get('color', False): file.write('<color rgb="' + str(f['color']) + '"/>')
                if f.get('bold', False): file.write('<b val="true"/>')
                if f.get('italic', False): file.write('<i val="true"/>')
                if f.get('underline', False): file.write('<u val="single"/>')
                if f.get('strike', False): file.write('<strike val="true"/>')

                file.write('</font>')

        file.write('</fonts>')
        file.write('<fills count="' + str(len(fills)) + '">')
        file.write('<fill><patternFill patternType="none"/></fill>')
        file.write('<fill><patternFill patternType="gray125"/></fill>')

        for fill in fills:
            if fill:
                file.write('<fill><patternFill patternType="solid"><fgColor rgb="' + str(fill) + '"/><bgColor indexed="64"/></patternFill></fill>')

        file.write('</fills>')
        file.write('<borders count="' + str(len(borders)) + '">')
        file.write('<border diagonalDown="false" diagonalUp="false"><left/><right/><top/><bottom/><diagonal/></border>')

        for border in borders:
            if bool(border):
                pieces = json.loads(border)
                border_style = pieces['style'] if pieces.get('style', False) else 'hair'
                border_color = '<color rgb="' + str(pieces['color']) + '"/>' if pieces.get('color', False) else ''
                file.write('<border diagonalDown="false" diagonalUp="false">')

                for side in ('left', 'right', 'top', 'bottom'):
                    show_side = side in pieces['side']
                    file.write('<' + side + ' style="' + border_style + '">' + border_color + '</' + side + '>' if show_side else "<" + side + "/>")

                file.write('<diagonal/>')
                file.write('</border>')

        file.write('</borders>')
        file.write('<cellStyleXfs count="20">')
        file.write('<xf applyAlignment="true" applyBorder="true" applyFont="true" applyProtection="true" borderId="0" fillId="0" fontId="0" numFmtId="164">')
        file.write('<alignment horizontal="general" indent="0" shrinkToFit="false" textRotation="0" vertical="bottom" wrapText="false"/>')
        file.write('<protection hidden="false" locked="true"/>')
        file.write('</xf>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="43"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="41"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="44"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="42"/>')
        file.write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="9"/>')
        file.write('</cellStyleXfs>')
        file.write('<cellXfs count="' + str(len(style_indexes)) + '">')

        for v in style_indexes:
            applyAlignment = 'true' if 'alignment' in v else 'false'
            wrapText = 'true' if 'wrap_text' in v else 'false'
            horizAlignment = v['halign'] if 'halign' in v else 'general'
            vertAlignment = v['valign'] if 'valign' in v else 'bottom'
            applyBorder = 'true' if 'border_idx' in v else 'false'
            applyFont = 'true'
            borderIdx = int(v['border_idx']) if 'border_idx' in v else 0
            fillIdx = int(v['fill_idx']) if 'fill_idx' in v else 0
            fontIdx = int(v['font_idx']) if 'font_idx' in v else 0

            file.write('<xf applyAlignment="' + applyAlignment + '" applyBorder="' + applyBorder + '" applyFont="' + applyFont + '" applyProtection="false" borderId="' + str(
                borderIdx) + '" fillId="' + str(fillIdx) + '" fontId="' + str(fontIdx) + '" numFmtId="' + str(164 + v['num_fmt_idx']) + '" xfId="0">')
            file.write(
                '    <alignment horizontal="' + horizAlignment + '" vertical="' + vertAlignment + '" textRotation="0" wrapText="' + wrapText + '" indent="0" shrinkToFit="false"/>')
            file.write('    <protection locked="true" hidden="false"/>')
            file.write('</xf>')

        file.write('</cellXfs>')
        file.write('<cellStyles count="6">')
        file.write('<cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>')
        file.write('<cellStyle builtinId="3" customBuiltin="false" name="Comma" xfId="15"/>')
        file.write('<cellStyle builtinId="6" customBuiltin="false" name="Comma [0]" xfId="16"/>')
        file.write('<cellStyle builtinId="4" customBuiltin="false" name="Currency" xfId="17"/>')
        file.write('<cellStyle builtinId="7" customBuiltin="false" name="Currency [0]" xfId="18"/>')
        file.write('<cellStyle builtinId="5" customBuiltin="false" name="Percent" xfId="19"/>')
        file.write('</cellStyles>')
        file.write('</styleSheet>')
        file.close()

        return temporary_filename

    def _buildAppXML(self):
        app_xml = ""
        app_xml += '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "\n"
        app_xml += '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        app_xml += '<TotalTime>0</TotalTime>'
        app_xml += '<Company>' + self.xmlspecialchars(self._company) + '</Company>'
        app_xml += '</Properties>'

        return app_xml

    def _buildCoreXML(self):
        core_xml = ""
        core_xml += '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "\n"
        core_xml += '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        core_xml += '<dcterms:created xsi:type="dcterms:W3CDTF">' + time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()) + '</dcterms:created>'
        core_xml += '<dc:title>' + self.xmlspecialchars(self._title) + '</dc:title>'
        core_xml += '<dc:subject>' + self.xmlspecialchars(self._subject) + '</dc:subject>'
        core_xml += '<dc:creator>' + self.xmlspecialchars(self._author) + '</dc:creator>'

        if self._keywords is not None:
            core_xml += '<cp:keywords>' + self.xmlspecialchars(", ".join(self._keywords)) + '</cp:keywords>'

        core_xml += '<dc:description>' + self.xmlspecialchars(self._description) + '</dc:description>'
        core_xml += '<cp:revision>0</cp:revision>'
        core_xml += '</cp:coreProperties>'

        return core_xml

    def _buildRelationshipsXML(self):
        rels_xml = ""
        rels_xml += '<?xml version="1.0" encoding="UTF-8"?>' + "\n"
        rels_xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        rels_xml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        rels_xml += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        rels_xml += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        rels_xml += "\n"
        rels_xml += '</Relationships>'

        return rels_xml

    def _buildWorkbookXML(self):
        i = 0
        workbook_xml = ""
        workbook_xml += '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "\n"
        workbook_xml += '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        workbook_xml += '<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>'
        workbook_xml += '<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>'
        workbook_xml += '<sheets>'

        for sheet in self._sheets:
            sheetname = self.sanitize_sheetname(self._sheets[sheet]['sheetname'])
            workbook_xml += '<sheet name="' + self.xmlspecialchars(sheetname) + '" sheetId="' + str(i + 1) + '" state="visible" r:id="rId' + str(i + 2) + '"/>'
            i += 1

        workbook_xml += '</sheets>'
        workbook_xml += '<definedNames>'

        for sheet in self._sheets:
            if self._sheets[sheet]['auto_filter']:
                sheetname = self.sanitize_sheetname(self._sheets[sheet]['sheetname'])
                workbook_xml += '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\'' + self.xmlspecialchars(sheetname) + '\'!' + self._sheets[sheet][
                    'auto_filter'] + '</definedName>'
                i += 1

        workbook_xml += '</definedNames>'
        workbook_xml += '<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>'

        return workbook_xml

    def _buildWorkbookRelsXML(self):
        i = 0
        wkbkrels_xml = ""
        wkbkrels_xml += '<?xml version="1.0" encoding="UTF-8"?>' + "\n"
        wkbkrels_xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        wkbkrels_xml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'

        for sheet in self._sheets:
            wkbkrels_xml += '<Relationship Id="rId' + str(i + 2) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/' + (
            self._sheets[sheet]['xmlname']) + '"/>'
            i += 1

        wkbkrels_xml += "\n"
        wkbkrels_xml += '</Relationships>'

        return wkbkrels_xml

    def _buildContentTypesXML(self):
        content_types_xml = ""
        content_types_xml += '<?xml version="1.0" encoding="UTF-8"?>' + "\n"
        content_types_xml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        content_types_xml += '<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        content_types_xml += '<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'

        for sheet in self._sheets:
            content_types_xml += '<Override PartName="/xl/worksheets/' + (
            self._sheets[sheet]['xmlname']) + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'

        content_types_xml += '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        content_types_xml += '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        content_types_xml += '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        content_types_xml += '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        content_types_xml += "\n"
        content_types_xml += '</Types>'

        return content_types_xml

    @staticmethod
    def xlsCell(row_number, column_number, absolute=False):
        if column_number < 26 and absolute == False:
            return chr(column_number + 65) + str(row_number + 1)

        n = column_number
        r = ""

        while n >= 0:
            r = chr(n % 26 + 65) + r
            n = (n // 26) - 1

        if absolute:
            return "${}${}".format(r, row_number + 1)

        return r + str(row_number + 1)

    @staticmethod
    def sanitize_filename(filename):
        filename = str(filename)
        invalid_chars = ('<', '>', '?', '"', ':', '|', '\\', '/', '*', '&')
        for s in invalid_chars:
            filename = filename.replace(s, '_')

        return filename

    @staticmethod
    def sanitize_sheetname(sheetname):
        badchars = '\\/?*:[]\\'
        goodchars = '        '

        sheetname = str(sheetname).strip().translate(str.maketrans(badchars, goodchars))

        return sheetname[:30] if sheetname else 'Sheet' + str(random.randint(1, 255))

    @staticmethod
    def xmlspecialchars(val):
        if val is None:
            return ""

        if type(val) != str:
            return str(val)

        badchars = "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0c\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f\x7f"
        goodchars = "                              "

        return html_special_chars(val.translate(str.maketrans(badchars, goodchars)))

    @staticmethod
    def __determineNumberFormatType(num_format):
        num_format = re.sub(r"(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)", "", num_format, flags=re.I)

        if num_format == 'GENERAL': return 'n_auto'
        if num_format == '@': return 'n_string'
        if num_format == '0': return 'n_numeric'

        if re.search(r'[H]{1,2}:[M]{1,2}(?![^"]*\+")', num_format, flags=re.I) is not None: return 'n_datetime'
        if re.search(r'[M]{1,2}:[S]{1,2}(?![^"]*\+")', num_format, flags=re.I): return 'n_datetime'
        if re.search(r'[Y]{2,4}(?![^"]*\+")', num_format, flags=re.I): return 'n_date'
        if re.search(r'[D]{1,2}(?![^"]*\+")', num_format, flags=re.I): return 'n_date'
        if re.search(r'[M]{1,2}(?![^"]*\+")', num_format, flags=re.I): return 'n_date'
        if re.search(r'$(?![^"]*\+")', num_format): return 'n_numeric'
        if re.search(r'%(?![^"]*\+")', num_format): return 'n_numeric'
        if re.search(r'0(?![^"]*\+")', num_format): return 'n_numeric'

        return 'n_auto'

    @staticmethod
    def __numberFormatStandardized(num_format):
        if num_format == 'money':
            num_format = 'dollar'
        elif num_format == 'number':
            num_format = '# ###.##'
        elif num_format == 'string':
            num_format = '@'
        elif num_format == 'integer':
            num_format = '0'
        elif num_format == 'date':
            num_format = 'YYYY-MM-DD'
        elif num_format == 'datetime':
            num_format = 'YYYY-MM-DD HH:MM:SS'
        elif num_format == 'price':
            num_format = '#,##0.00'
        elif num_format == 'dollar':
            num_format = '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00'
        elif num_format == 'euro':
            num_format = '#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]'

        ignore_until = ''
        escaped = ''

        for i in range(0, len(num_format)):
            c = num_format[i]
            if ignore_until == '' and c == '[':
                ignore_until = ']'
            elif ignore_until == '' and c == '"':
                ignore_until = '"'
            elif ignore_until == c:
                ignore_until = ''
            if ignore_until == '' and c in ('-', '(', ')') and (i == 0 or num_format[i - 1] != '_'):  # ' '
                escaped += "\\" + c
            else:
                escaped += c
        return escaped

    @staticmethod
    def add_to_list_get_index(haystack, needle):
        if needle in haystack:
            existing_idx = haystack.index(needle)
        else:
            existing_idx = len(haystack)
            haystack.append(needle)

        return existing_idx

    @staticmethod
    def convert_date_time(date_input):
        days = seconds = 0
        year = month = day = 0
        hour = _min = sec = 0
        date_time = date_input

        matches = re.search(r'(\d{4})-(\d{2})-(\d{2})', date_time)
        if matches:
            (year, month, day) = map(int, matches.groups())

        matches = re.search(r'(\d{2}):(\d{2}):(\d{2})', date_time)
        if matches:
            (hour, _min, sec) = map(int, matches.groups())
            seconds = (hour * 3600 + _min * 60 + sec) / 86400

        ymd = "{}-{}-{}".format(year, month, day)
        if ymd == '1899-12-31':  return seconds
        if ymd == '1900-01-00':  return seconds
        if ymd == '1900-02-29':  return seconds + 60

        epoch = 1900
        offset = 0
        norm = 300
        _range = year - epoch

        leap = 1 if ((year % 400 == 0) or ((year % 4 == 0) and (year % 100 != 0))) else 0
        mdays = (31, (29 if leap == 1 else 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)

        if year < epoch or year > 9999: return 0
        if month < 1 or month > 12: return 0
        if day < 1 or day > mdays[month - 1]: return 0

        days = day
        days += sum(mdays[0:month - 1])
        days += _range * 365
        days += int(_range / 4)
        days -= int((_range + offset) / 100)
        days += int((_range + offset + norm) / 400)
        days -= leap

        if days > 59: days += 1

        return days + seconds

    def setStyles(self, _styles):
        styles = copy.deepcopy(_styles)
        for key in styles:
            if isinstance(styles[key], list):
                self._styles[key] = []
                for c in styles[key]:
                    number_format_type = c.pop('format', 'GENERAL')
                    number_format = self.__numberFormatStandardized(number_format_type)
                    cell_style_idx = self.__addCellStyle(number_format, cell_style_string=json.dumps(c))
                    self._styles[key].append((cell_style_idx, self.__determineNumberFormatType(number_format)))
            elif isinstance(styles[key], dict):
                number_format_type = styles[key].pop('format', 'GENERAL')
                number_format = self.__numberFormatStandardized(number_format_type)
                cell_style_idx = self.__addCellStyle(number_format, cell_style_string=json.dumps(styles[key]))
                self._styles[key] = (cell_style_idx, self.__determineNumberFormatType(number_format))

        self._styleFontIndexes()
