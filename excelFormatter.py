import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Alignment, Side

class ExcelFormatter:
    def formatCellAll(ws, cellRange, ft=Font(), fl=PatternFill(), bd=Border(), al=Alignment()):
        for row in ws[cellRange]:
            for cell in row:
                cell.font = ft
                cell.fill = fl
                cell.border = bd
                cell.alignment = al
                            
    def formatCellFont(ws, cellRange, ft=Font()):
        for row in ws[cellRange]:
            for cell in row:
                cell.font = ft
    
    def formatCellFill(ws, cellRange, fl=PatternFill()):
        for row in ws[cellRange]:
            for cell in row:
                cell.fill = fl
    
    def formatCellBorder(ws, cellRange, bd=Border()):
        for row in ws[cellRange]:
            for cell in row:
                cell.border = bd
                
    def formatCellAlign(ws, cellRange, al=Alignment()):
        for row in ws[cellRange]:
            for cell in row:
                cell.alignment = al
                
    def setColumnWidth(ws, col, width):
        for column in openpyxl.utils.get_column_interval(col[0], col[1]):
            ws.column_dimensions[column].width = width
            
    def fitColumnWidth(ws, cellRange):
        dim = {}
        for row in ws[cellRange]:
            for cell in row:
                dim[cell.column_letter] = max(dim.get(cell.column_letter, 0), len(str(cell.value)))
        for col, value in dim.items():
            ws.column_dimensions[col].width = value
            
    def formatCellType(ws, cellRange, format):
        for row in ws[cellRange]:
            for cell in row:
                cell.number_format = format
            
    '''
# These are the constructors for the formatting objects
Font(name='Calibri', size=11, bold=False, italic=False, vertAlign=None, underline='none', strike=False, color='FF000000')

PatternFill(fill_type=None, start_color='FFFFFFFF', end_color='FF000000')

Border(left=Side(border_style=None, color='FF000000'),
        right=Side(border_style=None, color='FF000000'),
        top=Side(border_style=None, color='FF000000'),
        bottom=Side(border_style=None, color='FF000000'),
        diagonal=Side(border_style=None, color='FF000000'), diagonal_direction=0,
        outline=Side(border_style=None, color='FF000000'),
        vertical=Side(border_style=None, color='FF000000'),
        horizontal=Side(border_style=None, color='FF000000'))
        
Alignment(horizontal='general', vertical='bottom', text_rotation=0,
        wrap_text=False, shrink_to_fit=False, indent=0)
'''