"""
This module defines two classes.
1.) Table class 
Use: 
    1)Intermediate b/w db records, html-table, xls records
    2)A Table() object is simply a table of rows and columns
    3)Initialised by queryset
    4) * Initialised by reading xls file: this feature to be added
    5)It can be shown as html-table using this format:
        put {% load tablefilters %} once in your template file
      To show table (an instance of Table) as a html-table:
        {{ table|attributes:" id='...' class='...' "|as_html}}
      where attributes and as_html are filters.
    6)table can be written to xls file using its to_xls() method

2.) xlspdfGenerator
Use:
    1)Generating xls/pdf file or HttpResponse using table. 
"""

import xlwt
import xlrd
from datetime import datetime
from django.db.models.query import QuerySet
from django.http import HttpResponse
from django.template import Context, Template
from django.utils.safestring import mark_safe
import ho.pisa as pisa

from reportlab.lib.pagesizes import letter
from reportlab.platypus.doctemplate import SimpleDocTemplate,Paragraph,Spacer
from reportlab.rl_config import defaultPageSize
from reportlab.platypus import Table as pdfTable,TableStyle
from reportlab.lib import colors
from reportlab.lib.units import inch

def default_handler(model_instance,column):
  return getattr(model_instance,column)
  
class Table(object):
  def __init__(self,queryset,columns,serialize = False,handler = default_handler,
               verbose_names_dict={},col_width_dict={},**kwargs):
    """
    queryset   : any model's queryset
    columns    : a list of columns(strings)
                 column either be a field of that model or a user-defined column-name
    serialize  : Boolean - To set first column as of 'Serial No.'
    handler    : if not given then all columns are assumed to be the fields of that model
                 if user adds some own-defined columns then user should define a handler function
                 that returns a value and pass this function in parameters list.
    
    verbose_names_dict : dict that maps columns to Column-Names that are shown in html-table and xls
                         if not given then it will automatically search in model's verbose_name of fields
                         even if it is not found in model's verbose_name for a field it will simply capitalize 
                         the first letter of first word of column-name.
                         (default {} )
      
    col_width_dict : dict that maps columns to their widths(integers in pixels used in html)
                     widths of html-table columns and xls columns then would get defined.
                     Use '_serial_' as key in col_width_dict to define width of column 'Serial No.'
                     (essential to be given if table is going to be rendered as html using as_html filter)
                     
    * columns and rows indexing starts with 0    
    """
    assert isinstance(queryset,QuerySet), "Not a Queryset"
    self.queryset = queryset
    self.columns = columns
    self.serialize = serialize
    self.handler = handler
    self.html_attributes = ""
    self.verbose_names_dict = verbose_names_dict
    self.col_width_dict = col_width_dict
 
  def no_of_columns(self):
    """ Returns number of columns
    """    
    return len(self.columns) + (1 if self.serialize else 0)
  
  def no_of_rows(self):
    """ Returns number of rows
    """    
    return len(self.queryset) + 1 
  
  def col_width(self,column_no):
    """ Returns column width for given column number
    """    
    if(column_no == 0 and self.serialize):
      return self.col_width_dict['_serial_']  
    column = self.columns[column_no - (1 if self.serialize else 0)]
    return self.col_width_dict[column]
  
  def total_width(self):
    """ Returns total width of all columns
    """    
    total = 0
    for i in range(self.no_of_columns()):
      total += self.col_width(i)
    return total

  def col_width_percent(self,column_no):
    """ Returns column width percentage for given column number
    """    
    return float(self.col_width(column_no)*100)/self.total_width()

  def cell(self,row_no,column_no):
    """Return value of the cell for given row number and column number
    """  
    if row_no == 0:
      if self.serialize and column_no == 0:
       if self.verbose_names_dict.has_key('_serial_'):
         return self.verbose_names_dict['_serial_']
       else:  
         return "S.No."
      else:
        column = self.columns[column_no - (1 if self.serialize else 0)]
        if column in self.verbose_names_dict:
          return self.verbose_names_dict[column]
        else:
          try:
            return self.queryset.model._meta.get_field(column).verbose_name.capitalize()
          except Exception as e:
            return column.capitalize()
    else:
      if column_no == 0:
        return str(row_no)
      else:
        entrant = self.queryset[row_no - 1]
        column = self.columns[column_no - (1 if self.serialize else 0)]   
        return str(self.handler(entrant,column))

  def to_xls(self,ws,start_row = 0,start_col = 0,width_ratio = 1):
    """ Writes self's content to xls file
        ws : Worksheet (* refer xlwt module)
        start_row : start row number in xls file to write data from
                    (default 0)
        start_col : start column number in xls file to write data from
                    (default 0)
        width_ratio : adjustable multiplying quantity to given col_width in col_width_dict
                      (default 1)
    """  
    if self.col_width_dict:  
      for c in range(self.no_of_columns()):
        ws.col(start_col+c).width = int(35*self.col_width(c)*width_ratio); 
  
    boldstyle = xlwt.XFStyle()
    boldstyle.font.bold = True
  
    for r in range(self.no_of_rows()):
      for c in range(self.no_of_columns()):
        if r == 0:
          ws.write(start_row + r,start_col + c,self.cell(r,c),boldstyle)
        else:
          ws.write(start_row + r,start_col + c,self.cell(r,c))
  
  def __unicode__(self):
    """ Returns queryset as display
    """  
    return str(self.queryset)



class xlspdfGenerator(object):
  """class to generate xls/pdf file or HttpResponse using table. 
  """  
  html_template = """
  <!DOCTYPE HTML PUBLIC>
  <html>
   <head>
    <style type="text/css">
     #tb{
      border:1px solid #aaa;      
      border-collapse:collapse;
      background:#efefef; 
      font:13px sans-serif;           
     }
     #tb tr th,#tb tr td{
      text-align:left;
      padding:2px 2px 2px 5px;
     } 
     .heading{
      width:{{table.total_width}}px;
      height:20px;
      padding:2px 2px 2px 5px;           
      font:bold 14px sans-serif; 
      text-align:center;      
     } 
     .date{
      width:{{table.total_width}}px;
      height:20px;
      padding:2px 2px 2px 5px;           
      font:bold 14px sans-serif; 
     } 
    </style>
   </head>
   <body>
    {% if date %}
     <div class="date">Date : {{date}}</div>
    {% endif %}
    {% for heading in headings %}
    <div class="heading">
     {{heading}}
    </div> 
    {% endfor %}
    {% load tablefilters %}
    {{ table|attributes:"id='tb' border='1' "|as_html}}
   </body>
  </html> 
  """
  def __init__(self,table,filename,headings=[],date = datetime.today().strftime('%d/%m/%Y')):
    """
       table : Table instance
       filename = name of the xls/pdf file to be saved with. (without extension)
       headings = list of string headings to be shown above the table.
       date = date to be shown on upper left corner of file.
              if not given today's date is used by default.
              pass date = None in argument list if you don't want date to be shown in file.
    """  
    self.table = table
    self.filename = filename
    self.headings = headings
    self.date = date
  
  def generate_xls(self):
    """
       generates xls and returns xls WorkBook object.
    """
    self.wb = xlwt.Workbook()
    ws = self.wb.add_sheet('Sheet1')
    heading_style = xlwt.easyxf('font: bold true; alignment: horizontal center, wrap true;')
    extra_row = 0
    if self.date:
      date_style = xlwt.easyxf('font: bold true; alignment: horizontal left, wrap true;')
      ws.write_merge(0,0,0,self.table.no_of_columns()-1,'Date : '+self.date,date_style)   
      extra_row = 1
    for i in range(len(self.headings)):
      ws.write_merge(i+extra_row,i+extra_row,0,self.table.no_of_columns()-1,self.headings[i],heading_style)
    ws.set_panes_frozen(True)
    ws.set_horz_split_pos(len(self.headings)+extra_row+1)
    ws.set_remove_splits(True)
    self.table.to_xls(ws,start_row=len(self.headings)+extra_row,start_col=0)
    return self.wb
      
  def xls_response(self):
    """
       returns xls HttpResponse.
    """
    response = HttpResponse(mimetype="application/ms-excel")
    response['Content-Disposition'] = 'attachment; filename='+ self.filename + '.xls'      
    self.generate_xls()
    self.wb.save(response)
    return response
  
  def save_xls(self,basepath=''):
    """
       basepath = basepath ending with forward slash.
       saves xls file.  
    """  
    self.generate_xls()
    self.wb.save(basepath+self.filename+'.xls')

  def html_response(self):
    """
       returns html HttpResponse.
    """
    t = Template(self.html_template)  
    c = Context({'table':self.table,'headings':self.headings,'date':self.date})
    return HttpResponse(t.render(c))


  def pdf_response(self):
    """
       returns pdf HttpResponse.
    """
    response = HttpResponse(mimetype="application/pdf")
    response['Content-Disposition'] = 'filename='+ self.filename + '.pdf'
    doc = SimpleDocTemplate(response,topMargin = inch/2,bottomMargin = inch/2,leftMargin = inch/2, rightMargin = inch/2)      
    elements = []
    data = []
    for i in range(self.table.no_of_rows()):
      q = []
      for j in range(self.table.no_of_columns()):
        q.append(self.table.cell(i,j))
      data.append(q) 
    header = []
    if self.date:
      header.append(['Date : '+self.date])
    for heading in self.headings:
      header.append([heading])
    header.append([''])
    er = len(header)    
    width,height = defaultPageSize
    width = width - inch    
    t=pdfTable(header+data,[int(width*self.table.col_width_percent(i)/100.) for i in range(self.table.no_of_columns())])
    style_list = []
    for i in range(len(header)):
      style_list.append(('SPAN',(0,i),(-1,i)))
    style_list+=[('ALIGN',(0,1 if self.date else 0),(-1,er-1),'CENTER'),
                 ('FONT',(0,0),(-1,er),'Helvetica-Bold'),
                 ('INNERGRID', (0,er), (-1,-1), 1, colors.black),
                 ('BOX', (0,er), (-1,-1), 1, colors.black),
                 ('BACKGROUND',(0,er),(-1,-1),colors.HexColor('#efefef'))]
    t.setStyle(TableStyle(style_list))
    elements.append(t) 
    doc.build(elements)
    return response
    
