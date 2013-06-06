from django import template
from api.table import Table
from django.utils.safestring import mark_safe

register = template.Library()

@register.filter
def as_html(table):
  """ A filter for Table() object
      return html code for the table
  """  
  if isinstance(table,Table):
    html = "<table width=\"" + str(table.total_width()) + "\"" + table.html_attributes + " ><colgroup>\n"
    if table.col_width_dict:
      for i in range(table.no_of_columns()):
        html += "<col width=\"" + str(table.col_width_percent(i)) + "%\"/>\n"
    html += "</colgroup><tbody>\n"    
    row = "<tr>"
    for c in range(table.no_of_columns()):
      row += "<th width=\""+str(table.col_width_percent(c))+"%\">" + table.cell(0,c) +"</th>"
    row += "</tr>\n"
    html += row
    for r in range(1,table.no_of_rows()):
      row = "<tr>"
      for c in range(table.no_of_columns()):
        row += "<td>" + table.cell(r,c) + "</td>"
      row += "</tr>\n"
      html += row
    return mark_safe(html)
  else:
    return table 

@register.filter
def attributes(table,attrs):
  """ To add attributes to html <table>
      Should be used before as_html filter
  """  
  if isinstance(table,Table):
    table.html_attributes = attrs
  return table
