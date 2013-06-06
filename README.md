This module defines two classes.

####Table class 
 
- Intermediate b/w db records, html-table, xls records
- A Table() object is simply a table of rows and columns
- Initialised by queryset
- It can be shown as html-table using this format:

        put {% load tablefilters %} once in your template file

    To show table (an instance of Table) as a html-table:

        {{ table|attributes:" id='...' class='...' "|as_html}}
 
    where attributes and as_html are filters.
- table can be written to xls file using its to_xls() method
    
####xlspdfGenerator

- Generating xls/pdf file or HttpResponse using table.
