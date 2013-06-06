This module defines two classes.

- Table class 
 
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

- xlspdfGenerator

    1)Generating xls/pdf file or HttpResponse using table.
