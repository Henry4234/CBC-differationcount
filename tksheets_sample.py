from tksheet import *
import tkinter as tk
from datetime import datetime, date, timedelta, time
from dateutil import parser, tz
from math import ceil
import re

date_replace = re.compile('|'.join(['\(', '\)', '\[', '\]', '\<', '\>']))

# --------------------- Custom formatter methods ---------------------
def round_up(x):
    try: # might not be a number if empty
        return float(ceil(x))
    except:
        return x

def only_numeric(s):
    return ''.join(n for n in f"{s}" if n.isnumeric() or n == '.')

def convert_to_local_datetime(dt: str, **kwargs):
    if isinstance(dt, datetime):
        pass
    elif isinstance(dt, date):
        dt = datetime(dt.year, dt.month, dt.day)
    else:
        if isinstance(dt, str):
            dt = date_replace.sub("", dt)
        try:
            dt = parser.parse(dt)
        except:
            raise ValueError(f"Could not parse {dt} as a datetime")
    if dt.tzinfo is None:
        dt.replace(tzinfo = tz.tzlocal())
    dt = dt.astimezone(tz.tzlocal())
    return dt.replace(tzinfo = None)

def datetime_to_string(dt: datetime, **kwargs):
    return dt.strftime('%d %b, %Y, %H:%M:%S')

# --------------------- Custom Formatter with additional kwargs ---------------------

def custom_datetime_to_str(dt: datetime, **kwargs):
    return dt.strftime(kwargs['format'])


class demo(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.grid_columnconfigure(0, weight = 1)
        self.grid_rowconfigure(0, weight = 1)
        self.frame = tk.Frame(self)
        self.frame.grid_columnconfigure(0, weight = 1)
        self.frame.grid_rowconfigure(0, weight = 1)
        self.sheet = Sheet(self.frame,
                           empty_vertical = 0,
                           empty_horizontal = 0,
                           )
        self.sheet.enable_bindings()
        self.sheet.set_sheet_data(data = [[f"{r}"]*11 for r in range(20)],reset_col_positions=True,reset_row_positions=True)
        self.frame.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.grid(row = 0, column = 0, sticky = "nswe")
        self.sheet.headers(['Non-Nullable Float Cell\n1 decimals places', 
                            'Float Cell', 
                            'Int Cell', 
                            'Bool Cell', 
                            'Percentage Cell\n0 decimal places', 
                            'Custom Datetime Cell',
                            'Custom Datetime Cell\nCustom Format String',
                            'Float Cell that\nrounds up', 
                            'Float cell that\n strips non-numeric', 
                            'Dropdown Over Nullable\nPercentage Cell', 
                            'Percentage Cell\n2 decimal places'])
        
        # ---------- Some examples of cell formatting --------
        self.sheet.format_cell('all', 0, formatter_options = float_formatter(nullable = False))
        self.sheet.format_cell('all', 1, formatter_options = float_formatter())
        self.sheet.format_cell('all', 2, formatter_options = int_formatter())
        self.sheet.format_cell('all', 3, formatter_options = bool_formatter(truthy = truthy | {"nah yeah"}, falsy = falsy | {"yeah nah"}))
        self.sheet.format_cell('all', 4, formatter_options = percentage_formatter())


        # ---------------- Custom Formatters -----------------
        # Custom using generic formatter interface
        self.sheet.format_cell('all', 5, formatter_options = formatter(datatypes = datetime, 
                                                                       format_function = convert_to_local_datetime, 
                                                                       to_str_function = datetime_to_string, 
                                                                       nullable = False,
                                                                       invalid_value = 'NaT',
                                                                       ))
        # Custom format
        self.sheet.format_cell('all', 6, datatypes = datetime, 
                                         format_function = convert_to_local_datetime, 
                                         to_str_function = custom_datetime_to_str, 
                                         nullable = True,
                                         invalid_value = 'NaT',
                                         format = '(%Y-%m-%d) %H:%M %p'
                                         )
        
        # Unique cell behaviour using the post_conversion_function
        self.sheet.format_cell('all', 7, formatter_options = float_formatter(post_format_function = round_up))
        self.sheet.format_cell('all', 8, formatter_options = float_formatter(), pre_format_function = only_numeric)

        self.sheet.create_dropdown('all', 9, values = ['', '104%', .24, "300%", 'not a number'], set_value = 1)
        self.sheet.format_cell('all', 9, formatter_options = percentage_formatter(), decimals = 0)
        self.sheet.format_cell('all', 10, formatter_options = percentage_formatter(decimals = 5))
        lst = []
        for x in range(18,-1,-1):
            lst.append(x)
        print(lst)
        self.sheet.insert_column(values=lst,idx="end")


app = demo()
app.mainloop()