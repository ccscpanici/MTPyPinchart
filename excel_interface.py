import utils
import time
import sys

if sys.platform == "win32":
    import win32com.client as win32
    import pythoncom
else:
    raise Exception("Excel interface only works on windows systems.")
# end if

class Interface(object):

    def __init__(self, file_name, sheet_name=None):

        self.filename = file_name
        self.sheet_name = sheet_name
        self.workbook = None
        self.app = None

    # end __init__()

    def get_app(self):
        if self.app is not None:
            return self.app
        else:
            
            # not sure if this will screw everything up
            pythoncom.CoInitialize()
            
            self.app = win32.gencache.EnsureDispatch("Excel.Application")
            self.app.Visible = True
            return self.app
        # end if
    # end get_app()

    def get_workbook(self):

        if self.workbook is None:
            # gets the application instance
            _app_instance = self.get_app()

            # grabs the workbook if it is running
            # or not open, then it opens it
            # self.workbook = blah
            _found_workbook = False
            
            # figure out if the workbook is already open
            if _app_instance.Workbooks.Count > 0:
                for i in range(1, _app_instance.Workbooks.Count + 1):
                    if _app_instance.Workbooks.Item(i).FullName == self.filename:
                        _workbook = _app_instance.Workbooks.Item(i)
                        _found_workbook = True
                        break
                    else:
                        continue
                    # end if
                # end for
            # end if

            # if it did not find the workbook, then open it
            if not _found_workbook:
                _workbook = _app_instance.Workbooks.Open(self.filename)
            # end if

            self.workbook = _workbook
        # end if

        return self.workbook
        
    # end load_workbook

    def update_sheet(self, plc_data_column, config_data):

        # gets the workbook
        _workbook = self.get_workbook()

        if self.sheet_name is None:
            raise Exception("No sheet name specified")
        # end if

        # gets the worksheet
        _worksheet = _workbook.Worksheets(self.sheet_name)

        # start updating the worksheet
        _data_type = plc_data_column['data']['type']
        
        # loop through the PLC data
        _plc_data = plc_data_column['plc_data']

        for i in _plc_data:

            _row = i['value_cell']['row']
            _column = i['value_cell']['column']
            _value = i['value']

            if _data_type == "FLOAT DATA" or _data_type == "STRING DATA" or _data_type == "DINT DATA" or \
                _data_type == "INT DATA":
                # updates the value
                _worksheet.Cells(_row, _column).Value = _value
                
            elif _data_type == "DINT-32 DATA":

                _start_column = _column + config_data['DINT-32 BITS OFFSET']
                _end_column = _start_column + 32

                # get the binary representation in a list
                _bin = utils.get_32bit_bin(_value)

                for i in range(0, 32):
                    _bit_column = _start_column + i
                    _bit_value = _bin[i]
                    if _bit_value == 0:                    
                        # update the bit
                        _worksheet.Cells(_row, _bit_column).Value = ""
                    else:
                        _worksheet.Cells(_row, _bit_column).Value = _bit_value
                    # end if
                # end for

            elif _data_type == "INT-16 DATA":

                _start_column = _column + config_data['INT-16 BITS OFFSET']
                _end_column = _start_column + 16

                # get the binary representation in a list
                _bin = utils.get_16bit_bin(_value)

                for i in range(0, 16):
                    _bit_column = _start_column + i
                    _bit_value = _bin[i]
                    if _bit_value == 0:
                        # update the bit
                        _worksheet.Cells(_row, _bit_column).Value = ""
                    else:
                        _worksheet.Cells(_row, _bit_column).Value = _bit_value
                    # end if
                pass
            # end if
        # end for

    # end update_sheet

    def save_workbook(self):
        _workbook = self.get_workbook()
        _workbook.Save()
    # end save_workbook
    
# end Interface

