import utils
import time
import sys

if sys.platform == "win32":
    import win32com.client as win32
    import pythoncom
    
    # not sure if this will screw everything up
    #pythoncom.CoInitialize()
else:
    raise Exception("Excel interface only works on windows systems.")
# end if

class Interface(object):

    def __init__(self, file_name, sheet_name=None, CoInitialize=False):

        self.filename = file_name
        self.sheet_name = sheet_name
        self.workbook = None
        self.app = None

        if CoInitialize:
            pythoncom.CoInitialize()

    def get_app(self):
        if self.app is not None:
            return self.app
        else:
            self.app = win32.Dispatch("Excel.Application")
            self.app.Visible = True
            return self.app

    def disable_screenupdates(self):
        self.app.ScreenUpdating = False

    def enable_screenupdates(self):
        self.app.ScreenUpdating = True

    def disable_alerts(self):
        self.app.DisplayAlerts = False

    def enable_alerts(self):
        self.app.DisplayAlerts = True

    def disable_calculation(self):
        self.app.Calculation = -4135
    
    def enable_calculation(self):
        self.app.Calculation = -4105

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
        

    def update_range(self, sheet_name, range_dict):
    
        # gets the different attributes for the excel range,
        # we need the two cells and the data matrix
        _cell1_row = range_dict['cell1']['row']
        _cell1_col = range_dict['cell1']['column']
        _cell2_row = range_dict['cell2']['row']
        _cell2_col = range_dict['cell2']['column']

        # gets the workbook
        _workbook = self.get_workbook()

        # turn all of the shit off
        #self.disable_screenupdates()
        self.disable_alerts()
        self.disable_calculation()
        
        # gets the worksheet
        _worksheet = _workbook.Worksheets(sheet_name)

        # gets the worksheet cells
        _cell1 = _worksheet.Cells(_cell1_row, _cell1_col)
        _cell2 = _worksheet.Cells(_cell2_row, _cell2_col)

        # writes the range
        _worksheet.Range(_cell1, _cell2).Value = range_dict['data']

    def save_workbook(self):
        _workbook = self.get_workbook()

        # turn the shit back on
        #self.enable_screenupdates()
        self.enable_alerts()
        self.enable_calculation()

        _workbook.Save()
    # end save_workbook
    
# end Interface

