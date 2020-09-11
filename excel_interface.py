import win32com.client as win32
class Interface(object):

    def __init__(self, file_name, sheet_name):

        self.filename = file_name
        self.sheet_name = sheet_name
        self.workbook = None
        self.app = None

    # end __init__()

    def get_app(self):
        if self.app is not None:
            return self.app
        else:
            self.app = win32.gencache.EnsureDispatch("Excel.Application")
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

    def update_sheet(self, sheet_data):

        # gets the workbook
        _workbook = self.get_workbook()

        # gets the worksheet
        _worksheet = _workbook[self.sheet_name]

    # end update_sheet

    def save_workbook(self):
        _workbook = self.get_workbook()
        _workbook.Save()
    # end save_workbook
    
# end Interface

