import win32com.client as win32
class Interface(object):

    def __init__(self, file_name, sheet_name, sheet_data):

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

    def get_workbook(self, app_instance):

        # grabs the workbook if it is running
        # or not open, then it opens it
        # self.workbook = blah
        _found_workbook = False
        
        # figure out if the workbook is already open
        if app_instance.Workbooks.Count > 0:
            for i in range(1, app_instance.Workbooks.Count + 1):
                if app_instance.Workbooks.Item(i).FullName == self.filename:
                    _workbook = app_instance.Workbooks.Item(i)
                    _found_workbook = True
                    break
                else:
                    continue
                # end if
            # end for
        # end if

        # if it did not find the workbook, then open it
        if not _found_workbook:
            _workbook = app_instance.Workbooks.Open(self.filename)
        # end if
        else:
            _workbook = app_instance.Workbooks.Open(self.filename)
        # end if

        return _workbook
        
    # end load_workbook

    def update_sheet(self, sheet_data):

        pass

    # end update_sheet
# end Interface

