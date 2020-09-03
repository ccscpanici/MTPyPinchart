
CONFIG_FLAG = "$Config$"
DATA_HEADER_FLAG = "$HeaderRow$"
OPC_TOPIC_FLAG = "OPC Topic"
OPC_TOPIC_OFFSET = 1

DEFAULT_CONFIG = {
    "ADDRESS OFFSET" : 1,              # offset from the value column
    "DATA ROW OFFSET" : 2,             # offset from the $HeaderRow$ flag, could be a few rows down
    "DINT-32 BITS OFFSET" : -32,       # offset from the value column
    "INT-16 BITS OFFSET" : -16         # offset from the value column
}

VALID_FIELD_TYPES = ['DINT DATA', 'DINT-32 DATA', 'INT DATA', 'INT-16 DATA', 'FLOAT DATA', 'STRING DATA']

class SheetDataBase(object):
    def __init__(self, sheet_dict):
        self.sheet = sheet_dict
    # end __init__()

    def search_sheet(self, value):
        for a_row in self.sheet:
            for a_cell in a_row:
                _value = a_cell['value']
                if _value is not None:
                    if str(a_cell['value']).lower() == value.lower():
                        return a_cell
                    # end if
                # end if
            # end for
        # end for
        return False
    # end search sheet

    def concat_cell_values(self, first_cell, span, delimiter=""):
        # this method concatinates the first_cell value
        # with the columns to the right of that cell 
        # in the current span.
        _string_builder = first_cell['value']
        for i in range(1, span):
            _tmp_cell = self.get_cell(first_cell['row'], first_cell['column'] + i)
            _tmp_value = _tmp_cell['value']
            if _tmp_value is None:
                _tmp_value = ""
            # end if
            _string_builder = _string_builder + delimiter + _tmp_value
        # end for
        return _string_builder
    # end concat_cells_values()

    def get_row(self, row):
        if row < 1:
            raise ValueError(message="get_row needs to be one or greater.")
        else:
            return self.sheet[row-1]
        # end if
    # end get_row

    def get_cell(self, row, column):
        if row < 1 or column < 1:
            raise ValueError(message="get_cell needs to be based on 1, 1 being the first cell")
        else:
            # based on 1-value
            return self.sheet[row-1][column-1]
        # end if
    # end get_cell

    def search_in_row(self, row, value):
        for a_cell in row:
            _value = a_cell['value']
            if _value is not None:
                if a_cell['value'].lower() == value.lower():
                    return a_cell
                # end if
            # end if
        # end for
        return False
    # end search_in_row
# end SheetDataBase

class PLCSheetData(SheetDataBase):

    def __init__(self, sheet_dict, config_data):
        super(PLCSheetData, self).__init__(sheet_dict)
        self.config_data = config_data
    # end __init__()

    def get_plc_data(self):

        _header_row = self.get_header_row()
        _data_structure = self.get_plc_data_structure(_header_row)

        for a_plc_range in _data_structure:


        # end for
        
        return None

    # end get_plc_data

    def get_header_row(self):
        _header_cell = self.search_sheet(DATA_HEADER_FLAG)
        _header_row = self.get_row(_header_cell['row'])
        return _header_row
    # end get_header_row

    def get_plc_data_structure(self, header_row):
        
        # plc data value columns
        _columns = []

        # converts the types to lower case for comparison
        _valid_types = []
        for a_type in VALID_FIELD_TYPES:
            _valid_types.append(a_type.lower())
        # end for

        # get the header row
        _header_row = self.get_header_row()
        
        for a_cell in _header_row:
            _cell_value = a_cell['value']
            if _cell_value is not None:
                if _cell_value.lower in _valid_types:
                    # save the column

                    # this is the cell that should have the address
                    _address_cell = self.get_cell(a_cell['row'], a_cell['column'] + self.config_data['ADDRESS OFFSET'])
                    _address_span = _address_cell['col_span']
                    _data_row_start = a_cell['row'] + self.config_data['DATA ROW OFFSET']
                    
                    _column_data = {
                        'value':{
                            'column':a_cell['column'],
                            'col_span':a_cell['col_span']
                        },
                        'address': {
                            'column':_address_cell['column'],
                            'col_span':_address_cell['col_span']
                        },
                        'data': {
                            'type':_cell_value,
                            'start_row':_data_row_start
                        }
                    }
                    if _cell_value == "int-16 data":
                        _bits_start = a_cell['column'] + self.config_data['INT-16 BITS OFFSET']
                        _bits_end = _bits_start + 16
                        _column_data['bits'] = {
                            'start_col':_bits_start,
                            'end_col':_bits_end
                        }
                    elif _cell_value == 'dint-32 data':
                        _bits_start = a_cell['column'] + self.config_data['DINT-32 BITS OFFSET']
                        _bits_end = _bits_start + 32
                        _column_data['bits'] = {
                            'start_col':_bits_start,
                            'end_col':_bits_end
                        }
                    else:
                        pass
                    # end if

                    _columns.append(_column_data)
                # end if
            # end if
        # end for
        return _columns
    # end get_plc_value_column_numbers
# end PLCSheetData

class MainSheetData(SheetDataBase):

    config_data = None
    opc_topic = None

    def get_config_cell(self):
        return self.search_sheet(CONFIG_FLAG)
    # end get_config_cell

    def get_opc_topic(self):
        if self.opc_topic is None:
            # search for the flag
            _cell1 = self.search_sheet(OPC_TOPIC_FLAG)
            _cell2 = self.get_cell(_cell1['row'], _cell1['column'] + OPC_TOPIC_OFFSET)
            _value = _cell2['value']
            if _value is not None:
                self.opc_topic = _value
            else:
                raise Exception("Could not locate valid OPC Topic in sheet.")
            # end if
        # end if
        return self.opc_topic
    # end get_opc_topic

    def get_config_data(self):
        if self.config_data is None:

            # instatiate the data
            _config_data = {}

            # get the config cell
            _config_cell = self.get_config_cell()
            if not _config_cell:
                return None
            else:
                _config_row_number = _config_cell['row']
                _config_row_data = self.get_row(_config_row_number)

                # get the key column cell
                _key_cell = self.search_in_row(_config_row_data, "key")
                _key_column = _key_cell['column']

                # get the value column cell
                _value_cell = self.search_in_row(_config_row_data, "value")
                _value_column = _value_cell['column']

                # loop through and get all of the key / value pairs in the
                # configuration. Stop when you get to a blank key.
                _end_of_config = False
                _key_index = 0
                while not _end_of_config:
                    _k = self.get_cell(_config_row_number + 1 + _key_index, _key_column)['value']
                    if _k is not None:
                        _v = self.get_cell(_config_row_number + 1 + _key_index, _value_column)['value']
                        _config_data[_k] = _v
                    else:
                        _end_of_config = True
                    # end if
                    _key_index = _key_index + 1
                # end while
                self.config_data = _config_data
                return self.config_data
            # end if
        # end if
    # end get_config_data

    def process_configuration_data(self, data):
        # process configuration data makes sure that the required
        # keys for configuration are in place. If they are not, it
        # will use the default values
        if data is None:
            data = {}
        # end if

        # loop through the keys
        for a_key in DEFAULT_CONFIG:
            if not a_key in data:
                data[a_key] = DEFAULT_CONFIG[a_key]
            # end if
        return data
    # end if

    def validate_configuration_data(self, data):
        # similar to process, just returns a true or
        # a faulse if the required keys are not in the data
        for a_key in DEFAULT_CONFIG:
            if not a_key in data:
                return False
            # end if
        return True
    # end validate_configuration_data
# end class

# if __name__ == '__main__':

#     import filereader
#     import os

#     file = os.getcwd() + "\\RO Pinchart.xlsm"

#     print("Loading workbook....")
#     # gets a dictionary of the whole workbook
#     workbook = filereader.get_xl_data(file)
#     print("Loading workbook....Done.")
    
#     # get the sheet names
#     sheet_names = list(workbook.keys())

#     _sheet_name = sheet_names[0]

#     if _sheet_name.lower() == 'main program':
#         _sheet_data = workbook[_sheet_name]
#         _sheet_object = MainSheetData(_sheet_data)

#         _config_data = _sheet_object.get_config_data()
#         _config_data = _sheet_object.process_configuration_data(_config_data)
#         if not _sheet_object.validate_configuration_data(_config_data):
#             print("we're all fucked.")
#         else:
#             print("looks. Good.")
#         # end if
#         print(_config_data)
#         print(_sheet_object.get_opc_topic())