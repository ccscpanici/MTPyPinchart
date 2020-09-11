import numpy
import utils

CONFIG_FLAG = "$Config$"
DATA_HEADER_FLAG = "$HeaderRow$"
OPC_TOPIC_FLAG = "OPC Topic"
OPC_TOPIC_OFFSET = 1
SLC_FILE_FLAG = "PC5 File"
SLC_FILE_OFFSET = 1

DEFAULT_CONFIG = {
    "ADDRESS OFFSET" : 1,              # offset from the value column
    "DATA ROW OFFSET" : 2,             # offset from the $HeaderRow$ flag, could be a few rows down
    "DINT-32 BITS OFFSET" : -32,       # offset from the value column
    "INT-16 BITS OFFSET" : -16         # offset from the value column
}

VALID_FIELD_TYPES = ['DINT DATA', 'DINT-32 DATA', 'INT DATA', 'INT-16 DATA', 'FLOAT DATA', 'STRING DATA']

class SheetDataBase(object):
    def __init__(self, sheet_dict, thread_id):
        self.sheet = sheet_dict
        self.thread_id = thread_id
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

    def get_search_offset_value(self, search_string, row_offset=0, col_offset=0):

        # gets the first cell searching for the string
        # that we need
        _cell1 = self.search_sheet(search_string)

        # if the search function returns false, then this function
        # returns none
        if not _cell1:
            return None
        else:
            # if the search function does not return false, then we get the
            # second cell with the offsets and return the cell value
            _cell2 = self.get_cell(_cell1['row'] + row_offset, _cell1['column'] + col_offset)
            return _cell2['value']
        # end if
    # end get_search_offset_value

    def concat_cell_values(self, first_cell, span, delimiter=""):
        # this method concatinates the first_cell value
        # with the columns to the right of that cell 
        # in the current span.
        _string_builder = first_cell['value']
        if _string_builder is None:
            return None
        # end if

        for i in range(1, span):
            _tmp_cell = self.get_cell(first_cell['row'], first_cell['column'] + i)
            _tmp_value = str(_tmp_cell['value'])
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

    def get_cell_in_row(self, row, column_number):
        return row[column_number - 1]
    # end get_cell_in_row

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

    def __init__(self, sheet_dict, config_data, thread_id):
        super(PLCSheetData, self).__init__(sheet_dict,thread_id)
        self.config_data = config_data
        self.opc_topic = None
    # end __init__()

    def get_opc_topic(self):
        return self.opc_topic
    # end get_opc_topic

    def set_opc_topic(self, value):
        if self.opc_topic != value:
            self.opc_topic = value
        # end if
    # end set_opc_topid

    def get_plc_data_for_column(self, plc_data_column):

        # returns dictionary of address
        # and value, as well as the cell
        # references.
        _header_row = self.get_header_row()
        _header_row_number = _header_row[0]['row']
        _data_row_start = plc_data_column['data']['start_row']

        _value_column = plc_data_column['value']['column']
        _value_column_span = plc_data_column['value']['column_span']
        _address_column = plc_data_column['address']['column']
        _address_column_span = plc_data_column['address']['column_span']

        _plc_column_data = []

        for a_row in self.sheet[_data_row_start - 1:]:
            _value_cell = self.get_cell_in_row(a_row, _value_column)
            _value_cell_value = self.concat_cell_values(_value_cell, _value_column_span)

            _address_cell = self.get_cell_in_row(a_row, _address_column)
            _address_cell_value = self.concat_cell_values(_address_cell, _address_column_span)

            _topic = self.get_opc_topic()
            if _topic is not None:
                _full_address = "[%s]%s" % (_topic, _address_cell_value.strip())
            else:
                _full_address = _address_cell_value.strip()
            # end if

            _plc_data = {
                'address':_full_address,
                'value':_value_cell_value,
                'address_cell':_address_cell,
                'value_cell':_value_cell,
                'full_address':_full_address
            }

            _plc_column_data.append(_plc_data)
        # end for
        
        return _plc_column_data

    # end get_plc_data_for_column

    def get_plc_data(self, data_row_offset):

        # returns dictionary of address
        # and value, as well as the cell
        # references.
        _header_row = self.get_header_row()
        _header_row_number = _header_row[0]['row']
        _data_row_start = _header_row_number + data_row_offset
        _data_structure = self.get_plc_data_structure(_header_row)

        for a_plc_range in _data_structure:
            _value_column = a_plc_range['value']['column']
            _value_column_span = a_plc_range['value']['column_span']
            _address_column = a_plc_range['address']['column']
            _address_column_span = a_plc_range['address']['column_span']

            _plc_column_data = []

            for a_row in self.sheet[_data_row_start - 1:]:
                _value_cell = self.get_cell_in_row(a_row, _value_column)
                _value_cell_value = self.concat_cell_values(_value_cell, _value_column_span)

                _address_cell = self.get_cell_in_row(a_row, _address_column)
                _address_cell_value = self.concat_cell_values(_address_cell, _address_column_span)

                _topic = self.get_opc_topic()
                if _topic is not None:
                    _full_address = "[%s]%s" % (_topic, _address_cell_value.strip())
                else:
                    _full_address = _address_cell_value.strip()
                # end if

                _plc_data = {
                    'address':_full_address,
                    'value':_value_cell_value,
                    'address_cell':_address_cell,
                    'value_cell':_value_cell,
                    'full_address':_full_address
                }
                _plc_column_data.append(_plc_data)
            # end for

            a_plc_range['plc_data'] = _plc_column_data
        # end for
        
        return _data_structure

    # end get_plc_data

    def get_address_value_list(self, plc_data, data_type_string):
        # takes in the plc data, converts the values,
        # and puts them in a list of tuples to be written to
        # the opc server.
        _value_type = data_type_string.lower()
        tuple_list = []
        for i in plc_data:

            # converts the value to the correct data type
            _value = utils.data_converter(i['value'], data_type_string)

            # if the address is None, we still have to append it to the
            # values. This is because when we read the values back they
            # are in the same order. The downloader should only get errors
            # for the invalid tags. We should be ok
            _address = i['address']
            tuple_list.append((_address, _value))
        # end for
        return tuple_list
    # end get_address_value_list

    def get_address_list(self, plc_data):
        address_list = []
        for i in plc_data:
            _address = i['address']
            address_list.append(_address)
        # end for
        return address_list
    # end get_address_list

    def get_header_row(self):
        _header_cell = self.search_sheet(DATA_HEADER_FLAG)
        _header_row = self.get_row(_header_cell['row'])
        return _header_row
    # end get_header_row

    def get_header_row_number(self):
        _header_cell = self.search_sheet(DATA_HEADER_FLAG)
        return _header_cell['row']
    # end get_header_row_number

    def get_plc_data_structure(self):

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
                if _cell_value.lower() in _valid_types:
                    # save the column

                    # this is the cell that should have the address
                    _address_cell = self.get_cell(a_cell['row'], a_cell['column'] + self.config_data['ADDRESS OFFSET'])
                    _address_span = _address_cell['column_span']
                    _data_row_start = a_cell['row'] + self.config_data['DATA ROW OFFSET']
                    
                    _column_data = {
                        'value':{
                            'column':a_cell['column'],
                            'column_span':a_cell['column_span']
                        },
                        'address': {
                            'column':_address_cell['column'],
                            'column_span':_address_cell['column_span']
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

    def update_data_with_new_values(self, data_type, plc_data, new_values_from_opc):
        _data_index = 0
        for address, value, quality, timestamp in new_values_from_opc:

            _plc_data_item = plc_data[_data_index]
            _value = utils.data_converter(_plc_data_item['value'], data_type)

            if _plc_data_item['address'] == address:
                plc_data[_data_index]['value'] = _value
            else:
                raise Exception("Data out of sync.")
            # end if
            _data_index = _data_index + 1
        # end for
        return plc_data
    # end update_data_with_new_values
# end PLCSheetData

class MainSheetData(SheetDataBase):

    config_data = None
    opc_topic = None
    slc_file = None

    def get_config_cell(self):
        return self.search_sheet(CONFIG_FLAG)
    # end get_config_cell

    def get_opc_topic(self):

        if self.opc_topic is None:
            # gets the topic value
            _topic = self.get_search_offset_value(OPC_TOPIC_FLAG, 0, OPC_TOPIC_OFFSET)

            if _topic is not None and _topic != "":
                self.opc_topic = _topic
            else:
                raise Exception("Could not locate valid OPC topic in MAIN PROGRAM sheet.")
        # end if

        return self.opc_topic
    # end get_opc_topic

    def get_pc5_file(self):

        if self.slc_file is None:
            _slc_file = self.get_search_offset_value(SLC_FILE_FLAG, 0, SLC_FILE_OFFSET)

            if _slc_file is not None and _slc_file != "":
                self.slc_file = _slc_file
            else:
                raise Exception("Could not locate valid SLC in MAIN PROGRAM sheet")
            # end if
        # end if

        return self.slc_file
    # end get_slc_file

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
