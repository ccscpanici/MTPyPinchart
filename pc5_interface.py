import utils
from utils import output
# this module is for SLC file operations
# import and export
class PC5_File(object):

    def __init__(self, filename):
        
        self.__filename__ = filename
        self.__data_table__ = {}
        self.__reference_table__ = {}
        self.__raw_file__ = None
        self.__output_file__ = None

        # loads the raw file into memory
        self.__load_raw_file__()

        # loads the data table on instantiation
        self.__load_data_tables__()

    # end __init__

    def __load_raw_file__(self):
        # this sucks in the file and closes
        # it.
        #output(self.__class__.__name__, "__load_raw_file__", "Loading raw SLC file (.PC5)...")
        file_obj = open(self.__filename__, 'r')
        self.__raw_file__ = []
        for line in file_obj:
            self.__raw_file__.append(line)
        # end for

        # close the file
        file_obj.close()
        #output(self.__class__.__name__, "__load_raw_file__", "Loading raw SLC file (.PC5)...Complete.")
    # end __load_raw_file__

    def __load_data_tables__(self):

        wanted_files = ['N', 'F', 'B', 'S']
        #output(self.__class__.__name__, "__load_data_tables__", "Parsing SLC Data Tables...%s" % wanted_files)

        # this function loads all of the data tables
        # into memory. This should one dictionary where the
        # key is the address and the value is the data
        # table value
        # start by looping through file lines
        in_data_file = False

        # this variable keeps track of the line index
        line_index = 0
        for _line in self.__raw_file__:

            if _line[0:4] == 'DATA' or in_data_file:

                if _line[0:4] == 'DATA':
                    
                    # parses the data file name
                    _data_file = _line[5:]
                    _data_file = _data_file.replace("\n", "")

                    if _data_file[0:1] in wanted_files:
                        # Sets the in file bit
                        in_data_file = True
                    # end if

                if _line == '\n' or _line == '':
                    in_data_file = False
                # end if

                if _line != '\n' and _line != '' and _line[0:4] != 'DATA':

                    # take off the end carraige return
                    _line = _line.replace("\n", "")

                    # gets the first address on the line
                    _first_address = _line.split("%")[1].strip()

                    # gets the data
                    _data = utils.clean_int_list(_line.split("%")[2].split(" "))

                    # gets the address file (prefix)
                    _first_address_file = _first_address.split(":")[0]

                    # gets the index of the file
                    _first_address_index = int(_first_address.split(":")[1])

                    # keeps track of the address index with position index
                    # variable
                    _pos_index = 0

                    # loops through the data line and assigns the point to
                    # an address and adds it to the dictionary
                    for _point in _data:
                        _address = _first_address_file + ":" + str((_first_address_index + _pos_index))
                        self.__data_table__[_address] = _point
                        self.__reference_table__[_address] = {"line_index":line_index, "value_index": _pos_index, "initial_value":_point}
                        _pos_index = _pos_index + 1
                    # end for
                # end if            
            # end if

            # add one to the line index
            line_index = line_index + 1
        # end looping through
        
        #output(self.__class__.__name__, "__load_data_tables__", "Parsing SLC Data Tables...Complete")
    # end __load_data__()

    def get_plc_value(self, address):
        # if the address contains the topic name, we strip that
        # off and get just the address.
        if address.__contains__("]"):
            _address_withou_topic = address.split("]")[1]
            if self.__data_table__.has_key(_address_withou_topic):
                return self.__data_table__[address.split("]")[1]]
            else:
                #output(self.__class__.__name__, "get_plc_value", "ERROR: Address not found in program: %s" % _address_withou_topic)
                return False
        else:
            return self.__data_table__[address]
        # end if
    # end get_plc_value

    def get_plc_values(self, plc_data):
        # this method gets all of the plc_data
        # and adds it to the dictionary. Then,
        # returns the dictionary so the excel sheet
        # can be updated
        for _column in plc_data.keys():
            _column_data = plc_data[_column]
            for _row in _column_data:
                _row['value'] = self.get_plc_value(_row['address'])
            # end for
        # end for
        return plc_data
    # end get_plc_values

    def __update_line__(self, line, address, value):

        # gets the first address in the line
        _first_address = line.split("%")[1].strip()

        # gets the first address number
        _first_address_number = int(_first_address.split(":")[1])

        # gets the input address number
        _input_address_number = int(address.split(":")[1])

        # gets the difference of the numbers
        _difference = _input_address_number - _first_address_number

        # gets the last index of the percent character
        _start_values = line.rindex("%")

        # create the new value string that is 6 characters long
        _new_value_string = str(value).rjust(7)

        # calculate the last position
        _last_digit_pos = (_start_values + (7 * (_difference + 1)))

        # calculate the start position
        _first_digit_pos = _last_digit_pos - 6

        _pre_value_string = line[0:_first_digit_pos]
        _post_value_string = line[_last_digit_pos + 1:]

        # formulates the new string
        new_string = _pre_value_string + _new_value_string + _post_value_string

        return new_string
    # end __update_line__()

    def update_data_tables(self, plc_data_from_excel):

        #output(self.__class__.__name__, "update_data_tables", "Updating SLC file with values in pin chart sheet...")
        
        # this method updates the values that are in the data
        # table.
        _data = plc_data_from_excel

        # copy the raw file into the output variable
        self.__output_file__ = self.__raw_file__

        for _column in _data.keys():

            # gets the column data 
            _column_data = _data[_column]

            # loops through the items in the column data
            for _column_data_item in _column_data:

                # gets the plc address
                _address = _column_data_item['address']
                _new_value = _column_data_item['value']

                if _address.__contains__("]"):
                    _address = _address.split("]")[1]
                # end if

                # get the file reference for the address
                _reference = self.__reference_table__[_address]

                # gets the reference line number
                _reference_line_number = _reference['line_index']

                # get the reference line
                _reference_line = self.__raw_file__[_reference_line_number]

                # update the line with the new value
                _updated_line = self.__update_line__(_reference_line, _address, _new_value)

                # replace the line with the updated one
                self.__output_file__[_reference_line_number] = _updated_line

            # end looping through column data
        # end looping through data columns
        
        #output(self.__class__.__name__, "update_data_tables", "Updating SLC file with values in pin chart sheet...Complete")
    # end update_data_tables

    def save_file(self):

        #output(self.__class__.__name__, "save_file", "Writing lines of new PC5 file located at: %s...." % self.__filename__)

        # writes the text file
        try:
            _file = open(self.__filename__, 'w')

            for _line in self.__output_file__:
                _file.write(_line)
            # end for

            # close the file
            _file.close()
            
            #output(self.__class__.__name__, "save_file", "Writing lines of new PC5 file located at: %s....Complete." % self.__filename__)
            return True
        except IOError as ex:
            #output(self.__class__.__name__, "save_file", "IO ERROR: Permission Denied - File Locked. Export not successful. %s" % ex)
            return False
        # end try
# end SLC_File

if __name__ == '__main__':

    _file = "Pinchart.PC5"
    pc5 = PC5_File(_file)

    print(pc5)