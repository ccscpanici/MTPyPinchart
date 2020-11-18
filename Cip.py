from pycomm3 import LogixDriver
import threading
import utils
import time

class ConnManager(object):
    """ 
     Connection manager keeps the concurrent
     connections down to a minimum. This is to
     protect the controller from maxing out their
     allowed number of CIP connections. Remember
     To call the remove_connection() from the 
     manager after CIP operation(s) are complete.
    """
    def __init__(self, max=2):
        self.max = max
        self.connections = 0

    def add_connection(self):
        if self.connections < self.max:
            self.connections += 1
        else:
            raise Exception("Max connection count exceeded.")
    
    def remove_connection(self):
        if self.connections > 0:
            self.connections -= 1
        else:
            raise Exception("Invalid connection total. Something is wrong.")

    def wait_for_connection(self):
        while self.connections == self.max:
            time.sleep(0.25)
        # end while

        # after if gets a connection, add it
        self.add_connection()

class LogixController(object):
    def __init__(self, ip_address_string, slot_number, tag_structure=None):
        self.cip_path = "%s/%s" % (ip_address_string, slot_number)
        self.plc_info = None
        self.tag_structure = tag_structure
        self.tag_validation_error = {}
        self.tag_validation_warnings = []
    def get_plc_tags(self):
        """
        This function reads the tags inside the controller
        and returns them. This only needs to be done once.
        """
        if self.tag_structure is None:
            c = LogixDriver(self.cip_path)
            self.plc_info = c.info
            self.tag_structure = c._tags
            return self.tag_structure
        else:
            return self.tag_structure

    def read_tags(self, tag_list):
        return self.__plc_operation__(tag_list, False)

    def write_tags(self, tag_list):
        return self.__plc_operation__(tag_list, True)

    def validate_data_types(self, tags_list, data_type_string):

        # controller data types that exist in tag list
        _data_types = []

        # clears out the errors and warnings on the first run
        self.tag_validation_warnings = []
        self.tag_validation_error = {}

        if data_type_string is None:
            self.tag_validation_error = {
                "message": "function can not pass None in for data type"
            }
            return False
        
        # make sure all of the tags in the list are of the same data type
        # and are compatible with the main sheet data type
        for i in tags_list:

            # get the data type from the controller tag list
            _path_string = utils.get_tag_indexing_string(i)
            _controller_data_type = eval("self.get_plc_tags()" + _path_string)
            
            # append the controller data type to the list if
            # it is not already in there.
            if _controller_data_type not in _data_types: 
                _data_types.append(_controller_data_type)

            # figuring out the sheet data type
            _sheet_type = None
            _dint_type_strings = ["DINT", "DINT DATA", "DINT-32 DATA", "DINT-32"]
            _real_type_strings = ["REAL", "FLOAT", "FLOAT DATA"]
            _string_type_strings = ["STRING", "STRING DATA"]
            _int_type_strings = ["INT", "INT DATA", "INT-16 DATA"]
            if data_type_string in _dint_type_strings:
                _sheet_type = "DINT"
            elif data_type_string in _real_type_strings:
                _sheet_type = "REAL"
            elif data_type_string in _string_type_strings:
                _sheet_type = "STRING"
            elif data_type_string in _int_type_strings:
                _sheet_type = "INT"
            else:
                self.tag_validation_error = {
                    "message":"unsupported data type",
                    "data type": data_type_string,
                    "tag": i
                }
                return False
            
            # if the tag in the controller does not match the tag in the sheet
            # there are a couple of things we can do.
            # if the data types are not equal
            # error - if one is a string and the other is not
            # error - if the sheet is a dint and the controller is an int or sint
            # error - if the sheet is a float and the controller is not
            # warn - if the sheet is a dint and the controller is a float
            # warn - if the sheet is an int and the controller is a dint
            if _sheet_type == "STRING" and _controller_data_type != "STRING" or \
                _sheet_type != "STRING" and _controller_data_type == "STRING":
                self.tag_validation_error = {
                    "message":"sheet is a string type and controller type is not",
                    "tag":i
                }
                return False
            elif _sheet_type == "DINT" and _controller_data_type == "INT":
                self.tag_validation_error = {
                    "message":"sheet is a dint and controller data type is an INT",
                    "tag":i
                }
                return False
            elif _sheet_type == "REAL" and _controller_data_type == "DINT":
                self.tag_validation_error = {
                    "message":"sheet is a FLOAT and the controller type is a DINT",
                    "tag": i
                }
                return False
            elif _sheet_type == "DINT" and _controller_data_type == "REAL":
                self.tag_validation_warnings.append(
                    {
                        "message":"Data mismatch warning. Sheet data is DINT and controller type is REAL",
                        "tag":i
                    }
                )
            elif _sheet_type == "INT" and _controller_data_type == "DINT":
                self.tag_validation_warnings.append(
                    {
                        "message":"Data mismatch warning. Sheet data is INT and controller type is DINT",
                        "tag":i
                    }
                )

        if len(_data_types) > 1:
            self.tag_validation_error = {
                "message":"tags in data chunk are not of the same data type",
                "sheet tag types": _data_types
            }
            return False

        # if it gets here, return True
        return True

    def __plc_operation__(self, tag_list, write=False):
        _tag_structure = self.get_plc_tags()
        c = LogixDriver(path=self.cip_path, init_tags=False, init_program_tags=False)
        c._tags = _tag_structure

        if write:
            results = c.write(*tag_list)
        else:
            results = c.read(*tag_list)

        # close the connection after the
        # operation.
        c.close()

        return results
        