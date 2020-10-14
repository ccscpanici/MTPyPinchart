from pycomm3 import LogixDriver
import threading

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

    def __plc_operation__(self, tag_list, write=False):
        _tag_structure = self.get_plc_tags()
        c = LogixDriver(path=self.cip_path, init_tags=False, init_program_tags=False)
        c._tags = _tag_structure
        if write:
            results = c.write(*tag_list)
        else:
            results = c.read(*tag_list)
        
        return results
        