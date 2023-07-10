from pycomm3 import LogixDriver, RequestError
from pycomm3.logger import configure_default_logger, LOG_VERBOSE
import threading
import utils
import time
import utils
from typing import List, Tuple, Optional, Union
from settings import DEBUG_MODE


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
    def __init__(self, ip_address_string, slot_number):
        self.cip_path = "%s/%s" % (ip_address_string, slot_number)
        self._tag_database = None
        if DEBUG_MODE:
            configure_default_logger(filename='pycomm3.log')

    def get_tag_database(self):
        driver = LogixDriver(self.cip_path, init_tags=True)
        driver.open()
        _tag_database = driver._tags
        driver.close()
        return _tag_database

    def set_tag_database(self, tag_database):
        self._tag_database = tag_database
    
    def read_tag(self, tag, tag_database):
        
        driver = LogixDriver(self.cip_path, init_tags=False)
        driver._tags = tag_database
        driver.open()
        result = driver.read(tag)
        driver.close()
        return result
        
    def read_tags(self, tag_list, tag_database):
        return self.__plc_operation__(tag_list, tag_database, False)

    def write_tags(self, tag_list, tag_database):
        return self.__plc_operation__(tag_list, tag_database, True)

    def write_tag(self, tag, value, tag_database):

        driver = LogixDriver(self.cip_path, init_tags=False)
        driver._tags = tag_database
        driver.open()
        result = driver.write((tag, value))
        driver.close()
        return result

    def __plc_operation__(self, tag_list, tag_database, write=False):

        # initiates the driver
        driver = LogixDriver(self.cip_path, init_tags=False)
        
        # sets the tags to the tag database that is in
        # memory
        driver._tags = tag_database

        # open the connection
        driver.open()

        if write:
            results = driver.write(*tag_list)
        else:
            results = driver.read(*tag_list)

        # close the driver at the end of the operation
        driver.close()

        return results