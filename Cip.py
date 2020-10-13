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
            self.plc_info = c.plc_info
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
            return c.write_tags(*tag_list)
        else:
            return c.read_tags(*tag_list)
        

# ----------------- initial testing code ---------------------------
# def output(thread_name, message, sdtlock):
#     sdtlock.acquire()
#     print("%s: %s" % (thread_name, message))
#     sdtlock.release()

# def get_tags_from_plc(cip_path, stdlock):
#     try:
#         output("main", "connecting to controller %s" % cip_path, sdtlock)
#         c = LogixDriver(cip_path)
#         output("main", "PLC_INFO %s" % c.info, sdtlock)
#     except Exception as ex:
#         output("main", "ERROR: %s" % ex, sdtlock)
#     # end try
#     if c.connected:
#         _tags = c._tags
#         c.close()
#         output("main", "complete.", sdtlock)
#         return _tags
#     else:
#         return None

# def worker(thread_id, cip_manager, cip_path, tags, sdtlock):

#     read_tags = ['CycleTMR9.ACC','CycleTMR10.ACC']

#     # wait for a valid connection
#     output("THREAD-%s" % thread_id, "waiting for a connection...", sdtlock)
#     cip_manager.wait_for_connection()
#     output("THREAD-%s" % thread_id, "connection available.", sdtlock)

#     try:
#         controller = LogixDriver(path=cip_path, init_tags=False, init_program_tags=False)
#         controller._tags = tags
#     except Exception as ex:
#         output("THREAD-%s" % thread_id, "ERROR: %s" % ex, sdtlock)
    
#     if controller.connected:

#         output("THREAD-%s" % thread_id, "reading tags.", sdtlock)
#         ret = controller.read(*read_tags)

#         # remember to remove the connection after data transfer
#         # has been completed.
#         cip_manager.remove_connection()

#         for i in ret:
#             # read i.TagName, i.Value, i.Status
#             output("THREAD-%s" % thread_id, i, sdtlock)
#         # end for
#         output("THREAD-%s" % thread_id, "reading complete.", sdtlock)
#         controller.close()

# if __name__ == '__main__':

#     manager = CIPConnectionManager.Manager()

#     sdtlock = threading.Lock()

#     # connect once to get the tag structure
#     cip_path = "192.168.59.211/0"
#     tags = get_tags_from_plc(cip_path, sdtlock)
#     if tags is not None:
#         for i in range(0, 10):
#             t = threading.Thread(target=worker, args=(i, manager, cip_path, tags, sdtlock))
#             t.start()
#         # end for
#     # end if

# # end main
