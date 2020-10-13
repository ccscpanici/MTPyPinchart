import OpenOPC
import serialize
import utils
import excel_interface
import pc5_interface
import time
import random
from main import PLC_OPERATION_DOWNLOAD, PLC_OPERATION_UPLOAD, PLC_OPERATION_EXPORT, PLC_OPERATION_IMPORT, OPC_SERVER
import Cip

def threads_running(thread_list):
    count = 0
    for i in thread_list:
        if i.is_alive():
            count = count + 1
        # end if
    # end for
    return count
# end threads_running

# # WORKER MODULE IS FOR PROCESSING A SHEETs
# ON A SEPARATE THREAD
def process_sheet(**kwargs):

    # working out the kwargs
    thread_id = kwargs['thread_id']
    sheet_name = kwargs['sheet_name']
    sheet_dict = kwargs['sheet_dict']
    config_data = kwargs['config_data']
    operation = kwargs['operation']
    slock = kwargs['slock']

    utils.output(thread_id, "worker", "process_sheet", "PROCESSING SHEET %s" % sheet_name, slock)
    
    # sleep a random time between 0 and 3 seconds.
    # this should give the system a chance to get the proper
    # locks in
    _sleep_time = random.uniform(0,3)
    time.sleep(_sleep_time)
    
    # get the sheet serializer
    sheet_object = serialize.PLCSheetData(sheet_dict, config_data, thread_id)

    # gets the data row offset
    _data_row_offset = config_data['DATA ROW OFFSET']

    utils.output(thread_id, "worker", "process_sheet", "%s-GETTING PLC SCHEMA FROM SHEET..." % sheet_name, slock)

    # gets the PLC data schema
    _plc_data_structure = sheet_object.get_plc_data_structure()

    # sets a volitile bit - this is used to exit the loop
    # on an error
    running = True

    # loop through different chunks of data
    # gets the number of chunks that need to be processed
    _data_chunks = _plc_data_structure.__len__()
    _data_chunk_index = 1

    # this keeps track of how many errors there were
    operation_errors = 0

    for plc_data_column in _plc_data_structure:

        if not running:
            utils.output(thread_id, "worker", "process_sheet", "TERMINATED UNEXPECTIDLY", slock)
            break
        # end if

        if operation == PLC_OPERATION_UPLOAD or operation == PLC_OPERATION_DOWNLOAD:
                
            # get the plc data for the column
            plc_data_column['plc_data'] = sheet_object.get_plc_data_for_column(plc_data_column)

            #utils.output(thread_id, "worker", "process_sheet", "CONNECTED TO OPC SERVER.", slock)
        # end if

        if operation == PLC_OPERATION_IMPORT or operation == PLC_OPERATION_EXPORT:
            
            pc5_file = kwargs['pc5_file']
            pc5_lock = kwargs['pc5_lock']

            # lock the file
            pc5_lock.acquire()

            # read the pc5 file
            pc5 = pc5_interface.PC5_File(pc5_file)

            # release the pc5 file lock
            pc5_lock.release()

            # get the plc data for the column. We don't need to set an opc topic
            plc_data_column['plc_data'] = sheet_object.get_plc_data_for_column(plc_data_column)

            # make sure that the file is a PLC-5 or SLC type address pattern.
            # if it is not, then we need to raise an error
            if plc_data_column['plc_data'][0]['address'].__contains__(":") == False:
                raise Exception("OPERATION IS NOT FOR A CONTROL LOGIX DATA STRUCTURE.")
            # end if

        # end if

        if operation == PLC_OPERATION_UPLOAD or operation == PLC_OPERATION_DOWNLOAD:

            cip_manager = kwargs['cip_manager']
            plc_tags = kwargs['plc_tags']
            ip_address = kwargs['ip_address']
            slot_number = kwargs['slot_number']
            controller = Cip.LogixController(ip_address, slot_number, plc_tags)

        if operation == PLC_OPERATION_UPLOAD or operation == PLC_OPERATION_IMPORT:
            elock = kwargs['elock']
            excel_file_path = kwargs['excel_file_path']

        if operation == PLC_OPERATION_DOWNLOAD:

            #utils.output(thread_id, "worker", "process_sheet", "%s-GETTING PLC ADDRESSES AND VALUES..." % sheet_name, slock)
            data_tuples = sheet_object.get_address_value_list(plc_data_column['plc_data'], plc_data_column['data']['type'])

            # wait for a  CIP connection from the manager, 
            # this is a blocking call so it won't continue until
            # it has one
            cip_manager.wait_for_connection()

            # print the downloading message
            utils.output(thread_id, "worker", "process_sheet", "%s--DOWNLOADING-- Data Chunk [%s] of [%s]" % (sheet_name, _data_chunk_index, _data_chunks), slock)

            # write the controller tags
            response = controller.write_tags(data_tuples)

            # remove the connection from the manager that way
            # another thread can access it.
            cip_manager.remove_connection()

            if not all(response):
                # there were errors during transmission
                for i in response:
                    if i['error']:
                        utils.output(thread_id, "worker", "process_sheet", "Tag Error: %s, \tValue: %s" % (i['tag'], i['value']), slock)
                    # end if
                # end for
            # end if           

        elif operation == PLC_OPERATION_UPLOAD:
            
            # gets the address list
            addresses = sheet_object.get_address_list(plc_data_column['plc_data'])
            
            # grab a cip connection
            cip_manager.wait_for_connection()

            utils.output(thread_id, "worker", "process_sheet", "%s--UPLOADING-- DATA CHUNK.[%s] of [%s]" % (sheet_name, _data_chunk_index, _data_chunks), slock)

            # upload the data
            response = controller.read_tags(addresses)
                        
            if response:
                # process the return data, if there are a bunch
                # of errors, maybe kick it out?
                plc_data_column['plc_data'] = sheet_object.update_data_with_new_values(plc_data_column['data']['type'], plc_data_column['plc_data'], response)

                # lock the execl
                elock.acquire()

                # create the win32com excel interface class
                #utils.output(thread_id, "worker", "process_sheet", "%s-GETTING WORKBOOK..." % sheet_name, slock)
                _excel = excel_interface.Interface(excel_file_path, sheet_name)

                # hand the interface class the data that needs 
                # to be updated
                utils.output(thread_id, "worker", "process_sheet", "%s--UPDATING WORKSHEET-- WITH DATA CHUNK.[%s] of [%s]" % (sheet_name, _data_chunk_index, _data_chunks), slock)
                _excel.update_sheet(plc_data_column, config_data)

                # after it is all updated, release the lock
                elock.release()
            # end if

        elif operation == PLC_OPERATION_IMPORT:

            # import the data from the file.
            # updates the dictionary
            plc_data_column['plc_data'] = pc5.get_plc_values(plc_data_column)

            # lock the execl
            elock.acquire()

            # create the win32com excel interface class
            #utils.output(thread_id, "worker", "process_sheet", "%s-GETTING WORKBOOK..." % sheet_name, slock)
            _excel = excel_interface.Interface(excel_file_path, sheet_name)

            # hand the interface class the data that needs 
            # to be updated
            utils.output(thread_id, "worker", "process_sheet", "%s--UPDATING WORKSHEET-- WITH DATA CHUNK.[%s] of [%s]" % (sheet_name, _data_chunk_index, _data_chunks), slock)
            _excel.update_sheet(plc_data_column, config_data)

            # after it is all updated, release the lock
            elock.release()
            
        elif operation == PLC_OPERATION_EXPORT:

            # lock the pc5 file so we can write to
            # it.
            pc5_lock.acquire()

            # export the operation from the sheet
            # to the PC5 file
            pc5.update_data_tables(plc_data_column)

            # release the lock
            pc5_lock.release()

        else:
            elock.release()
            pc5_lock.release()
            elock.release()
            raise Exception("Invalid PLC Operation")
        # end if
        
        _data_chunk_index = _data_chunk_index + 1
    # end for

    # if there were errors - print this at the end so the user knows.
    if operation_errors > 0:
        utils.output(thread_id, "worker", "process_sheet", "%s:***IMPORTANT****OPERATION COMPLETED WITH %s ERRORS****" % (sheet_name, operation_errors), slock)
    # end if

# end if