import OpenOPC
import serialize
import utils
import excel_interface
import pywintypes
import pc5_interface
from main import PLC_OPERATION_DOWNLOAD, PLC_OPERATION_UPLOAD, PLC_OPERATION_EXPORT, PLC_OPERATION_IMPORT, OPC_SERVER

pywintypes.datetime = pywintypes.TimeType

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
def process_sheet(elock, opclock, slock, pc5lock, main_sheet_object, \
                    sheet_name, sheet_dict, config_data, operation, \
                    thread_id, full_file_path):
    
    utils.output(thread_id, "worker", "process_sheet", "PROCESSING SHEET %s" % sheet_name, slock)

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
    for plc_data_column in _plc_data_structure:

        if not running:
            utils.output(thread_id, "worker", "process_sheet", "TERMINATED UNEXPECTIDLY", slock)
            break
        # end if

        if operation == PLC_OPERATION_UPLOAD or operation == PLC_OPERATION_DOWNLOAD:

            # get the topic from the main sheet
            _topic = main_sheet_object.get_opc_topic()

            # set the serializer class topic
            sheet_object.set_opc_topic(_topic)

            # get the plc data for the column
            plc_data_column['plc_data'] = sheet_object.get_plc_data_for_column(plc_data_column)

            # waits until no other thread is using
            # the opc object.
            opclock.acquire()
            opc = OpenOPC.client()
            opc.connect(OPC_SERVER)

            #utils.output(thread_id, "worker", "process_sheet", "CONNECTED TO OPC SERVER.", slock)
        # end if

        if operation == PLC_OPERATION_IMPORT or operation == PLC_OPERATION_EXPORT:

            # get the PLC file
            _pc5_file = main_sheet_object.get_pc5_file()

            # lock the file
            pc5lock.aquire()

            # read the pc5 file
            pc5 = pc5_interface.PC5_File(_pc5_file)

            # release the pc5 file lock

        # end if

        if operation == PLC_OPERATION_DOWNLOAD:

            utils.output(thread_id, "worker", "process_sheet", "%s-GETTING PLC ADDRESSES AND VALUES..." % sheet_name, slock)
            data_tuples = sheet_object.get_address_value_list(plc_data_column['plc_data'], plc_data_column['data']['type'])

            utils.output(thread_id, "worker", "process_sheet", "%s-GETTING DOWNLOADING DATA CHUNK..." % sheet_name, slock)

            try:
                _return_data = opc.write(data_tuples)
            except Exception as ex:
                utils.output(thread_id, "worker", "process_sheet", "UNHANDLED ERROR: %s" % ex, slock)
                running = False
                _return_data = None
            # end try

            opclock.release()

            # process the return data
            # on the download - we probably don't have to, maybe print
            # any errors?

            if _return_data:
                # process the data from the opc function. Maybe raise an error,
                # if too many errors were returned?
                pass
            # end _return_data

        elif operation == PLC_OPERATION_UPLOAD:
            
            # gets the address list for upload
            utils.output(thread_id, "worker", "process_sheet", "%s-GETTING PLC ADDRESSES..." % sheet_name, slock)
            addresses = sheet_object.get_address_list(plc_data_column['plc_data'])

            utils.output(thread_id, "worker", "process_sheet", "%s-GETTING UPLOADING DATA CHUNK..." % sheet_name, slock)
            
            try:
                _return_data = opc.read(addresses)
            except Exception as ex:
                _return_data = None
                utils.output(thread_id, "worker", "process_sheet", "UNHANDLED ERROR: %s" % ex, slock)
                running = False
            # end try
            
            opclock.release()
            
            if _return_data:
                # process the return data, if there are a bunch
                # of errors, maybe kick it out?
                plc_data_column['plc_data'] = sheet_object.update_data_with_new_values(plc_data_column['data']['type'], plc_data_column['plc_data'], _return_data)

                # lock the execl
                elock.acquire()

                # create the win32com excel interface class
                utils.output(thread_id, "worker", "process_sheet", "%s-GETTING WORKBOOK..." % sheet_name, slock)
                _excel = excel_interface.Interface(full_file_path, sheet_name)

                # hand the interface class the data that needs 
                # to be updated
                utils.output(thread_id, "worker", "process_sheet", "%s-UPDATING WORKSHEET WITH DATA CHUNK..." % sheet_name, slock)
                _excel.update_sheet(plc_data_column, config_data)

                # after it is all updated, release the lock
                elock.release()
            # end if

        elif operation == PLC_OPERATION_IMPORT:

            # import the data from the file

            # release the lock
            pc5lock.release()

        elif operation == PLC_OPERATION_EXPORT:

            # export the operation from the sheet
            # to the PC5 file
            pass

            # release the lock
            pc5lock.release()

        else:
            elock.release()
            opclock.release()
            pc5lock.release()
            elock.release()
            raise Exception("Invalid PLC Operation")
        # end if
        
    # end for
# end if