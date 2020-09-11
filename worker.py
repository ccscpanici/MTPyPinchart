import OpenOPC
import serialize
import utils
from main import PLC_OPERATION_DOWNLOAD, PLC_OPERATION_UPLOAD, PLC_OPERATION_EXPORT, PLC_OPERATION_IMPORT, OPC_SERVER

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

    utils.output(thread_id, "worker", "process_sheet", "GETTING PLC DATA...", slock)

    # get the PLC data
    _plc_data = sheet_object.get_plc_data(_data_row_offset)

    # loop through different chunks of data
    for plc_data_column in _plc_data:

        if operation == PLC_OPERATION_UPLOAD or operation == PLC_OPERATION_DOWNLOAD:

            # get the topic from the main sheet
            _topic = main_sheet_object.get_opc_topic()

            opclock.acquire()
            #opc = OpenOPC.client()
            #opc.connect(OPC_SERVER)

            #utils.output(thread_id, "worker", "process_sheet", "CONNECTED TO OPC SERVER.", slock)
        # end if

        if operation == PLC_OPERATION_IMPORT or operation == PLC_OPERATION_EXPORT:

            # get the PLC file
            _pc5_file = main_sheet_object.get_pc5_file()

            # lock the file
            pc5lock.aquire()

        # end if

        if operation == PLC_OPERATION_DOWNLOAD:

            data_tuples = sheet_object.get_address_value_list(plc_data_column['plc_data'], plc_data_column['data']['type'], _topic)
            #return_data = opc.write(data_tuples)
            opclock.release()

            # process the return data
            # on the download - we probably don't have to, maybe print
            # any errors?

        elif operation == PLC_OPERATION_UPLOAD:
            
            # gets the address list for upload
            addresses = sheet_object.get_address_list(plc_data_column['plc_data'], _topic)

            #return_data = opc.read(addresses)
            opclock.release()
            
            # process the return data

            # lock the execl
            elock.acquire()

            # create the win32com excel interface class

            # hand the interface class the data that needs 
            # to be updated

            # save the workbook

            # after it is all updated, release the lock
            elock.release()

        elif operation == PLC_OPERATION_IMPORT:

            # import the data from the file
            pass

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