import OpenOPC
import serialize
from main import PLC_OPERATION_DOWNLOAD, PLC_OPERATION_UPLOAD, OPC_SERVER

# # WORKER MODULE IS FOR PROCESSING A SHEETs
# ON A SEPARATE THREAD
def process_sheet(thread_id, elock, opclock, sheet_name, sheet_dict, config_data, operation):

    # get the sheet serializer
    sheet_object = serialize.PLCSheetData(sheet_dict, config_data)

    # gets the data row offset
    _data_row_offset = config_data['DATA ROW OFFSET']

    # get the PLC data
    _plc_data = sheet_object.get_plc_data(_data_row_offset)

    # loop through different chunks of data
    for plc_data_column in _plc_data:
        opc = None

        if operation == PLC_OPERATION_UPLOAD or operation == PLC_OPERATION_DOWNLOAD:
            
            opclock.acquire()
            #opc = OpenOPC.client()
            #opc.connect(OPC_SERVER)

        if operation == PLC_OPERATION_DOWNLOAD:

            data_tuples = sheet_object.get_address_value_list(plc_data_column['plc_data'])

        elif operation == PLC_OPERATION_UPLOAD:
            
            # gets the address list for upload
            addresses = sheet_object.get_address_list(plc_data_column['plc_data'])

        else:
            raise Exception("Invalid PLC Operation")
        # end if

        # release the lock
        opclock.release()
  
    # end for

# end if