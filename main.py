import os
import filereader
import serialize
import threading
import worker

PLC_OPERATION_DOWNLOAD = 0
PLC_OPERATION_UPLOAD = 1

OPC_SERVER = 'RSLinx OPC Server'

if __name__ == '__main__':

    # gets the full path of the excel file
    excel_file = os.getcwd() + "//RO Pinchart.xlsm"

    # gets all of the data from the workbook
    excel_dict = filereader.get_xl_data(excel_file)

    # sheets to process list. These will be given to the sheet
    # processor to process on different threads
    sheets_to_process = []

    # setup the thread locks.
    # excel lock
    elock = threading.Lock()
    # opc lock
    olock = threading.Lock()

    _operation = "upload"

    if _operation.lower() == "upload":
        program_operation = PLC_OPERATION_UPLOAD
    elif _operation.lower() == "download":
        program_operation = PLC_OPERATION_DOWNLOAD
    else:
        raise Exception("Invalid Input for Operation: %s" % _operation)
    # end if

    for sheet_name in excel_dict:

        # gets the sheet data
        sheet_data = excel_dict[sheet_name]

        if sheet_name.lower() == 'main program':

            # serialized the main sheet into a class
            main_sheet = serialize.MainSheetData(sheet_data)

            # gets the configutation dictionary for the workbook. If it doesn't
            # find the configuration, it uses default values.
            config_data = main_sheet.get_config_data()

        elif sheet_name.lower()[0:8] == 'pinchart' or sheet_name.lower()[0:8] == 'plc data':

            # add the sheet name to process on a different
            # thread
            sheets_to_process.append(
                {
                    'elock':elock,
                    'opclock':olock,
                    'sheet_name':sheet_name,
                    'sheet_dict':sheet_data,
                    'config_data':config_data,
                    'operation':program_operation
                }
            )
        # end if
    # end for

    # loop through the sheets that need to get processed
    thread_id = 0
    for a_sheet in sheets_to_process:

        # increment thread id
        thread_id = thread_id + 1
        a_sheet['thread_id'] = thread_id

        #elock, opclock, sheet_name, sheet_dict, config_data, operation
        
        # process the sheet
        worker.process_sheet(thread_id, a_sheet['elock'], a_sheet['opclock'], 
                            a_sheet['sheet_name'], a_sheet['sheet_dict'], 
                            a_sheet['config_data'], a_sheet['operation'])

    # end worker
# end main