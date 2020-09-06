import os
import filereader
import serialize
import threading
import worker
import getopt
import sys

PLC_OPERATION_DOWNLOAD = 0
PLC_OPERATION_UPLOAD = 1

OPC_SERVER = 'RSLinx OPC Server'

if __name__ == '__main__':

    # temporary user arguments
    temp_user_args = ['-f', os.getcwd() + "/" + "RO Pinchart.xlsm", '-o', 'DOWNLOAD', '-s', 'PINCHART-PROC, PINCHART-CIP']

    # user arguments
    #user_args = sys.argv[1:]
    #print("user_args = %s" % user_args)

    # system arguments possibilities
    # f - filename
    # o - operation
    # s - sheet names - if this doesn't exist then we will assume all the sheets
    # print the arguments
    args = getopt.getopt(temp_user_args, "f:o:s:")

    # loop through the arguments and set the variables
    arg_excel_file = False
    arg_operation = False
    arg_sheets = False
    process_all_sheets = False
    for switch, value in args[0]:
        if switch == '-f':
            arg_excel_file = value
        elif switch == '-o':
            arg_operation = value
        elif switch == '-s':

            # gets the user called sheets to process
            arg_sheets = value.replace(" ", "").strip().split(",")

            # lower case the sheet
            arg_sheets = [i.lower() for i in arg_sheets]
        else:
            raise Exception("Invalid switch %s" % switch)
            sys.exit(1)
        # end else
    # end for

    if not arg_excel_file or not arg_operation:
        raise Exception("Must specify excel file and operation")
        sys.exit(1)
    # end if
    if not arg_sheets:
        arg_sheets = []
        process_all_sheets = True
    # end if
    
    # gets all of the data from the workbook
    excel_dict = filereader.get_xl_data(arg_excel_file)

    # sheets to process list. These will be given to the sheet
    # processor to process on different threads
    sheets_to_process = []
    
    # setup the thread locks.
    # excel lock
    elock = threading.Lock()

    # opc lock
    olock = threading.Lock()

    if arg_operation.lower() == "upload":
        program_operation = PLC_OPERATION_UPLOAD
    elif arg_operation.lower() == "download":
        program_operation = PLC_OPERATION_DOWNLOAD
    else:
        raise Exception("Invalid Input for Operation: %s" % arg_operation)
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

        elif process_all_sheets or (sheet_name.lower() in arg_sheets):

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