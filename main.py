import os
import filereader
import serialize
import threading
import worker
import getopt
import sys
import utils
import time
import excel_interface
import pywintypes

pywintypes.datetime = pywintypes.TimeType

PLC_OPERATION_DOWNLOAD = 0
PLC_OPERATION_UPLOAD = 1
PLC_OPERATION_IMPORT = 2
PLC_OPERATION_EXPORT = 3

# USE THIS BOOL TO TURN MULTITHREADED
# ON AND OFF
MULTITHREAD = True

OPC_SERVER = 'RSLinx OPC Server'

if __name__ == '__main__':

    # temporary user arguments
    #temp_user_args = ['-f', os.getcwd() + "/" + "RO Pinchart.xlsm", '-o', 'DOWNLOAD', '-s', 'PINCHART-PROC, PINCHART-CIP']
    #temp_user_args = ['-f', os.getcwd() + "/" + "RO Pinchart.xlsm", '-o', 'UPLOAD', '-s', 'PINCHART-PROC, PINCHART-CIP']
    temp_user_args = ['-f', os.getcwd() + "/" + "SLC Pinchart.xlsm", '-o', 'IMPORT']
    #temp_user_args = ['-f', os.getcwd() + "/" + "SLC Pinchart.xlsm", '-o', 'export']

    # sets the thread-id
    THREAD_ID = "MAIN"

    # this locks the output so if we 
    # wanted to write the output to 
    # a file, it would not error. This
    # does make the program a little slower
    # but we'll see how much.
    slock = threading.Lock()

    # setup the thread locks.
    # excel lock
    elock = threading.Lock()

    # opc lock
    olock = threading.Lock()

    # PC5 file lock - this protects
    # the PC5 file from getting overwritten
    # by other threads
    pc5lock = threading.Lock()

    utils.output(THREAD_ID, "__main__", "__main__", "PROCESSING USER ARGUMENTS.", slock)

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

    utils.output(THREAD_ID, "__main__", "__main__", "OPENPYXL - READING EXCEL FILE INTO MEMORY.", slock)
    
    # lock the excel
    elock.acquire()

    # gets all of the data from the workbook
    excel_dict = filereader.get_xl_data(arg_excel_file)

    # release the lock()
    elock.release()

    utils.output(THREAD_ID, "__main__", "__main__", "OPENPYXL - READING EXCEL FILE INTO MEMORY. COMPLETE.", slock)

    # sheets to process list. These will be given to the sheet
    # processor to process on different threads
    sheets_to_process = []
    
    # formulates the correct PLC OPERATION
    if arg_operation.lower() == "upload":
        utils.output(THREAD_ID, "__main__", "__main__", "OPERATION SET TO UPLOAD.")
        program_operation = PLC_OPERATION_UPLOAD
    elif arg_operation.lower() == "download":
        utils.output(THREAD_ID, "__main__", "__main__", "OPERATION SET TO DOWNLOAD.")
        program_operation = PLC_OPERATION_DOWNLOAD
    elif arg_operation.lower() == 'import':
        utils.output(THREAD_ID, "__main__", "__main__", "OPERATION SET TO IMPORT.")
        program_operation = PLC_OPERATION_IMPORT
    elif arg_operation.lower() == 'export':
        utils.output(THREAD_ID, "__main__", "__main__", "OPERATION SET TO EXPORT.")
        program_operation = PLC_OPERATION_EXPORT
    else:
        raise Exception("Invalid Input for Operation: %s" % arg_operation)
    # end if

    utils.output(THREAD_ID, "__main__", "__main__", "PREPPING SHEETS FOR PROCESSING.", slock)

    # get the main sheet first
    _main_sheet_data = excel_dict['Main Program']
    utils.output(THREAD_ID, "__main__", "__main__", "MAIN PROGRAM SHEET FOUND. EXTRACTING CONFIGURATION.", slock)

    # serialized the main sheet into a class
    main_sheet = serialize.MainSheetData(_main_sheet_data, "MAIN")

    # gets the configutation dictionary for the workbook. If it doesn't
    # find the configuration, it uses default values.
    config_data = main_sheet.get_config_data()

    utils.output(THREAD_ID, "__main__", "__main__", "CONFIGURATION FOUND: %s" % config_data, slock)

    for sheet_name in excel_dict:

        # gets the sheet data
        sheet_data = excel_dict[sheet_name]

        if (process_all_sheets or (sheet_name.lower() in arg_sheets)) \
            and sheet_name.lower() != "main program" and (sheet_name.lower().__contains__('plc') or \
                sheet_name.lower().__contains__('pinchart')):

            # add the sheet name to process on a different
            # thread
            sheets_to_process.append(
                {
                    'olock': olock,
                    'elock': elock,
                    'slock': slock,
                    'pc5lock': pc5lock,
                    'main_sheet_object': main_sheet,
                    'sheet_name': sheet_name,
                    'sheet_dict': sheet_data,
                    'config_data': config_data,
                    'operation': program_operation,
                    'full_file_path': arg_excel_file
                }
            )
        # end if
    # end for

    # loop through the sheets that need to get processed
    thread_id = 0

    utils.output(THREAD_ID, "__main__", "__main__", "PUTTING TOGETHER SHEET PROCESSING THREADS.", slock)

    sheet_threads = []
    for sheet_process_data in sheets_to_process:

        #elock, opclock, sheet_name, sheet_dict, config_data, operation
        if MULTITHREAD:

            # increment thread id
            thread_id = thread_id + 1
            sheet_process_data['thread_id'] = thread_id

            _worker_arguments = (
                sheet_process_data['olock'],
                sheet_process_data['elock'],
                sheet_process_data['slock'],
                sheet_process_data['pc5lock'],
                sheet_process_data['main_sheet_object'],
                sheet_process_data['sheet_name'],
                sheet_process_data['sheet_dict'],
                sheet_process_data['config_data'],
                sheet_process_data['operation'],
                sheet_process_data['thread_id'],
                sheet_process_data['full_file_path']
            )

            # process the sheet
            t = threading.Thread(group=None, target=worker.process_sheet, args=_worker_arguments)
            t.start()
            sheet_threads.append(t)

        else:

            worker.process_sheet(
                sheet_process_data['olock'],
                sheet_process_data['elock'],
                sheet_process_data['slock'],
                sheet_process_data['pc5lock'],
                sheet_process_data['main_sheet_object'],
                sheet_process_data['sheet_name'],
                sheet_process_data['sheet_dict'],
                sheet_process_data['config_data'],
                sheet_process_data['operation'],
                "MAIN",
                sheet_process_data['full_file_path']
            )
        # end if
    # end worker
    
    if MULTITHREAD:
        while worker.threads_running(sheet_threads) > 0:
            time.sleep(0.5)
        # end while
    # end if
    
    utils.output("MAIN", "__main__", "__main__", "Completed sheet processing.")
    
    # if there was an upload or import - save the workbook
    if program_operation == PLC_OPERATION_UPLOAD or program_operation == PLC_OPERATION_IMPORT:
        elock.acquire()
        utils.output("MAIN", "__main__", "__main__", "Saving workbook after upload.")
        xl = excel_interface.Interface(arg_excel_file)
        xl.save_workbook()
        elock.release()
    # end if
    
    utils.output("MAIN", "__main__", "__main__", "System Complete.")
# end main