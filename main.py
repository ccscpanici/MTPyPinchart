import os
import filereader
import serialize
import threading
import worker
import getopt
import sys
import utils
import time
import Cip

if sys.platform == "win32":
    #import pywintypes
    #pywintypes.datetime = pywintypes.TimeType
    pass
# end if

PLC_OPERATION_DOWNLOAD = 0
PLC_OPERATION_UPLOAD = 1
PLC_OPERATION_IMPORT = 2
PLC_OPERATION_EXPORT = 3

# USE THIS BOOL TO TURN MULTITHREADED
# ON AND OFF
MULTITHREAD = True
DEBUG_MODE = True

# Sets the maximum number of concurrent
# PLC connections
CIP_CONCURRENT_CONNECTIONS = 2

OPC_SERVER = 'RSLinx OPC Server'

if __name__ == '__main__':

    if DEBUG_MODE:
        # temporary user arguments
        _file = "Z:\\data\\Documents\\_Active Jobs\\CFR\\C 20034 Saputo Almena\\03_Documents\\03.1_PinCharts\\03.1.1_PinCharts\\_20034 Tags and Setpoints-rev2.3.xlsm"
        #temp_user_args = ['-f', _file, '-o', 'UPLOAD', '-s', 'PLC DATA-Valve Config']
        #temp_user_args = ['-f', os.getcwd() + "/" + "RO Pinchart.xlsm", '-o', 'DOWNLOAD', '-s', 'PINCHART-CIP']
        #temp_user_args = ['-f', os.getcwd() + "/" + "RO Pinchart.xlsm", '-o', 'DOWNLOAD', '-s', 'PINCHART-PROC, PINCHART-CIP']
        #temp_user_args = ['-f', os.getcwd() + "/" + "RO Pinchart.xlsm", '-o', 'DOWNLOAD']
        temp_user_args = ['-f', os.getcwd() + "/" + "RO Pinchart.xlsm", '-o', 'UPLOAD']
        #temp_user_args = ['-f', os.getcwd() + "/" + "RO Pinchart.xlsm", '-o', 'UPLOAD', '-s', 'PINCHART-PROC, PINCHART-CIP']
        #temp_user_args = ['-f', os.getcwd() + "/" + "SLC Pinchart.xlsm", '-o', 'IMPORT']
        #temp_user_args = ['-f', os.getcwd() + "/" + "SLC Pinchart.xlsm", '-o', 'export']
    # end if

    # sets the thread-id
    THREAD_ID = "MAIN"

    # this locks the output so if we 
    # wanted to write the output to 
    # a file, it would not error. This
    # does make the program a little slower
    # but we'll see how much.
    slock = threading.Lock()

    # initializes the worker keyword arguments
    worker_kwargs = {'slock':slock}

    utils.output(THREAD_ID, "__main__", "__main__", "PROCESSING USER ARGUMENTS.", slock)

    # user arguments
    #user_args = sys.argv[1:]
    #print("user_args = %s" % user_args)

    # system arguments possibilities
    # f - filename
    # o - operation
    # s - sheet names - if this doesn't exist then we will assume all the sheets
    # print the arguments
    if DEBUG_MODE:
        args = getopt.getopt(temp_user_args, "f:o:s:")
    else:
        args = getopt.getopt(sys.argv[1:], "hf:o:s:")
    # end if

    for i in args:
        utils.output(THREAD_ID, "__main__", "__main__", "User argument: %s" % i, slock)
    # end for

    # loop through the arguments and set the variables
    arg_excel_file = False
    arg_operation = False
    arg_sheets = False
    process_all_sheets = False
    for switch, value in args[0]:
        if switch in ['-f', '--file']:
            arg_excel_file = value
        elif switch in ['-o', '--operation']:
            arg_operation = value
        elif switch in ['-s', '--sheet']:

            # gets the user called sheets to process
            arg_sheets = []
            for i in value.strip().split(","):
                arg_sheets.append(i.strip())
            # end for

            # lower case the sheet
            arg_sheets = [i.lower() for i in arg_sheets]
        elif switch in ['-h', '--help']:
            print("--------------------------------------------------------------------------------")
            print("Usage: python main.py -f <excel file path>, -o <operation string>, (optional) -s <sheet list comma separated string>")
            print("--------------------------------------------------------------------------------")
            print("\tswitch\t\t\t\tdescription")
            print("\t-h\t--help\t\t\tOutput help for usage.")
            print("\t-f\t--file\t\t\tInput file for processing operations.")
            print("\t-o\t--operation\t\tOperation on file (UPLOAD, DOWNLOAD, IMPORT, or EXPORT")
            print("\t-s\t--sheet\t\t\tOPTIONAL: Sheet string to process (default all sheets). Comma separated.")
            sys.exit(1)
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

    # gets all of the data from the workbook
    excel_dict = filereader.get_xl_data(arg_excel_file)

    utils.output(THREAD_ID, "__main__", "__main__", "OPENPYXL - READING EXCEL FILE INTO MEMORY. COMPLETE.", slock)

    # sheets to process list. These will be given to the sheet
    # processor to process on different threads
    sheets_to_process = []
    
    # formulates the correct PLC OPERATION
    if arg_operation.lower().strip() == "upload":
        utils.output(THREAD_ID, "__main__", "__main__", "OPERATION SET TO UPLOAD.")
        program_operation = PLC_OPERATION_UPLOAD
    elif arg_operation.lower().strip() == "download":
        utils.output(THREAD_ID, "__main__", "__main__", "OPERATION SET TO DOWNLOAD.")
        program_operation = PLC_OPERATION_DOWNLOAD
    elif arg_operation.lower().strip() == 'import':
        utils.output(THREAD_ID, "__main__", "__main__", "OPERATION SET TO IMPORT.")
        program_operation = PLC_OPERATION_IMPORT
    elif arg_operation.lower().strip() == 'export':
        utils.output(THREAD_ID, "__main__", "__main__", "OPERATION SET TO EXPORT.")
        program_operation = PLC_OPERATION_EXPORT
    else:
        raise Exception("Invalid Input for Operation: %s" % arg_operation)
    # end if

    # adds the operation to the worker kwargs
    worker_kwargs['operation'] = program_operation

    utils.output(THREAD_ID, "__main__", "__main__", "PREPPING SHEETS FOR PROCESSING.", slock)

    # get the main sheet first
    _main_sheet_data = excel_dict['Main Program']
    utils.output(THREAD_ID, "__main__", "__main__", "MAIN PROGRAM SHEET FOUND. EXTRACTING CONFIGURATION.", slock)

    # serialized the main sheet into a class
    main_sheet = serialize.MainSheetData(_main_sheet_data, "MAIN")

    # gets the configutation dictionary for the workbook. If it doesn't
    # find the configuration, it uses default values.
    config_data = main_sheet.get_config_data()
    worker_kwargs['config_data'] = config_data
    utils.output(THREAD_ID, "__main__", "__main__", "CONFIGURATION FOUND: %s" % config_data, slock)

    if program_operation == PLC_OPERATION_DOWNLOAD or program_operation == PLC_OPERATION_UPLOAD:

        ip_address = main_sheet.get_search_offset_value("IP Address", 0, 1)
        slot_number = main_sheet.get_search_offset_value("Slot", 0, 1)

        worker_kwargs['ip_address'] = ip_address
        worker_kwargs['slot_number'] = slot_number

        # instantiates a new CIP manager
        cip_manager = Cip.ConnManager(CIP_CONCURRENT_CONNECTIONS)
        
        if ip_address is None or slot_number is None:
            raise Exception("Could not locate IP address or slot number.")
        else:
            cip_path = "%s/%s" % (ip_address, slot_number)
            cip_manager.wait_for_connection()

            utils.output(THREAD_ID, "__main__", "__main__", "Connecting to controller @ %s" % cip_path, slock)
            c = Cip.LogixController(ip_address, slot_number)

            # we give this to the worker class, that way it
            # will not have to do it again.
            ts1 = time.time()
            _tags = c.get_plc_tags()
            _plc_info = c.plc_info
            ts2 = time.time()
            ts_delta = ts2 - ts1
            utils.output(THREAD_ID, "__main__", "__main__", "Took %s seconds to connect to controller and read tags." % ts_delta)

            # remove the connection
            cip_manager.remove_connection()
        # end if

        # adds the data to the worker kwargs
        worker_kwargs['cip_manager'] = cip_manager
        worker_kwargs['plc_tags'] = _tags

    if program_operation == PLC_OPERATION_UPLOAD or program_operation == PLC_OPERATION_IMPORT:
        elock = threading.Lock()
        worker_kwargs['elock'] = elock
        worker_kwargs['excel_file'] = arg_excel_file

    if program_operation == PLC_OPERATION_IMPORT or program_operation == PLC_OPERATION_EXPORT:
        pc5_lock = threading.Lock()
        _pc5_file = main_sheet.get_pc5_file()

        worker_kwargs['pc5_lock'] = pc5_lock
        worker_kwargs['pc5_file'] = _pc5_file

    # end if

    sheets = []
    for sheet_name in excel_dict:

        # gets the sheet data
        sheet_data = excel_dict[sheet_name]

        if (process_all_sheets or (sheet_name.lower() in arg_sheets)) \
            and sheet_name.lower() != "main program" and (sheet_name.lower().__contains__('plc') or \
                sheet_name.lower().__contains__('pinchart')):

            # add the sheet name to process on a different
            # thread
            worker_kwargs['sheet_name'] = sheet_name
            worker_kwargs['sheet_dict'] = sheet_data

            # save a buffer
            sheets_to_process.append(worker_kwargs.copy())
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

            # process the sheet
            t = threading.Thread(group=None, target=worker.process_sheet, kwargs=sheet_process_data)
            t.start()
            sheet_threads.append(t)

        else:
            sheet_process_data['thread_id'] = "MAIN"
            worker.process_sheet(**sheet_process_data)
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
        import excel_interface
        
        elock.acquire()
        utils.output("MAIN", "__main__", "__main__", "Saving workbook after upload.")
        xl = excel_interface.Interface(arg_excel_file)
        xl.save_workbook()
        elock.release()
    # end if
    
    utils.output("MAIN", "__main__", "__main__", "System Complete.")
# end main