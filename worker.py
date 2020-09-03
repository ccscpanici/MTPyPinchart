
# # WORKER MODULE IS FOR PROCESSING A SHEET
# ON A SEPARATE THREAD
def process_sheet(self, elock, opclock, sheet_dict, config_data, operation):
    pass

    # get the sheet serializer

    # get the PLC data

    # loop through different chunks of data

    # if upload or download
    #   create an opc class
    #   put the opclock on to prevent other threads
    #   from accessing the opc server (just in case).
    #   maybe in the future we could even take this out

    # if upload
    #   get the plc data and values
    #   put the new values in the sheet serializer

    # if download
    #   download the data into the plc
    #
    # after the OPC operations are completed, release the
    # lock on the opc (opcclock.release())

    # if upload
    #   put a lock on the excel workbook
    #   create the win32com excel class
    #   update the data with the plc values
    #   

# end if