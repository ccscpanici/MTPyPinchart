import os

# MULTITHREADED MODE - TURNS ON AND OFF THE MULTITHREADED
# FUNCTIONALILITY OF THE SYSTEM. THIS SPAWNS A NEW THREAD
# FOR EACH SHEET THAT IS GOING TO GET PROCESSED. IF YOU NEED TO
# DEBUG WHATS HAPPENING, A GOOD THING TO TRY WOULD BE TO
# SWITCH THIS TO FALSE, THEN, IT WILL PUT ALL OF THE
# PROCESSING ON ONE THREAD.
MULTITHREAD = True

# DEBUG MODE - TURNING THIS ON GIVES THE USER THE ABILITY TO
# USE THE DEBUG FILE VARIABLE AND THE DEBUG_USER_ARGS VARIABLE
# TO RUN THE SYSTEM.
DEBUG_MODE = False

# DEBUG_FILE = C:\\Users\\CCS\\Documents\\Pinchart.xlsm
# process the pinchart in the directory.
DEBUG_FILE = os.getcwd() + "\\" + "Pinchart.xlsm"
#DEBUG_FILE = "Z:\\data\Documents\\_Active Jobs\\CFR\\C 20082 Saputo Las Cruces Brine UF\\03_Documents\\03.1_PinCharts\\03.1.1_PinCharts\\20082 Pinchart.xlsm"

# DEBUG USER ARGUMENTS - WHEN IN DEBUG MODE, THE SYSTEM WILL USE THESE.
# THE USER CAN MODIFY WHAT THEY NEED TO DO.
DEBUG_USER_ARGS = ['-f', DEBUG_FILE, '-o', 'DOWNLOAD']

# THIS SETS THE MAXIMUM NUMBER OF CIP CONNECTIONS AT ONCE
# PLEASE - DON'T GO BEYOND 2 CONNECTIONS FOR NOW.
CIP_CONCURRENT_CONNECTIONS = 2
