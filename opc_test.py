import OpenOPC
import threading
import time
import pywintypes

pywintypes.datetime = pywintypes.TimeType

def upload_opc_data(tags, group):

    opc = OpenOPC.client()
    opc.connect('RSLinx OPC Server')

    data = opc.read(tags=tags)
    print(data)

# end

def threads_running(thread_list):
    count = 0
    for i in thread_list:
        if i.is_alive():
            count = count + 1
        # end if
    # end for
    return count
# end threads_running

if __name__ == '__main__':
   
    tags1 = ["[UF_RO_POL]RO_Data[3].ACK_DESC[0]", "[UF_RO_POL]RO_Data[3].ACK_DESC[1]", "[UF_RO_POL]RO_Data[3].ACK_DESC[2]"]
    tags2 = ["[UF_RO_POL]RO_Data[3].ACK_DESC[3]", "[UF_RO_POL]RO_Data[3].ACK_DESC[4]", "[UF_RO_POL]RO_Data[3].ACK_DESC[5]"]
    tags3 = ["[UF_RO_POL]RO_Data[3].ACK_DESC[6]", "[UF_RO_POL]RO_Data[3].ACK_DESC[7]", "[UF_RO_POL]RO_Data[3].ACK_DESC[8]"]
    #print(opc.read(tags))

    olock = threading.Lock()

    t1 = threading.Thread(target=upload_opc_data, args=(tags1, 1))

    t1.start()

    t2 = threading.Thread(target=upload_opc_data, args=(tags2, 2))
    
    t2.start()

    t3 = threading.Thread(target=upload_opc_data, args=(tags3, 3))
    
    t3.start()

    while t1.is_alive() or t2.is_alive() or t3.is_alive():
        time.sleep(1)
    # end while
# end main