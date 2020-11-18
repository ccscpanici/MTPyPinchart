from datetime import datetime
import utils
import numpy

def strip_array(tag: str) -> str:
    """
    Strip off the array portion of the tag
    'tag[100]' -> 'tag'
    """
    if '[' in tag:
        return tag[:tag.find('[')]
    return tag

def get_tag_indexing_string(a_tag):
    # this line splits the tag up using the dot
    # as the delimeter
    a_tag_array = a_tag.split('.')

    # this chunck of code gets rid of any tag
    # indexing.
    buffer = []
    for i in a_tag_array:
        if i.__contains__("[") and i.__contains__("]"):
            buffer.append(i[0:i.find("[")])
        else:
            buffer.append(i)

    # this forms the indexing string
    s = "" 
    for i in range(0, len(buffer)):
        if i < len(buffer) - 1:
            s = s + "['" + buffer[i] + "']['data_type']['internal_tags']"
        else:
            s = s + "['" + buffer[i] + "']['data_type']"
    return s

# this function takes in a DINT and outputs a list
# of the binary data.
def dec2bin(decimal_number):
    res = [int(i) for i in bin(decimal_number)[2:]] 
    return res
# end dec2bin

def get_tag_structure(a_tag_string):
     # first this is to get rid of all of the brackets []
    while a_tag_string.find("[") >= 0:
        _open = a_tag_string.find("[")
        _closed = a_tag_string.find("]")
        a_tag_string = a_tag_string[:_open] + a_tag_string[_closed + 1:]
    # end while
    return a_tag_string.split('.')
# end get_tag_structure

def find(a_dictionary, a_key):
    if a_key in a_dictionary: return a_dictionary[a_key]
    for k, v in a_dictionary.items():
        if isinstance(v,dict):
            item = find(v, a_key)
            if item is not None:
                return item
            # end if
        # end if
    # end for
# end find

def binstring_to_list(a_bin_string):

    int_list = []
    for char in a_bin_string:
        int_list.append(int(char))
    # end for
    return int_list

# end binstring_to_list

def get_32bit_bin(decimal_number):

    # convert it to a 32 bit integer
    d = numpy.int32(decimal_number)
    
    # gets the binary number
    binary_number = numpy.binary_repr(d, width=32)

    # if the length is over 32 than we don't really care
    # about the extra 1's
    if binary_number.__len__() > 32:
        difference = binary_number.__len__() - 32
        binary_number = binary_number[difference:]
    # end if

    # gets the binary list
    binary_list = binstring_to_list(binary_number)

    # return the binary representation of it
    return binary_list

# end get_32bit_bin()

def get_16bit_bin(decimal_number):
    
    # convert it to a 16-bit integer
    d = numpy.int16(decimal_number)
    
    # gets the binary number
    binary_number = numpy.binary_repr(d, width=16)

    # if the length is over 32 than we don't really care
    # about the extra 1's
    if binary_number.__len__() > 16:
        difference = binary_number.__len__() - 16
        binary_number = binary_number[difference:]
    # end if

    # gets the binary list
    binary_list = binstring_to_list(binary_number)

    # return the binary representation of it
    return binary_list

# end get_16bit_bin

def output(thread_id, class_name, method_name, message, thread_lock=None):

    # this keeps the output in the correct format
    now = datetime.now()

    # forumulates the thread name
    thread_name = "THREAD-%s" % str(thread_id).upper()

    # string format
    time_format = "%H:%M:%S"

    if thread_lock is not None:
        # blocking call for the output thread
        thread_lock.acquire()
    # end if

    # HH:MM:SS (module.method) (message)
    try:
        print(now.strftime(time_format) + "\t" + thread_name + "\t" + class_name + "\t" + method_name + "\t" + message.upper())
    except UnicodeEncodeError as ex:
        print(now.strftime(time_format) + "\t" + thread_name + "\t" + class_name + "\t" + method_name + "\t" + "UNICODE ENCODE ERROR.")
    # end try/except

    if thread_lock is not None:
        thread_lock.release()
    # end if
# end output

def data_converter(value, data_type_string):
    # this expects a value and a data type string, it 
    # will return the value in the requested data type
    _dt_string = str(data_type_string.lower().split()[0])

    if _dt_string == "int" or _dt_string == "int-16":
        if value is not None:
            _converted = numpy.int16(value)
        else:
            _converted = numpy.int16(0)
        # end if
    elif _dt_string == "dint" or _dt_string == "dint-32":
        if value is not None:
            _converted = numpy.int32(value)
        else:
            _converted = numpy.int32(0)
        # end if
    elif _dt_string == "float":
        if value is not None:
            _converted = float(value)
        else:
            _converted = float(0.0)
        # end if
    elif _dt_string == "string":
        if value is not None:
            _converted = str(value).strip()
        else:
            _converted = ""
        # end if
    else:
        raise ValueError("Incorrect data type string in method.")
    # end if
    return _converted
# end validate_data_value

def clean_int_list(a_list):
    # this function takes in a list
    # and returns a cleaner version
    # of it.
    return_list = []
    for i in a_list:
        try:
            return_list.append(int(i))
        except:
            continue
        # end try
    # end for
    return return_list
# end clean_int_list