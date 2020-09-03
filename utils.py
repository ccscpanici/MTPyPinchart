from datetime import datetime
import numpy

def output(class_name, method_name, message):

    # this keeps the output in the correct format
    now = datetime.now()

    # string format
    time_format = "%H:%M:%S"

    # HH:MM:SS (module.method) (message)
    try:
        print(now.strftime(time_format) + "\t" + class_name + "\t" + method_name + "\t" + message)
    except UnicodeEncodeError as ex:
        print(now.strftime(time_format) + "\t" + class_name + "\t" + method_name + "\t" + "UNICODE ENCODE ERROR.")
    # end try/except
# end output

def data_converter(self, value, data_type_string):
    # this expects a value and a data type string, it 
    # will return the value in the requested data type
    _dt_string = str(data_type_string.lower())

    if _dt_string == "int" or _dt_string == "int-16":
        if value is not None:
            _value = numpy.Int16(_value)
        else:
            _value = numpy.Int16(0)
        # end if
    elif _dt_string == "dint" or _dt_string == "dint-32":
        if value is not None:
            _value = numpy.Int32(_value)
        else:
            _value = numpy.Int32(0)
        # end if
    elif _dt_string == "float":
        if value is not None:
            _value = float(value)
        else:
            _value = float(0.0)
        # end if
    elif _dt_string == "string":
        if value is not None:
            _value = str(value)
        else:
            _value = ""
        # end if
    else:
        raise ValueError("Incorrect data type string in method.")
    # end if
# end validate_data_value