import xml.etree.ElementTree as ET

# This file generates alarms in an xml file format. The input is the spreadsheet
# That the alarm

# controller name for the alarm points
controller = "WheyRO"

# define the alarm files. AlarmX should be the base tag
# inside the controller. The "length" attribute should be
# the length of the alarm array.
alarm_files = {
    'Alarm1': {"length" : 30},
    'Alarm2': {"length" : 30},
}

# ---------------------------------------------------------- DO NOT TOUCH ---------------------------------------------------------------
if __name__ == '__main__':

    # xml template file that we will use. The system will populated the triggers and messsages
    xml_file =          """<?xml version="1.0" encoding="UTF-8"?>
                            <alarms version="1.0" product="{E44CB020-C21D-11D3-8A3F-0010A4EF3494}" id="Alarms">
                                <alarm history-size="128" capacity-high-warning="90" capacity-high-high-warning="99" display-name="[ALARM]" hold-time="250" max-update-rate="1.00" embedded-server-update-rate="1.00" silence-tag="" remote-silence-exp="" remote-ack-all-exp="" status-reset-tag="" remote-status-reset-exp="" close-display-tag="" remote-close-display-exp="" use-alarm-identifier="false" capacity-high-warning-tag="" capacity-high-high-warning-tag="" capacity-overrun-tag="" remote-clear-history-exp="">
                                    <triggers>
                                    </triggers>
                                    <messages>
                                    </messages>
                                </alarm>
                            </alarms>"""

    # gets the element tree (base object)
    tree = ET.ElementTree(ET.fromstring(xml_file))

    # gets the root node
    root = tree.getroot()

    # gets the trigger element
    triggers = root.find('alarm/triggers')

    # gets the message element
    messages = root.find('alarm/messages')

    trigger_attributes = {
        "id":"T1", "type":"value", "ack-all-value":"1", "use-ack-all":"True", "ack-tag":"{[Topic]cAlarmTag.ALM.Ack}", "exp":"{[Topic]cAlarmTag.ALM.Active}",
        "message-tag":"", "message-handshake-exp":"", "message-notification-tag":"", "remote-ack-exp":"", "remote-ack-handshake-tag":"", "label":"Label1",
        "handshake-tag":""
    }
    message_attributes = {
        "id":"M1", "trigger-value":"1", "identifier":"5", "trigger":"#T1", "backcolor":"#800000", "forecolor":"#FFFFFF", "audio":"false", "display":"true",
        "print":"false", "message-to-tag":"false", "text":"/*S:0 {[Topic]cAlarmTag.Message}*/"
    }

    # gets a list if the alarm files
    a = alarm_files.keys()

    # creates the topic / shortcut string
    topic = "[" + controller + "]"
    
    # alarm trigger / message id index
    id_index = 0
    
    for name in a: # looping through alarm files
        
        alarm_file_length = alarm_files[name]['length']

        for i in range(0, alarm_file_length): # looping through indexes for generated alarms
            
            # generate the trigger node attributes
            trigger_id = id_index + 1
            trigger = trigger_attributes
            trigger["id"] = "T" + str(trigger_id)
            trigger["ack-tag"] = "{" + topic + name + "[" + str(i) + "].ALM.Ack"
            trigger["exp"] = "{" + topic + name + "[" + str(i) + "].ALM.Active"
            trigger["label"] = "Label" + str(trigger_id)

            # generate the message node attributes
            message = message_attributes
            message["id"] = "M" + str(trigger_id)
            message["identifier"] = str(trigger_id)
            message["trigger"] = "#" + trigger["id"]
            message["text"] = "/*S:0 {" + topic + name + "[" + str(i) + "].Message} */"

            # add the trigger node to the XML
            ET.SubElement(triggers, "trigger", trigger)

            # add the message node to the XML
            ET.SubElement(messages, "message", message)

            id_index = id_index + 1
        # end for
    # end for

    # indents the XML file so it is ledgible in the file instead of
    # one giant string block
    ET.indent(tree)

    # write the xml file
    tree.write("alarms.xml")
    
    print("----- Alarm Generation Complete -----")
