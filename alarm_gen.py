import xml.etree.ElementTree as ET

# This file generates alarms in an xml file format. The input is the spreadsheet
# That the alarm

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

if __name__ == '__main__':

    # gets the element tree (base object)
    tree = ET.ElementTree(ET.fromstring(xml_file))

    # gets the root node
    root = tree.getroot()

    # gets the trigger element
    triggers = root.find('alarm/triggers')

    # gets the message element
    messages = root.find('alarm/messages')

    topic = "[WheyRO]"

    trigger_attributes = {
        "id":"T1", "type":"value", "ack-all-value":"1", "use-ack-all":"True", "ack-tag":"{[Topic]cAlarmTag.ALM.Ack}", "exp":"{[Topic]cAlarmTag.ALM.Active}",
        "message-tag":"", "message-handshake-exp":"", "message-notification-tag":"", "remote-ack-exp":"", "remote-ack-handshake-tag":"", "label":"Label1",
        "handshake-tag":""
    }
    message_attributes = {
        "id":"M1", "trigger-value":"1", "identifier":"5", "trigger":"#T1", "backcolor":"#800000", "forecolor":"#FFFFFF", "audio":"false", "display":"true",
        "print":"false", "message-to-tag":"false", "text":"/*S:0 {[Topic]cAlarmTag.Message}*/"
    }

    for i in range(0, 30):

        t1_id = i + 1
        t1 = trigger_attributes
        t1["id"] = "T" + str(t1_id)
        t1["ack-tag"] = "{" + topic + "Alarm1[" + str(i) + "].ALM.Ack}"
        t1["exp"] = "{" + topic + "Alarm1[" + str(i) + "].ALM.Active}"
        t1["label"] = "Label" + str(t1_id)

        m1_id = i + 1
        m1 = message_attributes

        m1["id"] = "M" + str(m1_id)
        m1["identifier"] = str(m1_id)
        m1["trigger"] = "#" + t1["id"]
        m1["text"] = "/*S:0 {" + topic + "Alarm1[" + str(i) + "].Message}*/"

        # add the trigger
        ET.SubElement(triggers, "trigger", t1)

        # add the message
        ET.SubElement(messages, "message", m1)

    # end for

    for i in range(0, 30):

        t2_id = 31 + i
        t2 = trigger_attributes
        t2["id"] = "T" + str(t2_id)
        t2["ack-tag"] = "{" + topic + "Alarm2[" + str(i) + "].ALM.Ack}"
        t2["exp"] = "{" + topic + "Alarm2[" + str(i) + "].ALM.Active}"
        t2["label"] = "Label" + str(t2_id)

        m2_id = 31 + i
        m2 = message_attributes
        m2["id"] = "M" + str(m2_id)
        m2["identifier"] = str(m2_id)
        m2["trigger"] = "#" + t2["id"]
        m2["text"] = "/*S:0 {" + topic + "Alarm2[" + str(i) + "].Message}*/"

        # add the trigger
        ET.SubElement(triggers, "trigger", t2)

        # add the message
        ET.SubElement(messages, "message", m2)

    # end for

    # write the xml file
    ET.indent(tree)
    tree.write("alarms.xml")
