from gettext import find
from unicodedata import name
from xml.dom.minidom import Identified
import win32evtlog
import xml.etree.ElementTree as ET
import ctypes
import sys
import pandas as pd

eventos = list()

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if is_admin():


    # open event file
    query_handle = win32evtlog.EvtQuery(
        'C:\Windows\System32\winevt\Logs\Application.evtx',
        win32evtlog.EvtQueryFilePath)
    

    read_count = 0

    while True:

        # read 1 record(s)
        events = win32evtlog.EvtNext(query_handle, 10)
        read_count += len(events)
        # if there is no record break the loop
        if len(events) == 0:
            break

        for event in events:
            xml_content = win32evtlog.EvtRender(event, win32evtlog.EvtRenderEventXml)
            # parse xml content
            xml = ET.fromstring(xml_content)


            # xml namespace, root element has a xmlns definition, so we have to use the namespace
            ns = '{http://schemas.microsoft.com/win/2004/08/events/event}'

            substatus = xml[0][1].text
            level_id = xml.find(f'.//{ns}Level').text
            event_id = xml.find(f'.//{ns}EventID').text
            computer = xml.find(f'.//{ns}Computer').text
            channel = xml.find(f'.//{ns}Channel').text
            execution = xml.find(f'.//{ns}Execution')
            process_id = execution.get('ProcessID')
            thread_id = execution.get('ThreadID')
            time_created = xml.find(f'.//{ns}TimeCreated').get('SystemTime')
            provider_name = xml.find(f'.//{ns}Provider').get('Name')
            channel = xml.find(f'.//{ns}Channel').text
            event_data = f'Time: {time_created}, Computer: {computer}, Event Id: {event_id}, Channel: {channel}, Process Id: {process_id}, Thread Id: {thread_id}, Provider: {provider_name}, Channel: {channel}, Level: {level_id}'

            #print(event_data)
            #eventos.append(event_data)
            #print(eventos)
            data = {'Time': [time_created], 'Computer': [computer], 'Event Id': [event_id], 'Chanel': [channel], 'Process Id': [process_id], 'Thread Id': [thread_id], 'Provider': [provider_name], 'Channel': [channel], 'Level': [level_id]}
            df = pd.DataFrame(data, columns = ['Time', 'Computer', 'Sustatus', 'Event Id', 'Chanel', 'Process Id', 'Thread Id', 'Provider', 'Channel', 'Level', 'User'])
            print(df)
            df.to_excel (r'C:\Users\win\Desktop\export_dataframe.xlsx', header=True)

user_data = xml.find(f'.//{ns}UserData')
            # user_data has possible any data
