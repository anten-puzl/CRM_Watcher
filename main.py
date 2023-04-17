import win32com.client as win32
import json
import requests
import os.path

TOKEN = 'Token_of_bot'
login_1c = 'your_login_to_1c'
password_1c = 'your_password_to_1c'
name_of_base = 'name_of_base'
name_of_file = 'chats.json'


def get_list_of_chats():
    if os.path.exists(name_of_file):
        with open(name_of_file, 'r') as jsonFile:
            return json.load(jsonFile)
    else:
        return []


def add_chat():
    if command == 'admin':
        added_value = {id_of_chat: 'admin'}
    else:
        added_value = {id_of_chat: 'user'}
    if os.path.exists(name_of_file):
        with open(name_of_file, 'r') as jsonFile:
            data = json.load(jsonFile)
            if not str(id_of_chat) in data:
                with open(name_of_file, 'w') as jsonFileWrite:
                    data.update(added_value)
                    json.dump(data, jsonFileWrite, indent=2)
    else:
        with open(name_of_file, 'w') as jsonFile:
            json_string = json.dumps(added_value)
            jsonFile.write(json_string)


def delete_chat():
    if os.path.exists(name_of_file):
        with open(name_of_file, 'r') as jsonFile:
            data = json.load(jsonFile)
            if id_of_chat in data:
                del data[id_of_chat]
                json_file = open(name_of_file, 'w')
                json_string = json.dumps(data)
                json_file.write(json_string)


try:
    url = f"https://api.telegram.org/bot{TOKEN}/getUpdates"
    answer = requests.get(url).json()
    json_str = json.dumps(answer)
    resp = json.loads(json_str)
    for msg in resp['result']:
        id_of_chat = msg['message']['chat']['id']
        id_of_chat = str(id_of_chat)
        command = msg['message']['text']
        if command == '/stop':
            delete_chat()
        else:
            add_chat()
except Exception as e:
    list_of_chats = get_list_of_chats()
    # error only for admin
    for chat_id in list_of_chats:
        if list_of_chats[chat_id] == 'admin':
            message = str(e)
            url = f"https://api.telegram.org/bot{TOKEN}/sendMessage?chat_id={list_of_chats[0]}&text={message}"
            requests.get(url).json()  # this sends the message

list_of_chats = get_list_of_chats()
try:
    for chat_id in list_of_chats:
        Connector1c = win32.gencache.EnsureDispatch('V83.COMConnector')
        agent = Connector1c.ConnectAgent("tcp://127.0.0.1:1540")
        cluster = agent.GetClusters()[0]
        agent.Authenticate(cluster, "", "")
        process = agent.GetWorkingProcesses(cluster)[0]
        port_number = process.MainPort
        Connector1c = win32.gencache.EnsureDispatch('V83.COMConnector')
        Server1c = Connector1c.ConnectWorkingProcess("tcp://127.0.0.1:"+str(port_number))
        Server1c.AddAuthentication(login_1c, password_1c)
        bases = Server1c.GetInfoBases()
        for base in bases:
            if base.Name == name_of_base:
                if base.ScheduledJobsDenied:
                    message = 'Регламентні завдання вимкнені'
                    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage?chat_id={chat_id}&text={message}"
                    requests.get(url).json()  # this sends the message
except Exception as e:
    # error only for admin
    for chat_id in list_of_chats:
        if list_of_chats[chat_id] == 'admin':
            message = str(e)
            url = f"https://api.telegram.org/bot{TOKEN}/sendMessage?chat_id={chat_id}&text={message}"
            requests.get(url).json()  # this sends the message
