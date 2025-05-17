import win32com.client as win32
import json
import requests
import os.path
import os
from dotenv import load_dotenv

# Загрузка переменных окружения из файла .env
load_dotenv()

# Получение переменных окружения
TOKEN = os.getenv('TOKEN')
login_1c = os.getenv('LOGIN_1C')
password_1c = os.getenv('PASSWORD_1C')
name_of_base = os.getenv('NAME_OF_BASE')
name_of_file = os.getenv('NAME_OF_FILE')
error_message_file = os.getenv('ERROR_MESSAGE_FILE')


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


def read_error_messages():
    """Читает сообщения из файла с ошибками"""
    if os.path.exists(error_message_file):
        with open(error_message_file, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return None


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
            url = f"https://api.telegram.org/bot{TOKEN}/sendMessage?chat_id={chat_id}&text={message}"
            requests.get(url).json()  # this sends the message

list_of_chats = get_list_of_chats()
try:
    # Проверяем наличие сообщений в файле
    error_message = read_error_messages()
    if error_message:
        # Отправляем сообщение всем пользователям
        for chat_id in list_of_chats:
            url = f"https://api.telegram.org/bot{TOKEN}/sendMessage?chat_id={chat_id}&text={error_message}"
            requests.get(url).json()
    
    # Проверка 1С
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
