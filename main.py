import win32com.client as win32
import json
import requests
import os
from dotenv import load_dotenv

# Load environment variables from .env
load_dotenv()

# Environment variables
TOKEN = os.getenv('TOKEN')
login_1c = os.getenv('LOGIN_1C')
password_1c = os.getenv('PASSWORD_1C')
name_of_base = os.getenv('NAME_OF_BASE')
name_of_file = os.getenv('NAME_OF_FILE')
error_message_dir = os.getenv('ERROR_MESSAGE_DIR')
last_update_file = 'last_update_id.txt'


def get_list_of_chats():
    if os.path.exists(name_of_file):
        with open(name_of_file, 'r') as jsonFile:
            return json.load(jsonFile)
    else:
        return {}


def add_chat():
    # Add user without password check
    if command.startswith('/admin'):
        added_value = {id_of_chat: 'admin'}
    else:
        added_value = {id_of_chat: 'user'}

    if os.path.exists(name_of_file):
        with open(name_of_file, 'r') as jsonFile:
            data = json.load(jsonFile)
            if str(id_of_chat) not in data:
                with open(name_of_file, 'w') as jsonFileWrite:
                    data.update(added_value)
                    json.dump(data, jsonFileWrite, indent=2)
                message = "You have been successfully registered!"
                url = f"https://api.telegram.org/bot{TOKEN}/sendMessage?chat_id={id_of_chat}&text={message}"
                requests.get(url).json()
    else:
        with open(name_of_file, 'w') as jsonFile:
            json_string = json.dumps(added_value)
            jsonFile.write(json_string)
            message = "You have been successfully registered!"
            url = f"https://api.telegram.org/bot{TOKEN}/sendMessage?chat_id={id_of_chat}&text={message}"
            requests.get(url).json()


def delete_chat():
    if os.path.exists(name_of_file):
        with open(name_of_file, 'r') as jsonFile:
            data = json.load(jsonFile)
            if id_of_chat in data:
                del data[id_of_chat]
                with open(name_of_file, 'w') as json_file:
                    json.dump(data, json_file, indent=2)
                message = "You have been removed from the subscription list"
                url = f"https://api.telegram.org/bot{TOKEN}/sendMessage?chat_id={id_of_chat}&text={message}"
                requests.get(url).json()


def read_error_files():
    """Read messages from all error files in the given directory"""
    error_messages = []

    if not os.path.exists(error_message_dir) or not os.path.isdir(error_message_dir):
        return error_messages

    files = os.listdir(error_message_dir)

    for file in files:
        file_path = os.path.join(error_message_dir, file)
        if os.path.isfile(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                    if content:
                        error_messages.append((file_path, content))
            except Exception as e:
                print(f"Error reading file {file_path}: {str(e)}")

    return error_messages


def get_last_update_id():
    if os.path.exists(last_update_file):
        with open(last_update_file, 'r') as f:
            return int(f.read().strip())
    return None


def save_last_update_id(update_id):
    with open(last_update_file, 'w') as f:
        f.write(str(update_id))


# --- Main processing ---
try:
    last_update_id = get_last_update_id()
    url = f"https://api.telegram.org/bot{TOKEN}/getUpdates"
    if last_update_id is not None:
        url += f"?offset={last_update_id + 1}"

    answer = requests.get(url).json()
    json_str = json.dumps(answer)
    resp = json.loads(json_str)

    if not os.path.exists(name_of_file):
        with open(name_of_file, 'w') as f:
            json.dump({}, f)

    for msg in resp['result']:
        if 'update_id' in msg:
            save_last_update_id(msg['update_id'])

        if 'message' in msg and 'text' in msg['message']:
            id_of_chat = str(msg['message']['chat']['id'])
            command = msg['message']['text']

            list_of_chats = get_list_of_chats()

            if command == '/stop':
                delete_chat()
            elif command == '/start':
                if id_of_chat not in list_of_chats:
                    add_chat()
                else:
                    message = "You are already registered in the system"
                    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage?chat_id={id_of_chat}&text={message}"
                    requests.get(url).json()

except Exception as e:
    list_of_chats = get_list_of_chats()
    for chat_id in list_of_chats:
        if list_of_chats[chat_id] == 'admin':
            message = str(e)
            url = f"https://api.telegram.org/bot{TOKEN}/sendMessage?chat_id={chat_id}&text={message}"
            requests.get(url).json()

# --- Error reporting and 1C job check ---
list_of_chats = get_list_of_chats()

try:
    error_files = read_error_files()

    if error_files:
        for chat_id in list_of_chats:
            for file_path, error_message in error_files:
                url = f"https://api.telegram.org/bot{TOKEN}/sendMessage?chat_id={chat_id}&text={error_message}"
                requests.get(url).json()

        for file_path, _ in error_files:
            if os.path.exists(file_path):
                os.remove(file_path)

    for chat_id in list_of_chats:
        Connector1c = win32.gencache.EnsureDispatch('V83.COMConnector')
        agent = Connector1c.ConnectAgent("tcp://127.0.0.1:1540")
        cluster = agent.GetClusters()[0]
        agent.Authenticate(cluster, "", "")
        process = agent.GetWorkingProcesses(cluster)[0]
        port_number = process.MainPort
        Connector1c = win32.gencache.EnsureDispatch('V83.COMConnector')
        Server1c = Connector1c.ConnectWorkingProcess(f"tcp://127.0.0.1:{port_number}")
        Server1c.AddAuthentication(login_1c, password_1c)
        bases = Server1c.GetInfoBases()

        for base in bases:
            if base.Name == name_of_base:
                if base.ScheduledJobsDenied:
                    message = 'Scheduled jobs are disabled'
                    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage?chat_id={chat_id}&text={message}"
                    requests.get(url).json()

except Exception as e:
    for chat_id in list_of_chats:
        if list_of_chats[chat_id] == 'admin':
            message = str(e)
            url = f"https://api.telegram.org/bot{TOKEN}/sendMessage?chat_id={chat_id}&text={message}"
            requests.get(url).json()
