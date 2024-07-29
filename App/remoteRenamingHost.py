import tkinter as tk
from itertools import combinations
import requests
import csv
import os
import subprocess

root = tk.Tk()

def open_filtered_csv(team):
    filtered_file = f'C:\\Users\\cristiano.actit\\PycharmProjects\\WindowsApps\\App\\TerminalFunctions\\filtered{team}.csv'
    if os.path.isfile(filtered_file):
        subprocess.run(['notepad.exe', filtered_file])
    else:
        print(f"Filtered file {filtered_file} does not exist.")

def show_confirmation_dialog(team):
    def on_confirm():
        perform_hostname_changes(team)
        dialog.destroy()
    
    def on_cancel():
        dialog.destroy()
    
    dialog = tk.Toplevel(root)
    dialog.title("Confirmar Alterações")
    message = tk.Label(dialog, text=f"Você confirma as alterações nos hostnames do arquivo filtered{team}.csv?")
    message.pack(pady=10)
    
    confirm_button = tk.Button(dialog, text="Sim", command=on_confirm)
    confirm_button.pack(side=tk.LEFT, padx=20, pady=10)
    
    cancel_button = tk.Button(dialog, text="Não", command=on_cancel)
    cancel_button.pack(side=tk.RIGHT, padx=20, pady=10)

    dialog.wait_window(dialog)

def parseNames(fullNames: str):
    parsedNames = [name.strip() for name in fullNames.split(",")]
    return parsedNames

def generate_combinations(name_parts):
    combinations_list = []
    first_name = name_parts[0].lower()
    for last_name in name_parts[1:]:
        combinations_list.append(f"{first_name}.{last_name.lower()}")
    return combinations_list

def query_ad(username, password):
    print("Running AD functions as:", username)

    def find_sam(names):
        fullNames = parseNames(names)
        for fullName in fullNames:
            name_parts = fullName.split()
            combinations = generate_combinations(name_parts)
            for combo in combinations:
                print(f"Testing combination: {combo}")
                ninja_search(combo)

    find_sam(str(input_names.get()))

def ninja_search(query):
    url = f"https://app.ninjarmm.com/v2/devices/search?q={query}&limit=3"
    headers = {
        'accept': 'application/json',
        'Cookie': f'sessionKey={sessionId_ninja.get()}'
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        try:
            data = response.json()
            # print(f"Raw data for {query}: {data}")  # print raw response data
            filter_response(data, query)

        except ValueError as e:
            print(f"Error decoding JSON for {query}: {e}")
    else:
        print(f"Failed to retrieve data for {query}. Status code: {response.status_code}")

def filter_response(data, query):
    filtered_results = []
    devices = data.get('devices', [])
    if not devices:
        print(f"No devices found for {query}")
        return

    for device in devices:
        matchAttrValue = device.get('matchAttrValue')
        if matchAttrValue and matchAttrValue.startswith("RIHAPPY\\"):
            actual_value = matchAttrValue.split("\\", 1)[1].lower()
            if actual_value == query:
                filtered_item = {
                    'id': device.get('id'),
                    'systemName': device.get('systemName'),
                    'matchAttrValue': matchAttrValue,
                    'newName': '' # here we put new AD hostname
                }
                filtered_results.append(filtered_item)
    if filtered_results:
        print(f"Exact match results for {query}: {filtered_results}")
        save_to_csv(filtered_results)
    else:
        print(f"No exact matches found for {query}")

def save_to_csv(results):
    team = input_team.get()
    file_exists = os.path.isfile(f'C:\\Users\\cristiano.actit\\PycharmProjects\\WindowsApps\\App\\TerminalFunctions\\filtered{team}.csv')
    with open(f'filtered{team}.csv', 'a', newline='') as csvfile:
        fieldnames = ['id', 'systemName', 'matchAttrValue', 'newName']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        for result in results:
            writer.writerow(result)

def save_cred():
    def on_enter():
        global user_cred
        global user_pass
        user_cred = username.get()
        user_pass = password.get()
        prompt.destroy()
        query_ad(user_cred, user_pass)
    
    prompt = tk.Toplevel(root)
    prompt.title("Login ActiveDirectory:")
    title = tk.Label(prompt, text="Insira seu usuário e senha:")
    username = tk.Entry(prompt, takefocus=1)
    password = tk.Entry(prompt, show="*")
    add_domain = tk.Button(prompt, text="Salvar", command=on_enter)
    exit_button = tk.Button(prompt, text="Fechar", command=prompt.destroy)
    
    title.pack()
    username.pack()
    password.pack()
    add_domain.pack()
    exit_button.pack()
    
    # Aguardar até que a janela seja destruída
    prompt.wait_window(prompt)

def perform_ad_search(query):
    try:
        powershell_command = f"Get-ADComputer -Filter {{Name -like '*{query}*'}} -SearchBase 'dc=rihappy, dc=local' | Select-Object Name | Sort-Object Name | Export-Csv -Path 'C:\\Users\\cristiano.actit\\PycharmProjects\\WindowsApps\\App\\TerminalFunctions\\query{query}Results.csv' -NoTypeInformation"
        subprocess.run(['powershell.exe', powershell_command], check=True)
        with open(f'C:\\Users\\cristiano.actit\\PycharmProjects\\WindowsApps\\App\\TerminalFunctions\\query{query}Results.csv', 'r', newline='') as file:
            reader = csv.reader(file)
            data = [row for row in reader if row and 'Name' not in row]
            data = sorted(data)
            print(f"AD search results for {query}: {data}")
    except subprocess.CalledProcessError as e:
        print(f"Error running PowerShell command: {e}")

def host_search():
    query = input_team.get()
    perform_ad_search(query)
    compare_lists(query)

def perform_hostname_changes(team):
    filtered_file = f'C:\\Users\\cristiano.actit\\PycharmProjects\\WindowsApps\\App\\TerminalFunctions\\filtered{team}.csv'
    
    try:
        with open(filtered_file, 'r', newline='') as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['newName']:
                    try:
                        ps_process = subprocess.Popen(['powershell.exe', '-NoExit', '-Command', '-'],
                                                  stdin=subprocess.PIPE,
                                                  stdout=subprocess.PIPE,
                                                  stderr=subprocess.PIPE,
                                                  text=True)
                    
                        # Creating PowerShell Credential Object
                        ps_commands = f"""
                        $pass = ConvertTo-SecureString "{user_pass}" -AsPlainText -Force
                        $credential = New-Object System.Management.Automation.PSCredential ("{user_cred}", $pass)
                        Rename-Computer rihappy.local -ComputerName {row['systemName']} -NewName {row['newName']} -Credential $credential -Force
                        """

                        # Running hostname changes PowerShell commands
                        stdout, stderr = ps_process.communicate(ps_commands)
                        print("STDOUT:", stdout)
                        print("STDERR:", stderr)
                        print(f"Successfully renamed {row['systemName']} to {row['newName']}")
                    except subprocess.CalledProcessError as e:
                        print(f"Error renaming {row['systemName']} to {row['newName']}: {e}")
    except FileNotFoundError:
        print(f"Filtered file {filtered_file} does not exist.")

def compare_lists(team):
    filtered_file = f'C:\\Users\\cristiano.actit\\PycharmProjects\\WindowsApps\\App\\TerminalFunctions\\filtered{team}.csv'
    query_file = f'C:\\Users\\cristiano.actit\\PycharmProjects\\WindowsApps\\App\\TerminalFunctions\\query{team}Results.csv'
    temp_file = f'C:\\Users\\cristiano.actit\\PycharmProjects\\WindowsApps\\App\\TerminalFunctions\\tempQuery{team}Results.csv'

    filtered_data = []
    query_data = []

    # reading filtered data
    if os.path.isfile(filtered_file):
        with open(filtered_file, 'r', newline='') as file:
            reader = csv.DictReader(file)
            filtered_data = list(reader)

    # reading query data
    with open(query_file, 'r', newline='') as file:
        reader = csv.DictReader(file)
        query_data = [row for row in reader if row and 'Name' in row]

    # processing temp hostnames
    temp_data = []
    new_query_data = []
    for row in query_data:
        if 'TEMP' in row['Name']:
            temp_data.append({
                'id': '',
                'systemName': row['Name'],
                'matchAttrValue': '',
                'newName': row['Name']
            })
        else:
            new_query_data.append(row)

    # add temp data to filtered file
    if temp_data:
        with open(filtered_file, 'a', newline='') as file:
            fieldnames = ['id', 'systemName', 'matchAttrValue', 'newName']
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            if os.path.getsize(filtered_file) == 0:
                writer.writeheader()
            writer.writerows(temp_data)

    # update query file with new data (removing temp hostnames)
    with open(temp_file, 'w', newline='') as file:
        fieldnames = ['Name']
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(new_query_data)

    # replace the original query file with the temporary file
    os.replace(temp_file, query_file)

    # get the last numeric value from new_query_data to increment
    last_number = 0
    if new_query_data:
        last_name = new_query_data[-1]['Name']
        last_number = int(last_name.split('-')[-1])

    # update field newName on filtered_data
    for item in filtered_data:
        system_name = item['systemName']
        if any(d['Name'] == system_name for d in new_query_data):
            item['newName'] = system_name
        else:
            last_number += 1
            item['newName'] = f"CORP-{team}-{last_number:04d}"

    # write updated data to a new temporary file
    with open(filtered_file, 'w', newline='') as file:
        fieldnames = ['id', 'systemName', 'matchAttrValue', 'newName']
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(filtered_data)
    print(f"Updated filtered data with new names in {filtered_file}")
    
    # open CSV file and show the confirmation dialog
    open_filtered_csv(team)
    show_confirmation_dialog(team)

def compare_and_confirm():
    team = input_team.get()
    compare_lists(team)

# frontend objects creation
mainFrame = tk.Frame(root, takefocus=0, padx=10, pady=10, width=200, height=200)
input_namesLabel = tk.Label(mainFrame, text="Insira a lista de nomes, separados por vírgula:")
input_names = tk.Entry(mainFrame, takefocus=1)
input_teamLabel = tk.Label(mainFrame, text="Insira o padrão do hostname:")
input_team = tk.Entry(mainFrame)
sessionId_ninjaLabel = tk.Label(mainFrame, text="Insira sua SessionID do NinjaRMM:")
sessionId_ninja = tk.Entry(mainFrame)
query_adButton = tk.Button(mainFrame, text="Localizar Usuários", command=save_cred)
host_searchButton = tk.Button(mainFrame, text="Pesquisar Hostnames", command=host_search)
compare_and_confirmButton = tk.Button(mainFrame, text="Comparar e Confirmar", command=compare_and_confirm)

# frontend objects exhibition
mainFrame.pack()
input_namesLabel.pack()
input_names.pack()
sessionId_ninjaLabel.pack()
sessionId_ninja.pack()
query_adButton.pack()
input_teamLabel.pack()
input_team.pack()
host_searchButton.pack()
compare_and_confirmButton.pack()

# run
root.mainloop()