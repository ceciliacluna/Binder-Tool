from tkinter import Tk, Label, Frame, Entry, Button, END, TOP, X, HORIZONTAL, BOTTOM, S
from tkinter.ttk import Progressbar
from tkinter.messagebox import showerror
import tkinter.filedialog as filedialog
import requests
import pandas as pd
import traceback


class BinderTool:
    def __init__(self, master):
        self.master = master
        master.title('Binder Tool')

        self.input_frame = Frame(master)
        self.top_frame = Frame(master)
        self.bottom_frame = Frame(master)
        self.line = Frame(master, height=1, width=400, bg="grey80", relief='groove')

        self.url_input = Label(self.input_frame, text="  url:")
        self.url_entry = Entry(self.input_frame, text="", width=40)

        self.user_name_input = Label(self.input_frame, text="username:    ")
        self.user_name_entry = Entry(self.input_frame, text="", width=40)

        self.password_input = Label(self.input_frame, text="password:   ")
        self.password_entry = Entry(self.input_frame, text="", width=40)

        self.date_input = Label(self.input_frame, text="date (mm-dd-yyyy):   ")
        self.date_entry = Entry(self.input_frame, text="", width=40)

        self.input_path_label = Label(self.top_frame, text="Input File Path:")
        self.input_entry = Entry(self.top_frame, text="", width=40)
        self.browse1 = Button(self.top_frame, text="Browse", command=self.input_location)

        self.output_path_label = Label(self.bottom_frame, text="Output File Path:")
        self.output_entry = Entry(self.bottom_frame, text="", width=40)
        self.browse2 = Button(self.bottom_frame, text="Browse", command=self.output_location)

        self.progress = Progressbar(master, orient=HORIZONTAL, length=100, mode='indeterminate')

        self.begin_button = Button(self.bottom_frame, text='Begin!', command=self.data_automation)

        # LAYOUT

        self.input_frame.pack(side=TOP, pady=5)
        self.line.pack(pady=10)
        self.top_frame.pack()
        self.line.pack(pady=10)
        self.bottom_frame.pack()

        self.url_input.grid(row=0, column=0, pady=5)
        self.url_entry.grid(row=0, column=1, pady=5)

        self.user_name_input.grid(row=1, column=0, pady=5)
        self.user_name_entry.grid(row=1, column=1, pady=5)

        self.password_input.grid(row=2, column=0, pady=5)
        self.password_entry.grid(row=2, column=1, pady=5)

        self.date_input.grid(row=3, column=0, pady=5)
        self.date_entry.grid(row=3, column=1, pady=5)

        self.input_path_label.pack(pady=5)
        self.input_entry.pack(pady=5)
        self.browse1.pack(pady=5)

        self.output_path_label.pack(pady=5)
        self.output_entry.pack(pady=5)
        self.browse2.pack(pady=5)

        self.begin_button.pack(pady=20, fill=X)

    def input_location(self):
        global input_path
        input_path = filedialog.askopenfilename(title="Select a file", filetypes=[("Excel files", ".xlsx .xls")])
        self.input_entry.delete(1, END)  # Remove current text in entry
        self.input_entry.insert(0, input_path)  # Insert the 'path'

    def output_location(self):
        global output_path
        output_path = filedialog.askdirectory()
        self.output_entry.delete(1, END)  # Remove current text in entry
        self.output_entry.insert(0, output_path)  # Insert the 'path'

    def data_automation(self):
        url = self.url_entry.get()
        username = self.user_name_entry.get()
        password = self.password_entry.get()
        date = self.date_entry.get()
        full_url = url + '/api/v19.1/auth'
        payload = {'username': username,
                   'password': password}
        files = [

        ]
        headers = {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
        }
        try:
            response = requests.request("POST", full_url, headers=headers, params=payload, files=files)
            auth_content = response.json()
            session_id = auth_content['sessionId']
        except Exception as ex:
            showerror(title="Error", message=ex)
        full_url = url + '/api/v19.1/objects/binders/'
        payload = {}
        files = {}
        headers = {
            'Accept': 'application/json',
            'Authorization': session_id

        }
        binder = pd.read_excel(input_path)
        output_df = pd.DataFrame(columns=['Binder ID', 'name__v', 'id'])
        self.progress.pack(side=BOTTOM, anchor=S, pady=10)
        for index, row in binder.iterrows():
            self.progress.step()
            binder_id = str(row['Document ID'])
            url_id = full_url + binder_id + '?depth=all'
            response = requests.request("GET", url_id, headers=headers, data=payload, files=files)
            json_file = response.json()
            print(json_file)
            master.update()
            json_parse = json_file['binder']['nodes']
            for x in json_parse:
                name__v = x['properties']['name__v']
                ID = x['properties']['id']
                new_row = {'Binder ID': binder_id, 'name__v': name__v, 'id': ID}
                output_df = output_df.append(new_row, ignore_index=True)

        output_df.to_csv(output_path + '/binder_output_' + date + '.csv', index=False)
        completed = Label(master, text="Download Complete", fg="green", font="Helvetica 10 bold", pady=6)
        completed.pack()
        self.progress.pack_forget()


master = Tk()
my_gui = BinderTool(master)
master.mainloop()