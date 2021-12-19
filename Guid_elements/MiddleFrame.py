from tkinter import Button, OptionMenu
from tkinter.font import BOLD, Font
from tkinter import Frame, Label, StringVar, ttk
from tkinter.constants import E, END, N, W
import pandas as pd
from Enums.GuidParameter import GuidParameter
from tkinter.filedialog import *

from Guid_elements.SaveDialog import SaveDialog

class MiddleFrame:
    def __init__(self, root, tab):
        self.root = root
        # font
        self.courier_font = Font(root, family="Courier", size=18, weight=BOLD)
        self.courier_small_font = Font(root, family="Courier", size=16, weight=BOLD)
        self.segoe_font = Font(root, family="Segoe UI", size=12)
        self.segoe_small_font = Font(root, family="Segoe UI", size=10)
        self.segoe_bold_font = Font(root, family="Segoe UI", size=12, weight=BOLD)

        # middle_frame
        self.middle_frame = Frame(tab)
        self.middle_frame.grid(row=0, column=1)

        self.report_type_var = StringVar(self.middle_frame)
        # Report type label
        report_type_label = Label(self.middle_frame, textvariable=self.report_type_var, font=self.courier_font)
        report_type_label.grid(row=0, column=1, padx=150, pady=[20,10])

        self.time_var = StringVar(self.middle_frame, "")
        # time label
        time_label = Label(self.middle_frame, textvariable=self.time_var, font=self.courier_font)
        time_label.grid(row=1, column=1, padx=150)

        # sort utility
        self.sort_var = StringVar()
        self.sort_var.set("Sort")
        sort_options = ["Largest to Smallest", "Smallest to Largest"]
        sort_option_menu = OptionMenu(self.middle_frame, self.sort_var, *sort_options, command=self.sort_table_data)
        sort_option_menu.config(font=self.segoe_bold_font)
        sort_option_menu.grid(row=2, column=1, pady=10, sticky=E)
        

        # dataview
        dataview = Frame(self.middle_frame)
        dataview.grid(row=3, column=0, columnspan=3, pady=10)

        # treeview
        self.tree = ttk.Treeview(dataview, columns=('Column1', 'Column2'), style="mystyle.Treeview", height=15)
        self.tree.column('#0', anchor=W, width=100)
        self.tree.column('#1', anchor=W, width=450)
        self.tree.column('#2', anchor=W, width=100)
        self.tree.grid(row=0, column=0, sticky='nsew')

        self.tree.tag_configure('odd', background=GuidParameter.GRAY1)
        self.tree.tag_configure('even', background=GuidParameter.WHITE)
        self.tree.tag_configure('between', background = GuidParameter.GRAY2)

        # Constructing vertical scrollbar
        # with treeview
        verscrlbar = ttk.Scrollbar(dataview, 
                           orient ="vertical", 
                           command = self.tree.yview)
        verscrlbar.grid(row=0, column=1, sticky="ns")
        self.tree.config(xscrollcommand=verscrlbar.set)

        # table data
        self.table_data = dict() #place holder

        #Informative label (<middle_frame)
        self.informative_label_text = StringVar()

        informative_label = Label(self.middle_frame, textvariable=self.informative_label_text, font = self.segoe_small_font, foreground=GuidParameter.RED)
        informative_label.grid(row=4, column=1, pady=5, sticky=N)

        # save file button
        save_button = Button(self.middle_frame, text="Save as xlsx file", font=self.segoe_bold_font, command=self.open_save_file_win)
        save_button.grid(row=5, column=1, sticky=E)



    def insert_data_to_table(self, table_data):
        self.table_data = table_data

        for i in self.tree.get_children():
            self.tree.delete(i)

        id = 0
        for first_value in table_data:
            if not first_value == "Total": 
                if (id%2==0): tag = "even"
                else: tag = "odd"

                self.tree.insert(parent="", index=END, iid=str(id), text = str(id), values = (first_value, table_data[first_value]), tags=[tag])
                id+=1
        
        self.tree.insert(parent="", index=END, iid=str(id), text = str(id), values = ("Total", int(table_data["Total"])), tags=["between"])

    def sort_table_data(self, *args):
        if (len(self.table_data) == 0):
            self.informative_label_text.set("No data to sort.")
            return

        df = pd.DataFrame()
        for type in self.table_data:
            df = df.append({"Type": type, "Tickets": int(self.table_data[type])}, ignore_index=True)

        selected_option = self.sort_var.get()
        if selected_option == "Largest to Smallest":
            df = df.sort_values(by=["Tickets"], ascending=False)
        else:
            df = df.sort_values(by=["Tickets"], ascending=True)

        print(df)

        self.table_data = dict()
        for row in df.itertuples():
            self.table_data[row.Type] = int(row.Tickets)

        self.insert_data_to_table(self.table_data)

        self.informative_label_text.set("Sorting completed.")

    def open_save_file_win(self):
        save_dialog = SaveDialog(self.middle_frame, self.root, self.table_data, self.time_var.get(), self.report_type_var.get())
        