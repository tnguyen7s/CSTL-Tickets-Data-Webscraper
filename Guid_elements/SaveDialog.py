from posixpath import basename
from tkinter import Button, Entry, Label, StringVar, Toplevel
from tkinter.constants import W
from tkinter.filedialog import Directory, askdirectory, askopenfilename
from tkinter.font import BOLD, Font
from Enums.GuidParameter import GuidParameter
from Helpers.XlsxWriteHelper import XlsxWriteHelper
import json
import os
import random
import traceback

class SaveDialog():
    def __init__(self, frame, root, data, daterange, report_type):
        # The data
        self.data = data

        # date range of the data
        self.month_year = self.get_month_year_string(daterange)

        # the report type of the data
        self.report_type = report_type

        # list of months
        self.month_names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

        # open configuration file
        with open("configuration.json", "r") as config_file:
            self.config = json.load(config_file)

        # font
        courier_small_font = Font(root, family="Courier", size=14, weight=BOLD)
        segoe_small_font = Font(root, family="Segoe UI", size=10)
        segoe_bold_font = Font(root, family="Segoe UI", size=12, weight=BOLD)

        #win
        to_save_win = Toplevel(frame) 
        to_save_win.transient(frame)

        #save_to_new_file label
        save_to_new_file_label = Label(master=to_save_win, text="Save to a new file", font=courier_small_font)
        save_to_new_file_label.grid(column=0, columnspan=2, row=0, sticky=W, padx=10, pady=[10, 5])

        #choose directory button
        choose_dir_button = Button(master=to_save_win, text="Choose a folder...", font=segoe_small_font, command=self.open_ask_directory_dialog)
        choose_dir_button.grid(column=0, row=1, sticky=W, padx=[10, 2], pady=[0,10])

        # folder_path_var
        self.folder_path_var = StringVar(root)

        # folder name label
        folder_label = Label(to_save_win, textvariable=self.folder_path_var, font=segoe_small_font, foreground=GuidParameter.BLUE)
        folder_label.grid(column=1, row=1, columnspan=3, sticky=W)

        # file name entry
        self.file_name_input = Entry(master=to_save_win, textvariable=StringVar(to_save_win,"FileName.xlsx"), font=segoe_small_font)
        self.file_name_input.grid(column=0, row=2, columnspan=2, sticky=W, padx=10, pady=[0,10])

        # save_to_existing_file label
        save_to_existing_file_label = Label(master=to_save_win, text="Or save to an existing file", font=courier_small_font)
        save_to_existing_file_label.grid(column=0, columnspan=2, row=3, sticky=W, padx=10, pady=[10, 5])

        # choose_file button
        choose_file_button = Button(master=to_save_win, text="Choose a file...", font=segoe_small_font, command=self.open_ask_file_dialog)
        choose_file_button.grid(column=0, row=4, sticky=W, padx=[10, 2], pady=[0,10])

        # file_path_var
        self.file_path_var = StringVar(root)

        # file name label
        file_label = Label(to_save_win, textvariable=self.file_path_var, font=segoe_small_font, foreground=GuidParameter.BLUE)
        file_label.grid(column=1, row=4, columnspan=3, sticky=W)

        # informative_label
        self.information_var = StringVar(root)
        informative_label = Label(to_save_win, textvariable=self.information_var, font= segoe_small_font, foreground=GuidParameter.RED)
        informative_label.grid(row=6, column=0, columnspan=3, sticky=W, padx=10)

        #save button
        save_button = Button(to_save_win, text="Save", font=segoe_bold_font, command=self.save_to_xlsx_file)
        save_button.grid(row=5, column=1, pady=[0, 5])

        self.count = 0
        

    def open_ask_directory_dialog(self):
        selected_directory = askdirectory()
        self.folder_path_var.set(selected_directory)

    def open_ask_file_dialog(self):
        selected_file = askopenfilename(initialdir=self.config["DIRECTORY"], filetypes=[("excel file", ".xlsx")])
        self.file_path_var.set(selected_file)

    def save_to_xlsx_file(self):
        # check if data is present
        if len(self.data)==0:
            self.information_var.set("No data to save")
            return

        # Option 1: User may choose to save in a new file
        folder_path = self.folder_path_var.get() #default ""
        newly_created_file = self.file_name_input.get()
        newly_created_file_path = folder_path + "/" + newly_created_file

        # Option 2: User may choose to save in an existing file
        existing_file_path = self.file_path_var.get() #default ""
        
        # If user chose neither an existing file nor a folder_path to save new file-> not proceed and display message
        if (existing_file_path=="" and folder_path==""):
            self.information_var.set("Please specify the location to save.")
            return 0

        # User chose option 1
        if (not folder_path=="" and not newly_created_file==""):
            if (not self.check_valid_file_name(newly_created_file)):
                self.information_var.set("File name is invalid (extension must be xlsx).")
                return 0

            try:
                self.write_to_xlsx_file(newly_created_file_path, False)
            except Exception as e:
                print("{class_name}:{method}|{error}".format(class_name=self.__class__.__name__, method=self.save_to_xlsx_file.__name__, error=e))
                traceback.print_exc()
                return 0

            self.folder_path_var.set("")
            self.information_var.set("Save to {file} completed.".format(file=newly_created_file_path))
            return 1

        # User chose option 2:
        if (not existing_file_path==""):
            try:
                self.write_to_xlsx_file(existing_file_path, True)
            except Exception as e:
                print("{class_name}:{method}|{error}".format(class_name=self.__class__.__name__, method=self.save_to_xlsx_file.__name__, error=e))
                traceback.print_exc()
                return 0

            self.file_path_var.set("")
            self.information_var.set("Save to {file} completed.".format(file=existing_file_path))
            return 1

        # User specifies the folder_path but not file name
        if (newly_created_file==""):
            self.information_var.set("Please specify the filename.")
            return 0
        

    def write_to_xlsx_file(self, file_path, exist):
        writer = XlsxWriteHelper(file_path, exist=exist)
        writer.add_and_set_worksheet(self.month_year)
        chart_title = "{report_type} {time}".format(report_type=self.report_type, time=self.month_year)

        try:
            if (self.report_type == "A percentage breakdown by contact type"):
                writer.write_data_to_worksheet_no_color(self.data, self.month_year, 1)
            elif (self.report_type == "A percentage breakdown by request type"):
                writer.write_data_to_worksheet_no_color(self.data, self.month_year, 1)
            elif (self.report_type == "A percentage breakdown by Canvas type"):
                writer.write_data_to_worksheet_hightlight_orangebrown_top_5(self.data, self.month_year, 1)
            elif (self.report_type == "Support Ticket Assigned Tech"):
                tech_layer_start_row = writer.write_support_ticket_data_to_worksheet(self.data, 22)
                writer.create_pie_chart("A Percentage Breakdown of Who Is Answering Support Tickets", 2, tech_layer_start_row, tech_layer_start_row+3, 1, tech_layer_start_row, tech_layer_start_row+3, "A3")
            elif (self.report_type == "Top 10 faculty who consult with Kris Baranovic"):
                writer.write_data_to_worksheet_top_10_no_color(self.data, self.month_year, 1)
            elif (self.report_type == "Top 10 faculty who consult with Mary Harriet"):
                writer.write_data_to_worksheet_top_10_no_color(self.data, self.month_year, 1)
            elif (self.report_type == "Five most frequent issue handle by GA"):
                cursor = writer.write_data_to_worksheet_hightlight_yellow_top_5_and_fill_green(self.data, self.month_year, 1)
                writer.create_horizontal_bar_chart(chart_title, 3, 3, 1, cursor-1, 2, 1, cursor-1, "A"+str(cursor+1))

            elif (self.report_type == "Five most frequent issues handled by student workers"):
                cursor = writer.write_data_to_worksheet_hightlight_yellow_top_5_and_fill_green(self.data, self.month_year, 1)
                writer.create_horizontal_bar_chart(chart_title, 3, 3, 1, cursor-1, 2, 1, cursor-1, "A"+str(cursor+1))

        except Exception as e:
            writer.close_workbook_without_save()
            raise e

        writer.close_workbook()

    def check_valid_file_name(self, file_name):
        basename, ext = os.path.splitext(file_name)
        return ext==".xlsx"

    def get_month_year_string(self, daterange):
        month_names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]


        ls = daterange.split()[1].split("/")

        print(ls)
        month = int(ls[0])
        year = int(ls[2])

        return "{month_name}, {year}".format(month_name = month_names[month-1], year=year)

        
