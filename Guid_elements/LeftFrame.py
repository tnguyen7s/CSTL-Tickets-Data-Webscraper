# Import module
import json
import re
import traceback
from tkinter import Button, Canvas, Entry, Frame, Label, Listbox, PhotoImage, Scrollbar, StringVar, messagebox
from tkinter.constants import ANCHOR, BOTH, BOTTOM, CENTER, E, END, HORIZONTAL, LEFT, RIGHT, S, TOP, W
from tkinter.font import BOLD, Font
from pyodbc import IntegrityError

from Enums.GuidParameter import GuidParameter
from Enums.WorkerEnum import WorkerEnum
from Guid_elements.MiddleFrame import MiddleFrame
from Helpers.WHDDatabaseUpdater import WHDDatabaseUpdater
from Helpers.WHDScrapper import WHDScrapper

class LeftFrame:
    def __init__(self, root, tab, tab_name):
        with open("configuration.json", "r") as config_parser:
            self.config = json.load(config_parser)

        self.whd_scrapper = None
        if (tab_name=="update_tab"):
            self.whd_scrapper = WHDScrapper()
            #self.whd_scrapper.login(self.config["WHD"]["USERNAME"], self.config["WHD"]["PASSWORD"]) (if comment this, we disable the scrapping functionality)

        self.first = True

        self.database_updater = WHDDatabaseUpdater()
        self.data = []

        self.tab_name = tab_name
        # Font
        segoe_font = Font(root, family="Segoe UI", size=12)
        courier_font = Font(root, family="Courier", size=14, weight=BOLD)
        segoe_bold_font = Font(root, family="Segoe UI", size=12, weight=BOLD)
        segoe_small_font = Font(root, family="Segoe UI", size=10)

        # Left_frame (<update_tab)
        self.left_frame =Frame(tab, bg=GuidParameter.BLACK)
        self.left_frame.grid(column=0, row=0)

        # report_type_label (<left_frame)
        report_type_label = Label(self.left_frame, text= "CSTL WHD Ticket Report Type", font= courier_font, bg=GuidParameter.BLACK, fg = GuidParameter.WHITE)
        report_type_label.grid(row=0, column=0, padx = [10, 1], pady = 10)

        # icon_image = Image.open(GuidParameter.CATEGORY_ICON_FILE)
        # icon_image.resize((50, 50))

        # icon_label = Canvas(self.left_frame, width=50, height=50)
        # icon_label.create_image(0, 0, image=ImageTk.PhotoImage(icon_image))
        # icon_label.grid(row=0, column=1)

        #hor_ver_scroll_list (<left_frame)
        hor_ver_scroll_list = Frame(self.left_frame)
        hor_ver_scroll_list.grid(row=1, padx = 10, pady = [0, 10])

        ver_scrolled_list_box = Frame(hor_ver_scroll_list)
        ver_scrolled_list_box.grid(row=0)

        self.listbox = Listbox(ver_scrolled_list_box, width=(int)(GuidParameter.SCREEN_WIDTH/40), height=(int)(GuidParameter.SCREEN_HEIGHT/80), font=segoe_small_font)
        self.listbox.grid(column=0, row=0)

        values=["A percentage breakdown by contact type", 
                "A percentage breakdown by request type", 
                "A percentage breakdown by Canvas type", 
                "Support Ticket Assigned Tech", 
                "Top 10 faculty who consult with Kris Baranovic", 
                "Top 10 faculty who consult with Mary Harriet", 
                "Five most frequent issue handle by GA", 
                "Five most frequent issues handled by student workers"]
        for value in values:
            self.listbox.insert(END, value)


        # ver_scrollbar = Scrollbar(ver_scrolled_list_box)
        # ver_scrollbar.grid(column=1, row=0)
        # self.listbox.config(yscrollcommand=ver_scrollbar.set)
        # ver_scrollbar.config(command = self.listbox.yview)

        # hor_scrolllbar = Scrollbar(hor_ver_scroll_list, orient=HORIZONTAL)
        # hor_scrolllbar.grid(row=6)
        # self.listbox.config(xscrollcommand=hor_scrolllbar.set)
        # hor_scrolllbar.config(command=self.listbox.xview)

        # date_range_label (<left_frame)
        date_range_label = Label(self.left_frame, text="Date Range", font=courier_font, bg=GuidParameter.BLACK, fg = GuidParameter.WHITE)
        date_range_label.grid(row=2, sticky=W, padx=10, pady=10)

        # start_time_frame (<left_frame)
        start_time_frame = Frame(self.left_frame, bg=GuidParameter.BLACK)
        start_time_frame.grid(row=3, padx=10, pady=[0, 10])

        start_time_label = Label(start_time_frame, text = "Start Time", font=segoe_font, bg = GuidParameter.BLACK, fg = GuidParameter.WHITE)
        start_time_label.grid(row=0, column=0, sticky="w", padx=[0, 10])

        self.start_time_input = Entry(start_time_frame, font=segoe_small_font)
        self.start_time_input.insert(1, "10/01/2021")
        self.start_time_input.grid(row=0, column=1)

        # end_time_frame (<left_frame)
        end_time_frame = Frame(self.left_frame, bg=GuidParameter.BLACK)
        end_time_frame.grid(row = 4, padx=10, pady=[0, 10])

        end_time_label = Label(end_time_frame, text = "End Time  ", font = segoe_font, bg = GuidParameter.BLACK, fg = GuidParameter.WHITE)
        end_time_label.grid(row=0, column=0, sticky="w", padx=[0, 10])

        self.end_time_input = Entry(end_time_frame, font=segoe_small_font)
        self.end_time_input.insert(1, "11/01/2021")
        self.end_time_input.grid(row=0, column=1)

        # button_data_frame (<left_frame)
        button_text = StringVar()
        if (tab_name == 'update_tab'):
            button_text.set("Pull Data")
        elif (tab_name =='view_tab'):
            button_text.set("Get data")

        self.button = Button(self.left_frame, text=button_text.get(), font=segoe_bold_font)
        self.button.grid(row=5, sticky=E, padx=10, pady=10)

        #Informative label (<left_frame)
        self.informative_label_text = StringVar()

        informative_label = Label(self.left_frame, textvariable=self.informative_label_text, font = segoe_small_font, foreground=GuidParameter.RED, bg = GuidParameter.BLACK)
        informative_label.grid(row=6, pady=[10, 10])

        self.informative_label_text1 = StringVar()

        informative_label1 = Label(self.left_frame, textvariable=self.informative_label_text1, font = segoe_small_font, foreground=GuidParameter.RED, bg = GuidParameter.BLACK)
        informative_label1.grid(row=7, pady=[10, 180])


        # button command
        if (tab_name == "update_tab"):
            self.button.config(command=self.scrappe_whd_and_update_database)
        elif (tab_name == "view_tab"):
            self.button.config(command=self.get_data_from_the_database)

        # middle frame
        self.middle_frame = Frame() #placeholder

    def set_middle_frame(self, middle_frame):
        self.middle_frame = middle_frame

        if (self.tab_name == "update_tab"):
            self.middle_frame.report_type_var.set("This page is used to pull monthly ticket report data")
            self.middle_frame.time_var.set("from Web Help Desk")
        else:
            self.middle_frame.report_type_var.set("This page is used to get existing data")
            self.middle_frame.time_var.set("from SQL Server database")

    def scrappe_whd_and_update_database(self):
        report_type = self.listbox.get(ANCHOR)

        start_time = re.escape(self.start_time_input.get())
        end_time = re.escape(self.end_time_input.get())

        if not self.is_valid_date_input(start_time) or not self.is_valid_date_input(start_time):
            self.informative_label_text.set("Invalid date input.")
            return

        self.informative_label_text.set("Is Processing!")
        self.informative_label_text1.set("")
        try:
            if (report_type == "A percentage breakdown by contact type"):
                self.data = self.whd_scrapper.collect_tickets_data_breakdown_by_contact_type(start_time, end_time, self.first)
            elif (report_type == "A percentage breakdown by request type"):
                self.data = self.whd_scrapper.collect_tickets_data_breakdown_by_request_type(start_time, end_time, self.first)
            elif (report_type == "A percentage breakdown by Canvas type"):
                self.data = self.whd_scrapper.collect_tickets_data_breakdown_by_Canvas_type(start_time, end_time, self.first)
            elif (report_type == "Support Ticket Assigned Tech"):
                self.data = self.whd_scrapper.collect_support_ticket_assigned_tech(start_time, end_time, self.first)
            elif (report_type == "Top 10 faculty who consult with Kris Baranovic"):
                self.data = self.whd_scrapper.select_top_10_faculty_who_consult_with_Admin(start_time, end_time, WorkerEnum.KRIS_BARANOVIC, self.first)
            elif (report_type == "Top 10 faculty who consult with Mary Harriet"):
                self.data = self.whd_scrapper.select_top_10_faculty_who_consult_with_Admin(start_time, end_time, WorkerEnum.MARY_HARRIET, self.first)
            elif (report_type == "Five most frequent issue handle by GA"):
                self.data = self.whd_scrapper.collect_five_most_frequent_issues_handled_by_ga(start_time, end_time, self.first)
            elif (report_type == "Five most frequent issues handled by student workers"):
                self.data = self.whd_scrapper.collect_five_most_frequent_issues_handled_by_student_workers(start_time, end_time, self.first)
            else:
                self.informative_label_text.set("No report type specified in the request!")
                return
        except Exception as e:
            self.informative_label_text.set("Error occurs while scrapping data from WHD!")
            self.informative_label_text1.set("Logined again! Wait 5 minutes before next attempt!")
            self.first = False
            traceback.print_exc()
            return
        
        self.first = False
        self.informative_label_text.set("Pull data from WHD completed!")
        self.informative_label_text1.set("")
        
        self.middle_frame.report_type_var.set(report_type)
        self.middle_frame.time_var.set("From " + start_time + " To " + end_time)
        self.middle_frame.insert_data_to_table(self.data)

        # Ask "Do you wish to save this into your database"
        answer = messagebox.askquestion(message="Do you wish to save this into your SQL server database?", title="Database save")
        if (answer=="yes"):
            print("Yes, save to database!")
            self.save_data_table_to_SQL_database()

    def save_data_table_to_SQL_database(self):
        is_proceed, report_type, start_time, end_time = self.get_metadata_for_querying()
        if (not is_proceed):
            return

        self.informative_label_text.set("Sending queries to the database.")
        self.informative_label_text1.set("")
        try:
            if (report_type == "A percentage breakdown by contact type"):
                print("Here")
                self.database_updater.bulk_insert_tickets_data_breakdown_by_contact_type(start_time, end_time, self.data)
            elif (report_type == "A percentage breakdown by request type"):
                self.database_updater.bulk_insert_tickets_data_breakdown_by_request_type(start_time, end_time, self.data)
            elif (report_type == "A percentage breakdown by Canvas type"):
                self.database_updater.bulk_insert_tickets_data_breakdown_by_Canvas_type(start_time, end_time, self.data)
            elif (report_type == "Support Ticket Assigned Tech"):
                self.database_updater.bulk_insert_tickets_data_support_ticket_assigned_tech(start_time, end_time, self.data)
            elif (report_type == "Top 10 faculty who consult with Kris Baranovic"):
                self.database_updater.bulk_insert_tickets_data_top_10_faculty_who_consult_with_Admin(start_time, end_time, WorkerEnum.KRIS_BARANOVIC, self.data)
            elif (report_type == "Top 10 faculty who consult with Mary Harriet"):
                self.database_updater.bulk_insert_tickets_data_top_10_faculty_who_consult_with_Admin(start_time, end_time, WorkerEnum.MARY_HARRIET, self.data)
            elif (report_type == "Five most frequent issue handle by GA"):
                self.database_updater.bulk_insert_tickets_data_five_most_frequent_issues(start_time, end_time, WorkerEnum.GA, self.data)
            elif (report_type == "Five most frequent issues handled by student workers"):
                self.database_updater.bulk_insert_tickets_data_five_most_frequent_issues(start_time, end_time, WorkerEnum.STUDENT_WORKER, self.data)
        except IntegrityError as e:
            print(e)
            answer = messagebox.askquestion(message="Table data has already been saved. Do you want to replace it?", title="Database replacement")
            if (answer == 'yes'):
                self.delete_data_from_SQL_database()
                self.save_data_table_to_SQL_database()
            return
        except:
            self.informative_label_text.set("Error thrown while doing the save operation.")
            self.informative_label_text1.set("")
            traceback.print_exc()
            return

        self.informative_label_text.set("Save to SQL server database completed.")
        self.informative_label_text1.set("")

    def delete_data_from_SQL_database(self):
        is_proceed, report_type, start_time, end_time = self.get_metadata_for_querying()
        if (not is_proceed):
            return 

        self.informative_label_text.set("Deleting data table from database.")
        self.informative_label_text1.set("")

        try:
            if (report_type == "A percentage breakdown by contact type"):
                self.database_updater.delete_tickets_data_breakdown_by_contact_type(start_time, end_time)
            elif (report_type == "A percentage breakdown by request type"):
                self.database_updater.delete_tickets_data_breakdown_by_request_type(start_time, end_time)
            elif (report_type == "A percentage breakdown by Canvas type"):
                self.database_updater.delete_tickets_data_breakdown_by_Canvas_type(start_time, end_time)
            elif (report_type == "Support Ticket Assigned Tech"):
                self.database_updater.delete_tickets_data_support_ticket_assigned_tech(start_time, end_time)
            elif (report_type == "Top 10 faculty who consult with Kris Baranovic"):
                self.database_updater.delete_tickets_data_top_10_faculty_who_consult_with_Admin(start_time, end_time, WorkerEnum.KRIS_BARANOVIC)
            elif (report_type == "Top 10 faculty who consult with Mary Harriet"):
                self.database_updater.delete_tickets_data_top_10_faculty_who_consult_with_Admin(start_time, end_time, WorkerEnum.MARY_HARRIET)
            elif (report_type == "Five most frequent issue handle by GA"):
                self.database_updater.delete_tickets_data_five_most_frequent_issues(start_time, end_time, WorkerEnum.GA)
            elif (report_type == "Five most frequent issues handled by student workers"):
                self.database_updater.delete_tickets_data_five_most_frequent_issues(start_time, end_time, WorkerEnum.STUDENT_WORKER)
        except:
            self.informative_label_text.set("Error thrown while doing database delete operation.")
            traceback.print_exc()
            return
        
        self.informative_label_text.set("Delete the table data from database completed.")
        self.informative_label_text1.set("")

    def get_data_from_the_database(self):
        is_proceed, report_type, start_time, end_time = self.get_metadata_for_querying()
        if (not is_proceed):
            return 

        self.informative_label_text.set("Is Processing!")
        self.informative_label_text1.set("")

        try:
            if (report_type == "A percentage breakdown by contact type"):
                self.data = self.database_updater.read_tickets_data_breakdown_by_contact_type(start_time, end_time)
            elif (report_type == "A percentage breakdown by request type"):
                self.data = self.database_updater.read_tickets_data_breakdown_by_request_type(start_time, end_time)
            elif (report_type == "A percentage breakdown by Canvas type"):
                self.data = self.database_updater.read_tickets_data_breakdown_by_Canvas_type(start_time, end_time)
            elif (report_type == "Support Ticket Assigned Tech"):
                self.data = self.database_updater.read_tickets_data_support_ticket_assigned_tech(start_time, end_time)
            elif (report_type == "Top 10 faculty who consult with Kris Baranovic"):
                self.data = self.database_updater.read_tickets_data_top_10_faculty_who_consult_with_Admin(start_time, end_time, WorkerEnum.KRIS_BARANOVIC)
            elif (report_type == "Top 10 faculty who consult with Mary Harriet"):
                self.data = self.database_updater.read_tickets_data_top_10_faculty_who_consult_with_Admin(start_time, end_time, WorkerEnum.MARY_HARRIET)
            elif (report_type == "Five most frequent issue handle by GA"):
                self.data = self.database_updater.read_tickets_data_five_most_frequent_issues(start_time, end_time, WorkerEnum.GA)
            elif (report_type == "Five most frequent issues handled by student workers"):
                self.data = self.database_updater.read_tickets_data_five_most_frequent_issues(start_time, end_time, WorkerEnum.STUDENT_WORKER)
            else:
                self.informative_label_text.set("No report type specified in the request!")
                self.informative_label_text1.set("")
                return
        except Exception as e:
            self.informative_label_text.set("Error thrown while doing database retrieval operation.")
            self.informative_label_text1.set("")
            traceback.print_exc()
            return
        
        if (len(self.data)==0): 
            self.informative_label_text.set("No data found in the database.")
            self.informative_label_text1.set("")
            return

        self.informative_label_text.set("Pull existing data from the database completed!")
        self.informative_label_text1.set("")
        self.middle_frame.report_type_var.set(report_type)
        self.middle_frame.time_var.set("From " + self.start_time_input.get() + " To " + self.end_time_input.get())
        self.middle_frame.insert_data_to_table(self.data)


    """
    Check if the date input is valid.
    """
    def is_valid_date_input(self, start_time):
       
        # check the pattern
        parts = start_time.split("/")
        if (len(parts)!=3):
            return False
            
        dd = parts[1]
        mm = parts[0]
        yyyy = parts[2]

        # Check dgits
        if(re.search("[^0-9]", dd) !=None or re.search("[^0-9]", mm) !=None or re.search("[^0-9]", yyyy) != None):
            return False

        # check valid day partially, valid month, and valid year
        if (int(dd)<1 or int(mm)<1 or int(mm)>12 or int(yyyy)<2000):
            return False
            
        leap_year = False

        if (int(yyyy)%400==0 or int(yyyy)%4==0):
            leap_year = True

        if (leap_year and int(mm)==2 and int(dd)>29):
            return False
            

        number_of_days = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

        if (not leap_year and int(dd)>number_of_days[int(mm)]):
            return False

        return True

    def get_metadata_for_querying(self):
        report_type = self.listbox.get(ANCHOR)

        start_time = self.start_time_input.get()
        end_time = self.end_time_input.get()

        if not self.is_valid_date_input(start_time) or not self.is_valid_date_input(start_time):
            self.informative_label_text.set("Invalid date input.")
            return False, None, None, None
        
        start_time_parts = start_time.split("/")
        start_time = "{yyyy}-{mm}-{dd}".format(yyyy=start_time_parts[2], mm = start_time_parts[0], dd=start_time_parts[1])

        end_time_parts = end_time.split("/")
        end_time = "{yyyy}-{mm}-{dd}".format(yyyy=end_time_parts[2], mm = end_time_parts[0], dd=end_time_parts[1])

        return True, report_type, start_time, end_time

    def close_database_connection_and_close_whd(self):
        #self.database_updater.close_connection()
        if not self.whd_scrapper:
            self.whd_scrapper.close_whd()