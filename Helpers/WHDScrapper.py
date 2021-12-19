import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
import json

from Enums.WorkerEnum import WorkerEnum

class WHDScrapper:
    """
    WHDScrapper's function is to collect Web Help Desk tickets data (first: first method is being called => choose correct report button)
    """
    def __init__(self):
        # geckodriver's path needs to be in the path system 
        self.driver = webdriver.Firefox()
        self.wait = WebDriverWait(self.driver, 600)

        self.was_in_the_second_page = False

        with open("configuration.json", "r") as config_parser:
            self.config = json.load(config_parser)

    def login(self, username, password):
        # Open the web help desk page
        self.driver.get(self.config['WHD']['URL'])

        # Assert the title of the web page
        assert self.config['WHD']['TITLE'] in self.driver.title

        # Maximize the window of the web page
        self.driver.maximize_window()

        # enter username
        username_input_box = self.wait.until(EC.presence_of_element_located((By.ID, "userName")))
        username_input_box.send_keys(username)

        # enter password
        pwd_input_box = self.wait.until(EC.presence_of_element_located((By.ID, "password")))
        pwd_input_box.send_keys(password)

        # submit the form
        submit_button = self.wait.until(EC.presence_of_element_located((By.CLASS_NAME, "formLoginButton")))
        submit_button.submit()

    def collect_tickets_data_breakdown_by_contact_type(self, start_time, end_time, first = False):
        print("-----------START: " + self.collect_tickets_data_breakdown_by_contact_type.__name__)

        try:
            # Click on report button
            self.click_shared_report_button(first)

            # Select on the "A percentage breakdown by contact type" link tag
            selected_report = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['LINK_TAG']['BY_CONTACT_TYPE_REPORT_SHADED'])))
            selected_report.click()

            # enter time range and click on run report
            self.enter_time_range_and_run_report(start_time, end_time)

            # find the table data and read
            retrieved_data = self.find_table_data_and_read()
        except Exception as e:
            self.handle_error()
            raise

        print("-----------END: " + self.collect_tickets_data_breakdown_by_contact_type.__name__)
        print(retrieved_data)
        print()
        return retrieved_data
        

    def collect_tickets_data_breakdown_by_request_type(self, start_time, end_time, first = False): 
        print("-----------START: " + self.collect_tickets_data_breakdown_by_request_type.__name__)

        try: 
            # Click on report button
            self.click_shared_report_button(first)

            # Select on the "A percentage breakdown by request type" link tag
            selected_report = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['LINK_TAG']['BY_REQUEST_TYPE_REPORT_WHITE'])))
            selected_report.click()
            
            # enter time range and click on run report
            self.enter_time_range_and_run_report(start_time, end_time)

            #find the table data and read
            retrieved_data = self.find_table_data_and_read()
        except Exception as e:
            self.handle_error()
            raise e

        print("-----------END: " + self.collect_tickets_data_breakdown_by_request_type.__name__)
        print(retrieved_data)
        print()

        return retrieved_data

    def collect_tickets_data_breakdown_by_Canvas_type(self, start_time, end_time, first=False):
        print("-----------START: " + self.collect_tickets_data_breakdown_by_Canvas_type.__name__)

        try:
            # Click on report button
            self.click_shared_report_button(first)

            # Select on the "A percentage breakdown by Canvas type" link tag
            selected_report = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['LINK_TAG']['BY_CANVAS_TYPE_REPORT_WHITE'])))
            selected_report.click()

            # enter time range and click on run report
            self.enter_time_range_and_run_report(start_time, end_time)

            #find the table data and read
            retrieved_data = self.find_table_data_and_read()
        except Exception as e:
            self.handle_error()
            raise e

        print("-----------END: " + self.collect_tickets_data_breakdown_by_contact_type.__name__)
        print(retrieved_data)
        print()
        return retrieved_data


    def collect_five_most_frequent_issues(self, start_time, end_time, handle_by, first=False):
        print("-----------START: " + self.collect_five_most_frequent_issues.__name__)

        try:
            # Click on report button
            self.click_shared_report_button(first)

            # Select on the "Five most frequent issues handled by Student Workers" link tag
            if (not self.was_in_the_second_page):
                # in case we are in first page and need to switch to next page to select the tag
                next_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['NEXT'])))
                next_button.click()
                selected_report = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['LINK_TAG']['FIVE_MOST_FREQUENT_ISSUES_REPORT_WHITE'])))
                selected_report.click()
            else:
                selected_report = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['LINK_TAG']['FIVE_MOST_FREQUENT_ISSUES_REPORT_SHADED'])))
                selected_report.click()

            # Click on report filter button
            report_filter_bt = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['REPORT_FILTER'])))
            report_filter_bt.click()

            # Click on assigned tech button 
            assigned_tech_bt = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['ASSIGNED_TECH'])))
            assigned_tech_bt.click()
             # Click on delete button
            delete_bt = self.wait.until((EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['DELETE']))))
            delete_bt.click()
            # handle alert "really delete this filter"
            self.wait.until(EC.alert_is_present())
            self.driver.switch_to_alert().accept()

            # Click on new button
            new_bt = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['NEW'])))
            new_bt.click()

            # select the assigned tech option from the drop down menu
            filter_option = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['DROP_DOWN_MENU']['FILTER_ATTRIBUTE']['ASSIGNED_TECH'])))
            filter_option.click()

            #select the tech handles the issues
            assigned_tech_check_box = list()
            if (handle_by == WorkerEnum.STUDENT_WORKER):
                assigned_tech_check_box.append(self.config['HTML_ELEMENT']['CHECK_BOX']['STUDENT_WORKER_1'])
                assigned_tech_check_box.append(self.config['HTML_ELEMENT']['CHECK_BOX']['STUDENT_WORKER_2'])
            elif (handle_by == WorkerEnum.GA):
                assigned_tech_check_box.append(self.config['HTML_ELEMENT']['CHECK_BOX']['GA'])
            else:
                raise ValueError("Not implement yet! Please add a case for it in " + self.collect_five_most_frequent_issues.__name__)
            
            for tech in assigned_tech_check_box: 
                selected_tech = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, tech)))
                selected_tech.click()
        
            #click on save button
            save_bt = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['SAVE'])))
            save_bt.click()

            #click on done button
            done_bt = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['DONE'])))
            done_bt.click()

            # Select on the "Five most frequent issues handled by Student Workers" link tag
            selected_report = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['LINK_TAG']['FIVE_MOST_FREQUENT_ISSUES_REPORT_SHADED'])))
            selected_report.click()

            # enter the time range and click on run report
            self.enter_time_range_and_run_report(start_time, end_time)

            #find the table data and read
            retrieved_data = self.find_table_data_and_read()
        except Exception as e:
            self.handle_error()
            raise e

        print("-----------END: " + self.collect_five_most_frequent_issues.__name__)
        print(retrieved_data)
        print()

        self.was_in_the_second_page = True
        return retrieved_data

    def collect_five_most_frequent_issues_handled_by_student_workers(self, start_time, end_time, first=False):
        try:
            return self.collect_five_most_frequent_issues(start_time, end_time, WorkerEnum.STUDENT_WORKER, first)
        except Exception as e:
            raise e

    def collect_five_most_frequent_issues_handled_by_ga(self, start_time, end_time, first=False):
        try:
            return self.collect_five_most_frequent_issues(start_time, end_time, WorkerEnum.GA, first)
        except Exception as e:
            raise e

    def collect_support_ticket_assigned_tech(self, start_time, end_time, first=False):
        print("-----------START: " + self.collect_support_ticket_assigned_tech.__name__)
        
        try:
            # Click on report button
            self.click_shared_report_button(first)
            
            # Select on the "Support Ticket Assigned Tech" link tag
            if (not self.was_in_the_second_page):
                # in case we are in first page and need to switch to next page to select the tag
                next_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['NEXT'])))
                next_button.click()
                selected_report = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['LINK_TAG']['SUPPORT_TICKET_ASSIGNED_TECH_REPORT_SHADED'])))
                selected_report.click()
            else:
                selected_report = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['LINK_TAG']['SUPPORT_TICKET_ASSIGNED_TECH_REPORT_WHITE'])))
                selected_report.click()

            # enter time range and click on run report
            self.enter_time_range_and_run_report(start_time, end_time)

            #find the table data and read
            retrieved_data = self.find_table_data_and_read()
        except Exception as e:
            self.handle_error()
            raise e

        print("-----------END: " + self.collect_support_ticket_assigned_tech.__name__)
        print(retrieved_data)
        print()
        self.was_in_the_second_page = True

        return retrieved_data

    def select_top_10_faculty_who_consult_with_Admin(self, start_time, end_time, consult_by, first = False):
        print("-----------START: " + self.select_top_10_faculty_who_consult_with_Admin.__name__)
        
        try:
            # Click on report button
            self.click_shared_report_button(first)

            # Select on the "TOP 10 faculty consult with CSTL administrator April 2020" link tag
            if (not self.was_in_the_second_page):
                # in case we are in first page and need to switch to next page to select the tag
                next_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['NEXT'])))
                next_button.click()
                selected_report = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['LINK_TAG']['TOP_10_FACULTY_WHO_CONSULT_WITH_REPORT_SHADED'])))
                selected_report.click()
            else:
                selected_report = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['LINK_TAG']['TOP_10_FACULTY_WHO_CONSULT_WITH_REPORT_WHITE'])))
                selected_report.click()

            # Click on report filter button 
            report_filter_bt = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['REPORT_FILTER'])))
            report_filter_bt.click()

         # Click on assigned tech button 
            assigned_tech_bt = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['ASSIGNED_TECH']))) 
            assigned_tech_bt.click()
            # Click on delete button
            delete_bt = self.wait.until((EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['DELETE'])))) 
            delete_bt.click()
            # handle alert "really delete this filter"
            self.wait.until(EC.alert_is_present())
            self.driver.switch_to_alert().accept()

            # Click on new button
            new_bt = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['NEW'])))
            new_bt.click()

            # select the assigned tech option from the drop down menu
            filter_option = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['DROP_DOWN_MENU']['FILTER_ATTRIBUTE']['ASSIGNED_TECH'])))
            filter_option.click()

            #select the tech handles the issues
            assigned_tech_check_box = list()
            if (consult_by == WorkerEnum.KRIS_BARANOVIC):
                assigned_tech_check_box.append(self.config['HTML_ELEMENT']['CHECK_BOX']['KRIS'])
            elif (consult_by == WorkerEnum.MARY_HARRIET):
                assigned_tech_check_box.append(self.config['HTML_ELEMENT']['CHECK_BOX']['MARY'])
            else:
                raise ValueError("Not implement yet! Please add a case for it in " + self.collect_support_ticket_assigned_tech.__name__)
            
            for tech in assigned_tech_check_box: 
                selected_tech = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, tech)))
                selected_tech.click()
        
            #click on save button
            save_bt = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['SAVE'])))
            save_bt.click()

            #click on done button
            done_bt = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['DONE'])))
            done_bt.click()

            # Select on the "TOP 10 faculty consult with CSTL administrator April 2020" link tag
            selected_report = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,self.config['HTML_ELEMENT']['LINK_TAG']['TOP_10_FACULTY_WHO_CONSULT_WITH_REPORT_WHITE'])))
            selected_report.click()

            # enter time range and click on run report
            self.enter_time_range_and_run_report(start_time, end_time)

            #find the table data and read
            retrieved_data = self.find_table_data_and_read()
        except Exception as e:
            self.handle_error()
            raise e

        print("-----------END: " + self.select_top_10_faculty_who_consult_with_Admin.__name__)
        print(retrieved_data)
        print()

        self.was_in_the_second_page = True
        return retrieved_data

    def click_shared_report_button(self, first):
        if first: selector = self.config['HTML_ELEMENT']['BUTTON']['REPORT']
        else: selector = self.config['HTML_ELEMENT']['BUTTON']['SELECTED_REPORT']

        reportButton = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
        reportButton.click()


    def enter_time_range_and_run_report(self, start_time, end_time):
        #To handle StaleElementReferenceException: https://club.ministryoftesting.com/t/webdriver-stale-either-the-element-is-no-longer-attached-to-the-dom-or-the-page-has-been-refreshed/11347/4
# Input FROM and TO
        attempt= True
        while (attempt):
            try:
                from_input_box = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['INPUT_BOX']['FROM'])))
                from_input_box.clear()
                from_input_box.send_keys(start_time)
                        
                to_input_box = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['INPUT_BOX']['TO'])))
                to_input_box.clear()
                to_input_box.send_keys(end_time)

                run_report_button = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['BUTTON']['RUN_REPORT'])))
                run_report_button.click()
                attempt = False
            except Exception as e:
                print("{class_name}:{method}|{error}".format(class_name=self.__class__.__name__, method=self.enter_time_range_and_run_report.__name__, error=e))
                time.sleep(10)
    
        #self.driver.refresh() (does not work well, result in misleading data result)
        #*** REMEMEBER: the above approach work, since elements are on the same page, attempt the loop (no need to refresh, since it was shown that itworked without refreshing)
        #***If data is misleading, call the method again and it will work in the second call
    
    def find_table_data_and_read(self):
        tds = self.wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, self.config['HTML_ELEMENT']['TABLE'])))
        
        tmp_list = list()
        for td in tds:
            tmp_list.append(td.text)

        print(tmp_list)

        retrieved_data= dict()
        for i in range(0, len(tmp_list), 2):
            retrieved_data[tmp_list[i]] = int(tmp_list[i+1])
        
        return retrieved_data
    
    def close_whd(self):
        self.driver.close()

    def handle_error(self):
        time.sleep(30)
        # close the current one
        self.close_whd()

        #open new whd windown 
        self.driver = webdriver.Firefox()
        self.wait = WebDriverWait(self.driver, 600)

        self.was_in_the_second_page = False
        self.login(self.config["WHD"]["USERNAME"], self.config["WHD"]["PASSWORD"])

        





        




