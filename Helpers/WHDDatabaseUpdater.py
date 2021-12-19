import pyodbc
import json
class WHDDatabaseUpdater:
    """
    WHDDatabaseUpdater's function is to update the database with the ticket report data.
    The data in the database then used by the web app to intepret in a nice format.
    """
    def __init__(self):
        with open("configuration.json") as config_file:
            self.sql_connection_string = json.load(config_file)["DATABASE"]["SQL_CONNECTION_STRING"]

    def insert_report_time(self, start_time, end_time):
        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor = sql_connection.cursor()

        sql_query = "INSERT INTO ReportTime VALUES(?, ?)"
        with sql_connection:
            try:
                cursor.execute(sql_query, start_time, end_time)
                cursor.commit()
            except Exception as e: 
                cursor.close()
                return
        
        cursor.close()


    def bulk_insert_tickets_data_breakdown_by_contact_type(self, start_time, end_time, data):
        print(self.bulk_insert_tickets_data_breakdown_by_contact_type.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()
        
        sql_query = "INSERT INTO BreakdownByContactType VALUES"
        for contact_type in data:
            sql_query += "('{start}', '{end}', '{type}', {tickets}),".format(start=start_time, end=end_time, type=contact_type, tickets = data[contact_type])

        sql_query= sql_query[0:len(sql_query)-1] + ";"

        with sql_connection:
            try:
                cursor.execute(sql_query)
                cursor.commit()
            except Exception as e:
                cursor.close()
                raise e
            cursor.close()
        
        self.insert_report_time(start_time, end_time)
    
    def read_tickets_data_breakdown_by_contact_type(self, start_time, end_time):
        print(self.read_tickets_data_breakdown_by_contact_type.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()

        sql_query = "SELECT * FROM BreakdownByContactType WHERE StartTime = ? AND EndTime = ?;"
        with sql_connection:
            try:
                cursor = cursor.execute(sql_query, start_time, end_time)
                result = cursor.fetchall()
            except Exception as e:
                print(e)
                cursor.close()
                raise e

        data = dict()
        for d in result:
            data[d[2]] = int(d[3])

        cursor.close()

        return data

    def delete_tickets_data_breakdown_by_contact_type(self, start_time, end_time):
        print(self.delete_tickets_data_breakdown_by_contact_type.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()

        sql_query = "DELETE FROM BreakdownByContactType WHERE StartTime = ? AND EndTime = ?;"

        with sql_connection:
            try:
                sql_connection.autocommit = False
                cursor.execute(sql_query, start_time, end_time)
            except pyodbc.DatabaseError as e:
                cursor.rollback()
                cursor.close()
                raise e
            else:
                cursor.commit()
                
            finally:
                sql_connection.autocommit = True

        cursor.close()

    def bulk_insert_tickets_data_breakdown_by_request_type(self, start_time, end_time, data):
        print(self.bulk_insert_tickets_data_breakdown_by_request_type.__name__)
        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()
        
        sql_query = "INSERT INTO BreakdownByRequestType VALUES"
        for request_type in data:
            sql_query += "('{start}', '{end}', '{type}', {tickets}),".format(start=start_time, end=end_time, type=request_type, tickets = data[request_type])

        sql_query= sql_query[0:len(sql_query)-1] + ";"

        with sql_connection:
            try:
                cursor.execute(sql_query)
                cursor.commit()
            except Exception as e:
                cursor.close()
                raise e

        cursor.close()
        self.insert_report_time(start_time, end_time)
                  
    def read_tickets_data_breakdown_by_request_type(self, start_time, end_time):
        print(self.read_tickets_data_breakdown_by_request_type.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()

        sql_query = "SELECT * FROM BreakdownByRequestType WHERE StartTime = ? AND EndTime = ?;"

        with sql_connection:
            try:
                cursor = cursor.execute(sql_query, start_time, end_time)
                result = cursor.fetchall()
            except Exception as e:
                cursor.close()
                raise e

            data = dict()
        for d in result:
            data[d[2]] = int(d[3])

        cursor.close()
        
        return data

    def delete_tickets_data_breakdown_by_request_type(self, start_time, end_time):
        print(self.delete_tickets_data_breakdown_by_request_type.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()

        sql_query = "DELETE FROM BreakdownByRequestType WHERE StartTime = ? AND EndTime = ?;"

        with sql_connection:
            try:
                sql_connection.autocommit = False
                cursor.execute(sql_query, start_time, end_time)
            except pyodbc.DatabaseError as e:
                cursor.rollback()
                cursor.close()
                raise e
            else:
                cursor.commit()
                
            finally:
                sql_connection.autocommit = True

        cursor.close()

    def bulk_insert_tickets_data_five_most_frequent_issues(self, start_time, end_time, handle_by, data):
        print(self.bulk_insert_tickets_data_five_most_frequent_issues.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()
        
        sql_query = "INSERT INTO CstlFrequentIssues VALUES"
        for issue_type in data:
            sql_query += "('{start}', '{end}', '{type}', {handle_by}, {tickets}),".format(start=start_time, end=end_time, type=issue_type, handle_by=handle_by, tickets=data[issue_type])

        sql_query= sql_query[0:len(sql_query)-1] + ";"

        with sql_connection:
            try:
                cursor.execute(sql_query)
                cursor.commit()
            except Exception as e:
                cursor.close()
                raise e

        cursor.close() 
        self.insert_report_time(start_time, end_time)

    def read_tickets_data_five_most_frequent_issues(self, start_time, end_time, handle_by):
        print(self.read_tickets_data_five_most_frequent_issues.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()

        sql_query = "SELECT * FROM CstlFrequentIssues WHERE StartTime = ? AND EndTime = ? AND HandleByWorkerEnum = ?;"

        with sql_connection:
            try:
                cursor = cursor.execute(sql_query, start_time, end_time, handle_by)
                result=cursor.fetchall()
            except Exception as e:
                cursor.close()
                raise e

        data = dict()
        for d in result:
            data[d[2]] = int(d[4]) 

        cursor.close() 
        
        return data

    def delete_tickets_data_five_most_frequent_issues(self, start_time, end_time, handle_by):
        print(self.delete_tickets_data_five_most_frequent_issues.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()

        sql_query = "DELETE FROM CstlFrequentIssues WHERE StartTime = ? AND EndTime = ? AND HandleByWorkerEnum = ?;"

        with sql_connection:
            try:
                sql_connection.autocommit = False
                cursor.execute(sql_query, start_time, end_time, handle_by)
            except pyodbc.DatabaseError as e:
                cursor.rollback()
                cursor.close()
                raise e
            else:
                cursor.commit()
            finally:
                sql_connection.autocommit = True

        cursor.close() 

    def bulk_insert_tickets_data_support_ticket_assigned_tech(self, start_time, end_time, data):
        print(self.bulk_insert_tickets_data_support_ticket_assigned_tech.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()
        
        sql_query = "INSERT INTO SupportTicketAssignedTech VALUES"
        for tech in data:
            sql_query += "('{start}', '{end}', '{tech}', {tickets}),".format(start=start_time, end=end_time, tech=tech, tickets=data[tech])

        sql_query= sql_query[0:len(sql_query)-1] + ";"

        with sql_connection:
            try:
                cursor.execute(sql_query)
                cursor.commit()
            except Exception as e:
                cursor.close()
                raise e

        cursor.close() 
        self.insert_report_time(start_time, end_time)

    def read_tickets_data_support_ticket_assigned_tech(self, start_time, end_time):
        print(self.read_tickets_data_support_ticket_assigned_tech.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()

        sql_query = "SELECT * FROM SupportTicketAssignedTech WHERE StartTime = ? AND EndTime = ?;"

        with sql_connection:
            try:
                cursor = cursor.execute(sql_query, start_time, end_time)
                result = cursor.fetchall()
            except Exception as e:
                cursor.close()
                raise e

        data = dict()
        for d in result:
            data[d[2]] = int(d[3]) 
            print(d)

        cursor.close() 
        
        return data

    def delete_tickets_data_support_ticket_assigned_tech(self, start_time, end_time):
        print(self.delete_tickets_data_support_ticket_assigned_tech.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()

        sql_query = "DELETE FROM SupportTicketAssignedTech WHERE StartTime = ? AND EndTime = ?;"

        with sql_connection:
            try:
                sql_connection.autocommit = False
                cursor.execute(sql_query, start_time, end_time)
            except pyodbc.DatabaseError as e:
                cursor.rollback()
                cursor.close()
                raise e
            else:
                cursor.commit()
            finally:
                sql_connection.autocommit = True

        cursor.close() 
    
    def bulk_insert_tickets_data_top_10_faculty_who_consult_with_Admin(self, start_time, end_time, consult_by, data):
        print(self.bulk_insert_tickets_data_top_10_faculty_who_consult_with_Admin.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()
        
        sql_query = "INSERT INTO CSTLConsult VALUES"
        for faculty in data:
            sql_query += "('{start}', '{end}', '{faculty}', {consult_by}, {consult_times}),".format(start=start_time, end=end_time, faculty=faculty.replace("'", ""), consult_by=consult_by, consult_times=data[faculty])

        sql_query= sql_query[0:len(sql_query)-1] + ";"

        with sql_connection:
            try:
                cursor.execute(sql_query)
                cursor.commit()
            except Exception as e:
                cursor.close()
                raise e

        cursor.close() 
        self.insert_report_time(start_time, end_time)
    
    def read_tickets_data_top_10_faculty_who_consult_with_Admin(self, start_time, end_time, consult_by):
        print(self.read_tickets_data_top_10_faculty_who_consult_with_Admin.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()

        sql_query = "SELECT * FROM CSTLConsult WHERE StartTime = ? AND EndTime = ? AND ConsultByWorkerEnum = ?;"

        with sql_connection:
            try:
                cursor = cursor.execute(sql_query, start_time, end_time, consult_by)
                result = cursor.fetchall()
            except Exception as e:
                cursor.close()
                raise e

        data = dict()
        for d in result:
            data[d[2]] = int(d[4]) 

        cursor.close() 

        return data

    def delete_tickets_data_top_10_faculty_who_consult_with_Admin(self, start_time, end_time, consult_by):
        print(self.delete_tickets_data_top_10_faculty_who_consult_with_Admin.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()

        sql_query = "DELETE FROM CSTLConsult WHERE StartTime = ? AND EndTime = ? AND ConsultByWorkerEnum = ?;"

        with sql_connection:
            try:
                sql_connection.autocommit = False
                cursor.execute(sql_query, start_time, end_time, consult_by)
            except pyodbc.DatabaseError as e:
                cursor.rollback()
                cursor.close()
                raise e
            else:
                cursor.commit()
            finally:
                sql_connection.autocommit = True

        cursor.close() 
        

    def bulk_insert_tickets_data_breakdown_by_Canvas_type(self, start_time, end_time, data):
        print(self.bulk_insert_tickets_data_breakdown_by_Canvas_type.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()
        
        sql_query = "INSERT INTO BreakdownByCanvasType VALUES"
        for canvas_type in data:
            sql_query += "('{start}', '{end}', '{type}', {tickets}),".format(start=start_time, end=end_time, type=canvas_type, tickets = data[canvas_type])
            print(sql_query)
        sql_query= sql_query[0:len(sql_query)-1] + ";"

        with sql_connection:
            try:
                cursor.execute(sql_query)
                cursor.commit()
            except Exception as e:
                cursor.close()
                raise e
        cursor.close()
        self.insert_report_time(start_time, end_time)
    
    def read_tickets_data_breakdown_by_Canvas_type(self, start_time, end_time):
        print(self.read_tickets_data_breakdown_by_Canvas_type.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()

        sql_query = "SELECT * FROM BreakdownByCanvasType WHERE StartTime = ? AND EndTime = ?;"

        with sql_connection:
            try:
                cursor = cursor.execute(sql_query, start_time, end_time)
                result = cursor.fetchall()
            except Exception as e:
                cursor.close()
                raise e
        
        data = dict()
        for d in result:
            data[d[2]] = int(d[3]) 
        print(data)

        cursor.close()

        return data

    def delete_tickets_data_breakdown_by_Canvas_type(self, start_time, end_time):
        print(self.delete_tickets_data_breakdown_by_Canvas_type.__name__)

        sql_connection = pyodbc.connect(self.sql_connection_string)
        cursor =  sql_connection.cursor()

        sql_query = "DELETE FROM BreakdownByCanvasType WHERE StartTime = ? AND EndTime = ?;"

        with sql_connection:
            try:
                sql_connection.autocommit = False
                cursor.execute(sql_query, start_time, end_time)
            except pyodbc.DatabaseError as e:
                cursor.rollback()
                cursor.close()
                raise e
            else:
                cursor.commit()
            finally:
                sql_connection.autocommit = True
            
        cursor.close()

