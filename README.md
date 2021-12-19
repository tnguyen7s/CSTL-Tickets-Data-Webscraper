# CSTL-Tickets-Data-Webscraper
## Project Description
- In this project, I utilized Selenium packages to scrape data from the Web Help Desk website to collect monthly ticket data replacing the manual method and automate the monthly task at my working place (CTL). 
- The application provides a UI using Tkinter packages. The UI allows users to give inputs to the scraper and to view and insert the collected data to the SQL Server database.
- The limitation of the app that needs to be improved in the future: it is quite slow from the scrapping work and sometimes the application crashes due to the fact that the Scraper could not find the UI element to select. It also requires constant updates because sometimes the CSS Selector gets changed. 



## Installation and Running the Project
### Create a database
You can create the database for the project in any DBMS (MySQL, Microsoft SQL Server) based on your choice. Following is the SQL code that you can copy to create the database.
````
CREATE TALBE BreakdownByContactType
(StartTime Date NOT NULL, 
EndTime Date NOT NULL,
ContactType nvarchar(20) NOT NULL, 
NumberOfTickets int,
PRIMARY KEY (StartTime, EndTime, ContactType);
````

````
CREATE TABLE BreakdownByRequestType
(StartTime Date NOT NULL,
EndTime Date NOT NULL,
RequestType nvarchar(60) NOT NULL,
NumberOfTickets int,
PRIMARY KEY (StartTime, EndTime, RequestType);
````

````
CREATE TABLE BreakdownByCanvasType
(StartTime Date NOT NULL,
EndTime Date NOT NULL,
CanvasType nvarchar(100) NOT NULL,
NumberOfTickets int,
PRIMARY KEY (StartTime, EndTime, CanvasType);
SELECT * FROM BreakdownByCanvasType;
````

````
CREATE TABLE HandleBy
(WorkerEnum int NOT NULL PRIMARY KEY,
CSTLWorker nvarchar(30));

INSERT INTO HandleBy VALUES
(0, 'Everyone'),
(1, 'Student Worker'),
(2, 'GA'),
(4, 'Kris Baranovic'),
(3, 'Mary Harriet');
````
````
CREATE TABLE SupportTicketAssignedTech
(StartTime Date NOT NULL,
EndTime Date NOT NULL,
Tech nvarchar(50) NOT NULL, 
NumberOfTickets int,
PRIMARY KEY (StartTime, EndTime, Tech));
SELECT * FROM SupportTicketAssignedTech;
````
````
CREATE TABLE CSTLConsult
(StartTime Date NOT NULL,
EndTime Date NOT NULL,
FacultyName nvarchar(50) NOT NULL,
ConsultByWorkerEnum int NOT NULL,
ConsultTimes int,
PRIMARY KEY (StartTime, EndTime, FacultyName, ConsultByWorkerEnum),
FOREIGN KEY (ConsultByWorkerEnum) REFERENCES HandleBy(WorkerEnum));
SELECT * FROM CSTLConsult;
````

````
CREATE TABLE ReportTime (StartTime Date NOT NULL, EndTime Date NOT NULL, PRIMARY KEY(StartTime, EndTime));
````

### Set up the environment
- After you have cloned the repos to a directory in your local computer, you deletes the existing environment "whd-env" inside the repos and recreates a new environment by typing the following command in the terminal of the VS Code.
````
python -m venv env
````
- Then change to the env/scripts directory.
````
cd env
cd scripts
````
- Activate the env by typing the following command in VS Code's terminal:
````
./activate
````
- Now you are in env/scrips directory. To install Selenium packages for the project, you type the following command:
````
pip install selenium
````
- You also need to install pyodbc packages to allow the app to interact with the DBMS.
````
pip install pyodbc
````
- And installing the pandas packages to allow the app to use the dataframe to manipulate the scraped data.
````
pip install pandas 
````
- Last but not least, you change the interpreter path by using the environment that you have just created (env).
![image](https://user-images.githubusercontent.com/70489535/146679370-32f63b37-3ac0-4227-99c8-bc10d6f9a559.png)

- Now, you can run the code and see the running result through the main.py file

