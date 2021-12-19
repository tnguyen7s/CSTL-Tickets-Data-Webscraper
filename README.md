# CSTL-Tickets-Data-Webscraper
## Project Description
- In this project, I utilized Selenium packages to scrape data from the Web Help Desk website to collect monthly ticket data replacing the manual method and automate the monthly task at my working place (CTL). 
- The application provides a UI using Tkinter packages. The UI allows users to give inputs to the scraper and to view and insert the collected data to the SQL Server database.
- The limitation of the app that needs to be improved in the future: it is quite slow from the scrapping work and sometimes the application crashes due to the fact that the Scraper could not find the UI element to select. It also requires constant updates because sometimes the CSS Selector gets changed. 



## Installation and Running the Project
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

