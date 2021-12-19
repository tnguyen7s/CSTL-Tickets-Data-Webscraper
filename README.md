# CSTL-Tickets-Data-Webscraper
## Project Description
- In this project, I utilized Selenium packages to scrape data from the Web Help Desk website to collect monthly ticket data replacing the manual method and automate the monthly task at my working place (CTL). 
- The application provides a UI using Tkinter packages. The UI allows users to give inputs to the scraper and to view and insert the collected data to the SQL Server database.
- The limitation of the app that needs to be improved in the future: it is quite slow from the scrapping work and sometimes the application crashes due to the fact that the Scraper could not find the UI element to select. It also requires constant updates because sometimes the CSS Selector gets changed. 

## Installation and Running the Project
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
