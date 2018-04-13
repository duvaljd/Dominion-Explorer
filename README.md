# 1.	REQUIREMENTS

## 1.1	Required Files
All of these files should be present when you download this program:
- SI507F17_finalproject.py
- SI507F17_finalproject_tests.py
- creds.py
- requirements.txt (recommended)
- dominion_cache.json (recommended)

## 1.2	Required Python Version & Modules
This program was built in Pythong 3.6, and it is recommended that you have Python 3.6 installed to run it. All of the necessary modules can easily be installed via your command prompt; see step 3 of 2.1 below. The necessary modules are:

Module | Version | Function
------ | ------- | --------
BeautifulSoup 4 | v4.6.0 | (for parsing the data from the dominion wiki)
DateTime | | (for setting an expiration date on cache entries)
JSON | | (for keeping the cache organized)
html5lib | | (used in conjunction with bs4 to scrape pages)
html2text | v2017.10.4 | (for parsing html collected from bs4)
openpyxl | v2.4.9 | (for creating the excel output files)
psycopg2 | v2.7.3.2 | (for interacting with the database)
random | | (for generating random cards if users don't know the names of cards)
requests | v2.18.4 | (for scraping data from the dominion wiki)
sys | | (for debugging)
os | | (for debugging)
unittest | | (for testing)

## 1.3 Required Database
This program requires a PSQL database. See 2.1 below for setup instructions.


# 2.	HOW TO USE THIS PROGRAM

## 2.1	Setting Up (Please read entirely before beginning steps)

### 1.	Database
- Create a PSQL database with the name:

> duvaljd_507final

(Alternatively, you may name the database anything you want - you'll just have to update the DB_NAME variable in creds.py.)

- Write down the user name and password for your database

### 2.	creds.py
- Find and open the creds.py file in the program folder

- Update the DB_NAME variable with the name you chose for your database, or leave it as is if you called your database duvaljd_507final as instructed 

- Update the DB_USER variable with the username for your database 

- Update the DB_PASS variable with the password for the user in your database.

- Save and close creds.py

### 3.	Install required modules
- If you are setting up a virtual environment, do so now.

- To install the required modules, type:

> pip install -r requirements.txt

### 4.	Run the file for the first time
- In a command prompt, navigate to the project folder.

- From the command prompt, type:

> py SI507F17_finalproject.py

- The program will take several minutes to run the first time, because it collects a large amount of data from hundreds of individual web pages. Please be patient! If you wish to see please see section 2.3.

- I have included the cache file, dominion_cache.json, in case you do NOT want to sit through the several-minutes long process of building a new cache. IF YOU DO NOT DELETE THE INCLUDED CACHE, THE PROGRAM WILL LOAD ALL FILES FROM IT. Delete the cache to get fresh data.

- When the program has finished making tables, building a cache, and inserting data into the database, you will be presented with a prompt. See section 2.2 for how to interact with the prompt.

## 2.2	Interacting with the Prompt
### 0.	Some notes
- After running the program for the first time, it is recommended that you open the SI507F17_finalproject.py file and set the GETNEWDATA variable to False. This will prevent the program from re-inserting data into the database each time you run it.

- If you're at the main prompt and do not know what commands you can enter, type:

> help

to get a list of commands.

- If the prompt doesn't know what to do with your command, it will simply ask you for another input. It's very forgiving!

### 1.	Building a list of cards
- Type:

> list

to build a list of all the cards in the database that will be exported to a .xlsx file, dominion.xlsx, located in the same folder as SI507F17_finalproject.py. You can use this list to find cards to search for in the program.

### 2.	Selecting a specific card
- To see data for a specific card, type the name of the card into the main prompt. THE PROMPT IS CASE SENSITIVE - you must type the card name exactly as it appears on the card, including apostrophes, dashes, and spaces, or it will not find the card you request.

- If you don't know the name of a card, you can type:

> random

to let the program choose a random card for you.
- Once you have selected a card by typing its name or the random command, the details of the card will be displayed, followed by a new prompt that lists the commands you can run on the currently selected card.

### 3.	Generating files
- At the card prompt, you can input any of the following commands: card, recs, set, back.

- Inputting any of the card, recs, or set commands will generate files as detailed below, and then return you to the card prompt. Inputting back will take you to the main prompt.

- Typing:

> card

generates an xlsx file in the same folder as SI507F17_finalproject.py. The file name is the name of the card, prefixed by card_.

- Typing:

> recs

generates an xlsx file in the same folder as SI507F17_finalproject.py. The file name is the name of the card, prefixed by recs_.

- Typing:

> set

generates an xlsx file in the same folder as SI507F17_finalproject.py. The file name is the name of the card, prefixed by sets_.

### 4.	Exiting the program
- To exit the program, you must be at the main prompt. If you are at the card prompt, type:

> back

to return to the main prompt.

- Once you are at the main prompt, type:

> done

to exit the program.

- Alternatively, you can use your command window's keyboard shortcut to stop the program (in GitBash on windows, this is ctrl+c)

## 2.3	Options
### 1.	Turning off data collection & insertion.
- After running the file for the first time, it is recommended that you find the GETNEWDATA variable and change it to False. This will stop the program for emptying the database, creating new tables, fetching data from the cache, and inserting into the database each time you run it. It is located in the top of the SI507F17_finalproject.py file, under the comment header GLOBAL VARIABLES.

### 2.	Turning on debugging
- If you would like to watch the program's progress, or if you need to debug it, you should find the DEBUG variable and set it to True. This will flush messages out as the program completes tasks, and will give detailed error messages if applicable. It is located in the top of the SI507F17_finalproject.py file, under the comment header GLOBAL VARIABLES.

# 3.	HOW THIS PROGRAM WORKS
## 3.1 Overview
### 1.	Database
- After the user builds a database and updates creds.py with the name, user, and password for the database, it creates 4 tables: cards, csets, recs, and cardsInRecs.
- cards has columns for a card id, a set id (foreign key constraint), name, cost, types, illustrators, and description.
- csets has columns for a set id, name, cardnumber, themes, release, and coverart.
- recommendations has columns for recommendation id, name, set1, set2, set3.
- cardsInRecs has columns for card id and recommendation id.

### 2.	Cache
The program uses a cache function lightly adapted from the University of Michigan's School of Information SI507 Programming II graduate course from the Fall of 2017. The program checks the cache to see if the identifier is a key inside it; if the data is expired, or if it does not find the key, it fetches data from the supplied URL, and inserts the data into the cache with a key, a url of the page it got the data from, the data it got, the date and time it got the data, and the number of days in which it will expire.

### 3.	Collecting Data
- The program uses some supplied links to the dominion wiki to build a list of links to the cards and sets pages of the wiki. For each link, it uses the requests module to get the data from the website. The data is turned into a beautifulsoup object via the beautiful soup module, and only the useful data is inserted into the cache.
- Once the data is in the cache (or has been loaded from the cache), the program uses the html2text module to sort through the html and get the relevant data to be inserted into the database. Data insertion is done via the psycopg2 module.

### 4.	Classes
- The classes query the database based on a card id, set id, or recommendation id, and then populate the init variables with date from the tables.
- The Card class includes a method to build a list of recommendations based on its card id.
- The Set class includes a method to build a list of all the cards associated with it and assign them to a variable. It also has a contain method which looks to see if the the string the user inputs is in the list of cards associated with the set.
- The Rec class builds a list of cards associated with the input recommendation id.

### 5.	Functions
The program includes many functions. Here are the important ones:
- The makeList function creates a giant list of all the cards in the database. It makes use of the 'all' Functions, which query the database to find all the card, set, and recommendation ids and assign them to lists in variables.
- The makeFile functions take single card id, query the database for related set or recommendation ids if necessary, build class objects, and then use openpyxl to create excel files of relevant data.

### 5.	The Prompt
The prompt is a series of nested while loops and if-elif-else statements with code inside that allow the user to make use of all the code that comes before it. It makes use of nearly every class and function to turn commands input from the user into printed information or files.

# 4.	MISCELLANEOUS
## 4.1 Resources Used
- BeautifulSoup documentation: https://www.crummy.com/software/BeautifulSoup/bs4/doc/

- html2text documentation: https://github.com/aaronsw/html2text

- openpyxl documentation: http://openpyxl.readthedocs.io/en/default/

- An obscene amount of StackOverflow threads

## 4.2 Notes
- I apologize for the incredible amount of code here... I know it's not elegant, but 50 hours in I'm happy that it's all working!

- You probably already know, but when you run the tests file, just input 'done' at the prompt to run the tests.

- I'm interested in maybe going further with this code, so if you have any tips for consolidating, I would greatly appreciate it - even if that's only after you finish all the grading y'all have to do!

- Thanks!!