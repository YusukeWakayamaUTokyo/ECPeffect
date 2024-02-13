These .py files are for the measurement of ElectroChemical Peltier (ECP) effect.

## Preparation

In our experiments, the program has been run using ***jupyter*** which is a standard application executed on the browser you can use it when you download **anaconda**.

You can find two program files. Both require some special modules to download into your python environment, but the required modules are different.

Both files use pymeasure for comunicating and controlling sourcemeters, the process needs **EXTRA APPLICATIONS** to be downloaded other than anaconda. You can refer to this page https://qiita.com/matoarea/items/42292f27a9e669a8758a (though it is written in Japanese)

You have to donload in the **PROVIDED ODER HERE** or it does not work somehow.
- First one

https://www.ni.com/ja-jp/support/downloads/drivers/download.ni-488-2.html#345631

(You can change a language from Japanese to English from ”言語” on the right column on the web page)

- Second one

https://www.ni.com/ja/support/downloads/drivers/download.ni-visa.html#346210


## Using codes

You can copy & paste your preferd version of the program file (ECP_almost_pristine.py and ECP_csv.py) to your jupyter note book

### modules (python) needed both program files 
- concurrent.futures (enabling to run multiple functions parallely) #possibly not default module
- datetime (to get date)
- math (for calculation (you can do with numpy, if you want))
- numpy (for calculation)
- os (for saving files)
- pymeasure (for controlling sourcemeter in this case) #not default module
- time (for time management in the experiment)



## ECP_almost_pristine.py (the editted program file used in the paper)

### modules (python) needed in ECP_almost_pristine
- openpyxl (for save the results) #not default module
- request (for notification via Slack (the function for this is disabled as the default)) #possibly not default module

In order to change experimental conditions you have to vary the file itelf each time.
Advantage of this style is you don't have to input conditions, once you put necessary information and you don't change your experimental condition.
Result files are saved in an excel file.
Currently, the two functions in this program is in the state of comment out. One is for creating chart in the excel file automatically, the other is for sending a slack message when the experiment is done.
You can use them if you want.


## ECP_csv.py (written to make it easier to use hopefully)

### modules (python) needed in ECP_csv 
- pandas (for save the results)
- sys (to get detailed information when errors happen during running this program)

Requiring less non-default modules in python. 
Everytime you run this file, you input experimental conditoins not rewriting directly in this file but filling up entry field popped up.
Result files are text file separated by tab.
Until you put a "q" to the entry field, this program won't end.
