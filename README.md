These program files are for the measurement of ElectroChemical Peltier (ECP) effect.

## Preparation

In our experiments, the program has been run using ***Jupyter Notebook*** which is a standard application executed on the browser you can use it when you download **anaconda**.

You can find two program files. Both require some special modules to download into your python environment, but the required modules are different.

Both files use pymeasure for comunicating and controlling sourcemeters, the process needs **EXTRA APPLICATIONS** to be downloaded other than anaconda. You can refer to this page https://qiita.com/matoarea/items/42292f27a9e669a8758a (though it is written in Japanese). You have to download in the **FOLLOWING ORDER** or it does not work somehow.
- First one

https://www.ni.com/en/support/downloads/drivers/download.ni-488-2.html#345631

- Second one

https://www.ni.com/en/support/downloads/drivers/download.ni-visa.html#346210


## Using codes

You can copy & paste your preferd version of the program files (ECP.py and ECP_txt.py) to your jupyter note book.

Before using these codes, you should make sure some experimental conditions. First of all, you have to rewrite the code at the variable 'path_peltier' with your path to access the folder saving experimental results. If it is hard to find, you can use Crtl + f (Command + f; for Mac users) to identify the location of 'path_peltier'.

You, of courese, can change GPIB of the sourcemeters. And if you use different types of thermistor, please change the parameters (B_constant, R25 and T25).

### modules (python) needed both program files 
- concurrent.futures (enabling to run multiple functions parallely)
- datetime (to get date)
- math (for calculation (you can do with numpy, if you want))
- numpy (for calculation)
- os (for saving files)
- pymeasure (for controlling sourcemeter in this case) #not default module
- time (for time management in the experiment)



## ECP.py (the program file used in the paper)

In order to change experimental conditions you have to vary the file itself each time.

Advantage of this style is you don't have to input conditions, once you put necessary information and you don't change your experimental condition.

Result files are saved in an excel file.

Currently, the two functions in this program are in the state of comment out. One is for creating chart in the excel file automatically, the other is for sending a slack message when the experiment is done.
You can use them if you want.

### modules (python) needed in ECP.py
- openpyxl (for saving the results) #not default module
- requests (for notification via Slack (the function for this is disabled as the default)) #possibly not default module


## ECP_txt.py (more interactive version)

Once you input the path in the program file, you don't have to change the file anymore.

Requiring less non-default modules in python. 

Everytime you run this file, you have to input experimental conditoins not rewriting directly in this file but filling up entry fields popped up.

Result files are text file separated by tab.

Until you put a "q" to the entry field, this program won't end.

The style of inputting 1 period of measurement is different from ECP_almost_pristine.py. The former one requires half of a period, but this file needs hole period. (Just a matter of length of string in the input() function, you can change this if you want)

The default measurement interval was elongated from 0.4 to 0.5 [s] to be devisible any integer second.

### modules (python) needed in ECP_txt.py
- pandas (for saving the results)
- sys (to get detailed information when errors happen during running this program)

## Working environment
I checked the performance of the programs under following conditions.
- Windows PC
- python 3.9.7
- jupyter (ipykernel) 6.4.1
- openpyxl 3.0.9
- requests 2.26.0
- numpy 2.26.0
- pandas 1.3.4
- pymeasure 0.10.0

## Authors
Fumitoshi Matoba

Yusuke Wakayama
