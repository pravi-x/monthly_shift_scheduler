# monthly_shift_scheduler
## What it does
It provides a schedule for a given month for a given number of workers. It keeps track of the preferences of the workers (of their prefered days off) as well as the days on shifts on weekends and other holidays.

## How was made
- The program uses the library ortools from google and spesifically: `from ortools.sat.python import cp_model`
- For the UI the `PyQt6` library is used
- The program exports to the Documents folder of windows using `xlsxwriter` library

## How to run it
Download the zip file and run the main.exe on windows
or
By downloading and running the `main.py` file
```
git clone https://github.com/pravi-x/monthly_shift_scheduler
cd monthly_shift_scheduler
python main.py
```
or
For results only on the terminal
```
python main_cmd.py
```

# Screenshots of the program
![Alt text](image.png)

![Alt text](image-1.png)
