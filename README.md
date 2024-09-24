This code was generated through a mix of Claude Sonnet 3.5 and GPT 4-o. You may see problems in the code. For anybody who wants to use these tools, I found Claude Sonnet much better at generating the code that I wanted it to. It also was much better at refining the code. However the limits for Claude are more severe compared to ChatGPT.

Requirements:

- Python: https://www.python.org/downloads/

- Excel or Libreoffice to see the final output in a neat format.

- Openpyxl for excel integration. Install it with "pip install openpyxl". [Necessary for the code, but you can delete the excel components of the code and just let it output to the log file.]


Purported Features! (at least according to visual checks on the results. I am not familiar with python/actual software development.)
- Logging!
- Excel formatting!
- Limiting of x standbys per week! ( where x is the minimum number of days necessary per week for a given time frame and number of provosts)
- Consideration of special days! (Considers holidays and Sat/Sun as Special Days. Holidays have to be edited manually unfortunately.)
- Random rotations of the list!
- Making sure each person does an equal number of days! (roughly, I think as it fills up the last few days of the given time frame it will inevitably give a few people more shifts unless the total number of required shifts is a nice number)

