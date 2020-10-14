# Python_ExecuteSQL_WriteCSV
The code is used to automate and execute  sql query from sql server and store them in  csv and xlsx files.

# Install Dependencies
1. Install dependencies Using 
```
pip install csv
pip install pandas
pip install openpyxl
```
2. ODBC driver is needed 





# What the code does?
1. The code executes 4 sql queries  and stores the result which are part of differert tables of same DB.
2. It prints these on the console.
3. These consists of one particular same Column 'ColumnName2' on basis of which FINAL  sheet is to be produced.
4. Pandas is used to combine the generated Intermediate CSV.
5. There is provison also to drop or keep the Intermediate CSV.
6. The code also generates a  XLSX file.
7. The code also has beautify options for the EXCEL such as header formatiing and multiple options are available under the openpyxl.
8. It also gives user to dynamically add values to sql query-Here shown for two time taken as input from user.
