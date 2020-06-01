# DataExportFromMSAccess
Export all data from MS Access to any data source, here I have used Excel as target
This application helps you migrate any legacy Microsoft Access database to Excel or any other database

Requirements:
1. Source Database (Sample included)
2. Target Path (For genrated excel files to be copied to)
3. OleDb Provider (Connect to Microsoft Access)

The code is in C# and has 2 step process
//1. Get all data from the Source database 
DataSet ds = GetAllDataFromSource();

//2. Export all data to Excel
ExportDataSetToExcel(ds);


Results:
![Console Application Output](https://www.codeproject.com/KB/Articles/5269825/Working/Results.PNG)
![Target Database As Excel](https://www.codeproject.com/KB/Articles/5269825/Working/Results-Excel.PNG)

Feel free to extend this tool to include other database.
