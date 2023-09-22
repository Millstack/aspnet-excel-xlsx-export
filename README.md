# aspnet-excel-xlsx
Asp.net C# code for downloading the grid to Excel file with .xlsx format which is the current one. Here we can also customize the excel rows like color grading or the fix excel column width using the code behind

## Excel Format
Here the usual excel export code available in asp snippets is to download the Excel file in .Xls format which is only good to read
This code is to export the grid into an excel file with .xlsx format

## Code
first install the `EPPlus` library from NuGet Package Manager Console
`Install-Package EPPlus`

Then for using this library we need to provide the Non-Commercial licence  
`ExcelPackage.LicenseContext = LicenseContext.NonCommercial;`

Then we can start using the asp snippet

## Code Logic
- first we are setting the excel sheet's column name ie Header Row
- we are using cell formating like setting header color, column width, cell format, so on....
- using foreach loop we are traversing the datatable dt and setting the cell the value of selected column name
- then we are setting the excel file name and exporting it to download

### NOTE
There is an another way to fill the excel sheet directly using the GridView which is already binded with the datasource
for the click here to get the github repository
<a href="https://google.com" style="background-color: blue; color: white; padding: 5px 10px; border-radius: 20px; text-decoration: none;">
  View Source<span style="background-color: yellow; padding: 5px; border-radius: 0 20px 20px 0;"></span>
</a>
