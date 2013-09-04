T4ExcelTemplate
===============

Converts an excel table into a strongly typed class and sets up the file to be parsed. This tt file relies on EPPlus
package to read the excel file initially. EPPlus only works for xlsx files so the tt file will attempt to convert it
if necessary.

At the bottom of the tt file you will find the following variables. You must set the classname and the excel file name
the rest are optional. The tt file will search for the conversion utility in common places. If it cannot find the 
conversion utility and a conversion is needed, you will need to supply the path to excelcnv.exe.

    ClassName = "DrugSearch";
    ExcelFileName = "SampleExcelFile.xlsx";
    string Password = "";
    string ExcelCnvPath = "";


After the tt file has generated the the underlying code you can access your data by simply calling ExcelHelper.Parse();
    
    var stronglyTypedList = ExcelHelper.Parse();
    
    foreach(var item in stronglyTypedList) 
    {
      Console.WriteLine(item.ToString());
    }

   
