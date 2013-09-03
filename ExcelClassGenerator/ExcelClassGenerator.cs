﻿ 
 
// <autogenerated>
// This code was generated by a tool. Any changes made manually will be lost
// the next time this code is regenerated.
// </autogenerated>
using System;   
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using OfficeOpenXml;
public class DrugSearch
{

    public string CASE_NO_ {get;set;}
    public DateTime? DATE {get;set;}
    public string TIMETM9 {get;set;}
    public string COUNTY {get;set;}
    public string LOCATION {get;set;}
    public string REASON_FOR {get;set;}
    public string RACE {get;set;}
    public string SEX {get;set;}
    public string VEHICLE_YE {get;set;}
    public string VEHICLE_MA {get;set;}
    public string VEHICLE_MO {get;set;}
    public string GROUNDS_FO {get;set;}
    public string CONSENT_RE {get;set;}
    public string CONSENT_GI {get;set;}
    public string DRUGS_REC_ {get;set;}
    public string K9_USED {get;set;}
    public string K9_DRUGS_F {get;set;}
    public string LOCATION_O {get;set;}
    public string TYPE_OF_DR {get;set;}
    public double? AMOUNT_OF_ {get;set;}
    public string MONEY_RECO {get;set;}
    public double? MONEY_AMOU {get;set;}
    public string MONEY_LOCA {get;set;}
    public string TROOPER_NA {get;set;}
    public string UNIT {get;set;}
    public string COMMENTS {get;set;}
    public double? AMOUNTDRUG {get;set;}
    public double? _AMOUNTDRU {get;set;}

	public DrugSearch(IList<string> items) 
	{
			CASE_NO_ = items[0];
            DATE = ToNullable<DateTime>(items[1]);
            TIMETM9 = items[2];
            COUNTY = items[3];
            LOCATION = items[4];
            REASON_FOR = items[5];
            RACE = items[6];
            SEX = items[7];
            VEHICLE_YE = items[8];
            VEHICLE_MA = items[9];
            VEHICLE_MO = items[10];
            GROUNDS_FO = items[11];
            CONSENT_RE = items[12];
            CONSENT_GI = items[13];
            DRUGS_REC_ = items[14];
            K9_USED = items[15];
            K9_DRUGS_F = items[16];
            LOCATION_O = items[17];
            TYPE_OF_DR = items[18];
            AMOUNT_OF_ = ToNullable<Double>(items[19]);
            MONEY_RECO = items[20];
            MONEY_AMOU = ToNullable<Double>(items[21]);
            MONEY_LOCA = items[22];
            TROOPER_NA = items[23];
            UNIT = items[24];
            COMMENTS = items[25];
            AMOUNTDRUG = ToNullable<Double>(items[26]);
            _AMOUNTDRU = ToNullable<Double>(items[27]);
		
	}


	public override string ToString()
	{

    return string.Join("|", typeof(DrugSearch).GetProperties().Select(x => x.GetValue(this, null)));
	}


	public static T? ToNullable<T>(string s) where T: struct
	{
		s = s.Replace("NA","");
        var result = new T?();
        if (!string.IsNullOrEmpty(s) && s.Trim().Length > 0)
        {
            var conv = TypeDescriptor.GetConverter(typeof(T));
            result = (T?)conv.ConvertFrom(s);
        }
        return result;
	}


	}



public class ExcelHelper
{

   public List<DrugSearch>  ParseFile()
   {		
		var result = new List<DrugSearch>();	
		var excelFileInfo  = new FileInfo(@"C:\ExcelT4Template\ExcelClassGenerator\SampleExcelFile.xlsx");
		using (var package = new ExcelPackage(excelFileInfo))
		{
			ExcelWorkbook workBook = package.Workbook;
			if (workBook != null)
			{
				if (workBook.Worksheets.Count > 0)
				{
					ExcelWorksheet currentWorksheet = workBook.Worksheets.First();
					
					for(int row = 2; ; row++)
					{
						var propertyValues = new List<string>();
						for(int col = 1; col <= typeof(DrugSearch).GetProperties().Length; col++)
						{
							var cell = currentWorksheet.Cells[row,col];
							var colValue = cell.Value;
							propertyValues.Add((colValue ?? "").ToString());
						}	
						if(string.Join("",propertyValues).Trim() == "")
							break;
						result.Add(new DrugSearch(propertyValues));

					}
				}
			}
		}
		return result;
   }
	
}
