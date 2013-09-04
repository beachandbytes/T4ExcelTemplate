using System;

namespace ExcelClassGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var excelHelper = new ExcelHelper();
            var itemList = excelHelper.ParseFile();
            foreach (var item in itemList)
            {
                Console.WriteLine(item.ToString());
            }
            Console.ReadLine();
        }
    }
}
