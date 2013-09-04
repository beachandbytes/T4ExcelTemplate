using System;

namespace ExcelClassGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            foreach (var item in ExcelHelper.Parse())
            {
                Console.WriteLine(item.ToString());
            }
            Console.ReadLine();
        }
    }
}
