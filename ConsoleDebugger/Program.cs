using System;
using System.Collections.Generic;

using DataLaundry;
using System.Reflection;

namespace ConsoleDebugger
{
    class MainClass
    {
        public static void Main(string[] args)
        {
         

            List<Company> testList = new ExcelHandler().ExcelReader();

            new JsonHandler().JsonFixer(testList);

            foreach (var company in testList)
            {
               
                
                Console.WriteLine("NYTT FÖRETAG");
                Console.WriteLine();

                PropertyInfo[] props = company.GetType().GetProperties();

                foreach (var p in props)
                {
                    Console.WriteLine(p.GetValue(company));
               

                }
                Console.WriteLine();
            }

            Console.ReadKey();
        }
    }
}
