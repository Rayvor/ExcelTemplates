using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTemplates
{
    class Program
    {
        static void Main(string[] args)
        {
            var date = GetValue(DateTime.Now.ToShortDateString());
            var name = GetValue("John");
            var test = GetValue("Test");
            var names = GetValues("Name", 10).ToArray();
            var tableItems = GetValues("TItem", 5).ToArray();

            string pathIn = @"C:\Users\neoql\Downloads\1.xlsx";
            string pathOut = @"C:\Users\neoql\Downloads\2.xlsx";

            var xlTemplate = new ExcelTemplate(pathIn);

            xlTemplate.SetField("Date", date);
            xlTemplate.SetField("Name", name);
            xlTemplate.SetField("Test", test);
            xlTemplate.SetField("Names", names);
            xlTemplate.SetField("TableItem", tableItems);

            byte[] output = xlTemplate.GetAsByteArray();

            File.WriteAllBytes(pathOut, output);
        }

        private static IEnumerable<string> GetValues(string name, int count)
        {
            for (int i = 0; i < count; i++)
            {
                yield return $"{name} {i}";
            }
        }

        private static string GetValue(string name)
        {
            return name;
        }
    }
}
