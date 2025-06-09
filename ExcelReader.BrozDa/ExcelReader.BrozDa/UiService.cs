using ExcelReader.BrozDa.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spectre.Console;

namespace ExcelReader.BrozDa
{
    internal class UiService
    {

        public void PrintDatabase(List<Person> persons)
        {
            var table = new Table();
            table.AddColumns(GetTableHeaders());

            foreach (var person in persons) 
            {
                table.AddRow(GetTableRow(person));
            }

            AnsiConsole.Write(table);
        }
        private string[] GetTableHeaders() => typeof(Person).GetProperties().Select(x => x.Name).ToArray();
        private string[] GetTableRow(Person person) 
        {
            List<string> dataValues = new();

            var propertyInfo = typeof(Person).GetProperties();
            foreach (var property in propertyInfo) 
            {
                var test = property.GetValue(person);
                dataValues.Add(test.ToString());
            }

            return dataValues.ToArray();
        }
    }
}
