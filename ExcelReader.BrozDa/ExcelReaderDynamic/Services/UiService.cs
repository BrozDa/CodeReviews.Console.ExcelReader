using Spectre.Console;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderDynamic.Services
{
    internal class UiService
    {
        public int GetMenuInput(Dictionary<int, string> menuOptionMap)
        {
            var input = AnsiConsole.Prompt(
                new SelectionPrompt<int>()
                .AddChoices(menuOptionMap.Keys.ToArray())
                .UseConverter(x => menuOptionMap[x]));

            return input;
        }
        public void PrintRecords(List<Record> records)
        {
            var table = new Table();

            table.AddColumns(records[0].Headers.ToArray());
            foreach (var record in records) 
            { 
                table.AddRow(record.Data.ToArray());
            }

            AnsiConsole.Write(table);
        }
        public string GetFilePathFromUser()
        {
            var input = AnsiConsole.Prompt(
                new TextPrompt<string>("Please enter full path to file file: ")
                );

            return input;
        }

    }
}
