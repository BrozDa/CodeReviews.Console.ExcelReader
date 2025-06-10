using ExcelReader.Brozda.Helpers;
using Spectre.Console;


namespace ExcelReaderDynamic.Services
{
    /// <summary>
    /// Handles any input/ output operation for the application
    /// </summary>
    internal class UiService
    {
        /// <summary>
        /// Prints "Press any key to continue" text and awaits user input
        /// </summary>
        public void PressAnyKeyToContinue()
        {
            Console.WriteLine();
            Console.WriteLine(AppStrings.Io_PressAnyKeyToContinue);
            Console.ReadKey(); 
        }
        /// <summary>
        /// Clears existing output from the console
        /// </summary>
        public void ClearConsole()
        {
            Console.Clear();
        }
        /// <summary>
        /// Prints out the error message to the output and awaits user input
        /// </summary>
        /// <param name="errorMsg">A <see cref="string"/> containing the error message</param>
        public void PrintErrorMsg(string? errorMsg)
        {
            Console.WriteLine(errorMsg ?? AppStrings.Io_UnspecifiedError);
            PressAnyKeyToContinue();
        }
        /// <summary>
        /// Prints text to the output
        /// </summary>
        /// <param name="errorMsg">A <see cref="string"/> containing text to be printed</param>
        public void PrintText(string text)
        {
            Console.WriteLine(text);
        }
        /// <summary>
        /// Prints out the menu and gets user choice
        /// </summary>
        /// <param name="menuOptionMap">A <see cref="Dictionary{TKey, TValue}"/> of int/ string key value pairs. A int value representing enum value
        /// of menu choice. String representing text to be printed to the output"/></param>
        /// <returns>A int value representing enum value of menu choice.</returns>
        public int GetMenuInput(Dictionary<int, string> menuOptionMap)
        {
            var input = AnsiConsole.Prompt(
                new SelectionPrompt<int>()
                .AddChoices(menuOptionMap.Keys.ToArray())
                .UseConverter(x => menuOptionMap[x]));

            return input;
        }
        /// <summary>
        /// Prints out record to the output
        /// </summary>
        /// <param name="records">A list of <see cref="Record"/> to be printed</param>
        public void PrintRecords(List<Record> records)
        {
            var table = new Table();

            table.AddColumns(records[0].Headers.ToArray());
            foreach (var record in records) 
            { 
                table.AddRow(record.Data.ToArray());
            }

            AnsiConsole.Write(table);
            PressAnyKeyToContinue();
        }
        /// <summary>
        /// Gets a string representing the file path from the user
        /// </summary>
        /// <returns>A string representing the file path</returns>
        public string GetFilePathFromUser()
        {
            var input = AnsiConsole.Prompt(
                new TextPrompt<string>("Please enter full path to file file: ")
                );

            return input;
        }
        /// <summary>
        /// Gets confirmation from user to agree/ deny overwriting of the file
        /// </summary>
        /// <returns>A <see cref="bool"/> true if user agree with overwritting, false otherwise</returns>
        public bool ConfirmOverwrite()
        {
            var input = AnsiConsole.Prompt(
                new SelectionPrompt<bool>()
                .Title("\nFile already exists. Please confirm the operation")
                .AddChoices(true, false)
                .UseConverter(x => x == true ? "Confirm" : "Cancel")
                );

            return input;

        }

    }
}
