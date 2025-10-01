namespace Tempo.Exporter
{
    internal class ConsoleFileSelector
    {
        private string _folderPath;

        public ConsoleFileSelector(string folderPath)
        {
            _folderPath = folderPath;    
        }

        public string SelectExcelFile()
        {
            if (!Directory.Exists(_folderPath))
            {
                Console.WriteLine("Directory does not exist.");
                return null;
            }

            var files = Directory.EnumerateFiles(_folderPath, "*.xlsx")
                                 .Select(Path.GetFileName)
                                 .ToList();

            if (files.Count == 0)
            {
                Console.WriteLine("No Excel files found.");
                return null;
            }

            if (files.Count == 1)
            {
                return Path.Join(_folderPath, files.First());
            }

            int selectedIndex = 0;

            ConsoleKeyInfo keyInfo;

            do
            {
                Console.Clear();
                Console.WriteLine("Use TAB to cycle through Excel files. Press ESC to exit or ENTER to confirm.\n");

                for (int i = 0; i < files.Count; i++)
                {
                    if (i == selectedIndex)
                    {
                        // Highlight current entry (e.g. invert colors)
                        Console.BackgroundColor = ConsoleColor.White;
                        Console.ForegroundColor = ConsoleColor.Black;
                        Console.WriteLine(files[i]);
                        Console.ResetColor();
                    }
                    else
                    {
                        Console.WriteLine(files[i]);
                    }
                }

                keyInfo = Console.ReadKey(intercept: true);

                if (keyInfo.Key == ConsoleKey.Tab)
                {
                    selectedIndex = (selectedIndex + 1) % files.Count;
                } else if (keyInfo.Key == ConsoleKey.Enter)
                {
                    return Path.Join(_folderPath, files[selectedIndex]);
                }

            } while (keyInfo.Key != ConsoleKey.Escape);

            return null;
        }
    }
}
