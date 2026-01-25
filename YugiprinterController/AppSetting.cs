using System;
using System.IO;

namespace YugiprinterController
{
    /// <summary>
    /// Represents application settings read from a simple text file with "Key: Value" lines.
    /// Expected keys (case-insensitive): SelectedFolderPath, PrintCloseToCard
    /// </summary>
    public class AppSetting
    {
        public string SelectedFolderPath { get; }
        public bool PrintCloseToCard { get; }

        private AppSetting(string selectedFolderPath, bool printCloseToCard)
        {
            SelectedFolderPath = selectedFolderPath;
            PrintCloseToCard = printCloseToCard;
        }

        public static AppSetting Load(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentNullException(nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("Settings file not found.", filePath);

            // defaults
            string selectedFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            bool printClose = false;

            foreach (var raw in File.ReadAllLines(filePath))
            {
                if (string.IsNullOrWhiteSpace(raw))
                    continue;

                var line = raw.Trim();

                // allow comments
                if (line.StartsWith("#") || line.StartsWith("//"))
                    continue;

                var colonIndex = line.IndexOf(':');
                if (colonIndex < 0)
                    continue;

                var key = line.Substring(0, colonIndex).Trim();
                var value = line.Substring(colonIndex + 1).Trim().Trim('"');

                if (key.Equals("SelectedFolderPath", StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.IsNullOrWhiteSpace(value))
                        selectedFolder = value;
                }
                else if (key.Equals("PrintCloseToCard", StringComparison.OrdinalIgnoreCase))
                {
                    // accept true/false, 1/0, yes/no
                    bool parsedBool;
                    if (bool.TryParse(value, out parsedBool))
                    {
                        printClose = parsedBool;
                    }
                    else
                    {
                        switch (value.Trim().ToLowerInvariant())
                        {
                            case "1":
                            case "yes":
                            case "y":
                                printClose = true;
                                break;
                            default:
                                printClose = false;
                                break;
                        }
                    }
                }
            }

            return new AppSetting(selectedFolder, printClose);
        }
    }
}