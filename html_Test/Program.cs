using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Threading;
using HtmlAgilityPack;
using static System.Runtime.InteropServices.JavaScript.JSType;

class Program
{
    // Define your connection string based on your Access database location
    private static readonly string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\oliver\\Documents\\Tooldb_test.accdb;";
    static void Main()
    {
        // Define a dictionary of predetermined websites with user-friendly names and corresponding parsers
        Dictionary<int, (string Name, Action<string> Parser)> websites = new Dictionary<int, (string, Action<string>)>
        {
            { 1, ("OSG", PerformOSGWebScraping) },
            { 2, ("ISCAR", PerformISCARWebScraping) },
            // Add more websites as needed
        };

        while (true)
        {
            // Prompt the user to select a website
            Console.WriteLine("Select a website to scrape (or type 'exit' to close the program):");
            foreach (var entry in websites)
            {
                Console.WriteLine($"{entry.Key}. {entry.Value.Name}");
            }

            // Get user's choice
            Console.Write("Enter the number of the website (or 'exit' to close): ");
            string userInput = Console.ReadLine();

            if (userInput.ToLower() == "exit")
            {
                break; // Exit the loop and close the program
            }

            if (int.TryParse(userInput, out int selectedWebsite) && websites.ContainsKey(selectedWebsite))
            {
                // Prompt the user to enter a text (page number)
                Console.Write("Enter text to search the webpage: ");
                string searchText = Console.ReadLine();

                // Call the corresponding parser based on user's selection
                websites[selectedWebsite].Parser(searchText);
            }
            else
            {
                Console.WriteLine("Invalid selection. Please enter a valid number or 'exit'.");
                Console.WriteLine("");
            }
        }

        Console.WriteLine("Program closed.");
    }

    static void PerformOSGWebScraping(string searchText)
    {
        // OSG-specific scraping logic
        string url = $"https://nl.osgeurope.com/{searchText}.html";
        Console.WriteLine($"Performing OSG web scraping for URL: {url}");

        string[] Classes = {    "specification dc",
                                "specification apmx",
                                "specification lu",
                                "specification lf",
                                "specification re",
                                "specification dn",
                                "specification dcon",
                                "specification zefp",
                                "specification sig",
                                "specification oal",
                                "specification pl",
                                "specification uldr",
                                "specification flute_type",
                                "specification nof",
                                "specification tcdc",
                                "specification tcdmm",
                                "specification thread_type",
                                "specification td",
                                "specification tp_tpi",
                                "specification din",
                                "specification drvs",
                                "specification hole_type_application",
                                "specification tap_type",
                                "specification thchl",
                                "specification thcht",
                                "specification thlgth",
                                "specification thread_tolerance",
                                "specification type",
                                "specification premachined_hole_diameter__phd_min",
                                "specification premachined_hole_diameter_phd_max_",
                                "product-description",
                                "materials"};

        // Get the initial results
        string[] initialResults = ExtractSecondChildText(url, Classes, searchText);

        // Add the last entry
        string edp = searchText;
        string[] finalResults = new string[initialResults.Length + 3];

        Array.Copy(initialResults, finalResults, initialResults.Length);

        // Add the last entry to the end of the array
        //finalResults[finalResults.Length - 4] = edp;
        finalResults[finalResults.Length - 3] = "";
        finalResults[finalResults.Length - 2] = "";
        finalResults[finalResults.Length - 1] = url;
        finalResults[30] = edp;

        
        // Loop over the array and add "" to empty entries
        for (int i = 0; i < finalResults.Length; i++)
        {
            if (string.IsNullOrEmpty(finalResults[i]))
            {
                finalResults[i] = "0";
            }
        }

        // Alternatively, you can use Array.ConvertAll
        finalResults = Array.ConvertAll(finalResults, element => string.IsNullOrEmpty(element) ? "" : element);

        // Loop over the array and add single quotes to every entry
        string[] quotedArray = new string[finalResults.Length];

        for (int i = 0; i < finalResults.Length; i++)
        {
            quotedArray[i] = "'" + finalResults[i] + "'";
        }

        // Connection string to your Access database
        string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\oliver\\Documents\\Tooldb_test.accdb;";

        // Specify the table name in the database
        string tableName = "OSG_1"; // Replace this with your actual table name

        // Specify the column names you want to use in the database
        List<string> columnNames = new List<string>
        {
            "dc", "apmx", "lu", "lf", "re", "dn", "dcon", "zefp", "sig", "oal", "pl", "uldr", "flute_type", "nof", "tcdc", "tcdmm", "thread_type", "td", "tp_tpi", "din", "drvs", "hole_type_application", "tap_type", "thchl", "thcht", "thlgth", "thread_tolerance", "type", "premachined_hole_diameter__phd_min", "premachined_hole_diameter_phd_max_", "[product-description]", "P", "M", "K", "N", "S", "H", "Amount", "[Reorder amount]", "link"
        };

        // Call the function to insert data into the database
        InsertDataIntoAccessDB(connectionString, tableName, columnNames, quotedArray);

        Console.WriteLine("Data inserted successfully.");

    }

    static void PerformISCARWebScraping(string searchText)
    {
        // ISCAR-specific scraping logic
        string url = $"https://www.iscar.com/eCatalog/item.aspx?cat={searchText}&fnum=4105&mapp=ML&app=59&GFSTYP=M&isoD=1";
        Console.WriteLine($"Performing ISCAR web scraping for URL: {url}");

        // Initialize HtmlWeb and HtmlDocument from HtmlAgilityPack
        HtmlWeb web = new HtmlWeb();
        HtmlDocument document = web.Load(url);

        Thread.Sleep(2000);

        // Specify XPath to select the text within all elements of the specified XPath
        string targetXPath = "/html/body/form/div[3]/section[2]/div/div[2]/div[7]/div[2]/div//text()";

        // Use SelectNodes to get all text within the elements of the specified XPath
        var textWithinTarget = document.DocumentNode.SelectNodes(targetXPath);

        if (textWithinTarget != null)
        {
            // Iterate through each text node, replace "&nbsp;" and check if it is not empty, then display its content
            foreach (var textNode in textWithinTarget)
            {
                string trimmedText = textNode.InnerText.Trim().Replace("&nbsp;", "");
                if (!string.IsNullOrEmpty(trimmedText))
                {
                    Console.WriteLine($"Text Content: {trimmedText}");
                }
            }
        }
        else
        {
            Console.WriteLine("No text found within the specified XPath.");
        }

    }

    static string[] ExtractSecondChildText(string htmlContent, string[] classNames, string edp)
    {
        // Create HtmlDocument and load HTML content
        HtmlWeb web = new HtmlWeb();
        HtmlDocument document = web.Load(htmlContent);

        // Initialize a list to store results
        List<string> resultList = new List<string>();

        for (int i = 0; i < classNames.Length; i++)
        {
            // Use XPath to select the element with the specified class
            HtmlNode node = document.DocumentNode.SelectSingleNode($"//div[@class='{classNames[i]}']");

            switch (classNames[i])
            {
                case ("product-description"):
                    // Extract text from the specified XPath
                    HtmlNode productDescriptionNode = document.DocumentNode.SelectSingleNode($"//div[@class='product-description']");

                    string productDescriptionConverted = productDescriptionNode?.InnerText ?? "";
                    string productDescriptionExtractedText = ExtractTextAfter(productDescriptionConverted, edp);

                    // Check if the target node is found before extracting text
                    resultList.Add(productDescriptionExtractedText);
                    break;

                case ("materials"):
                    // Extract text from the specified XPath
                    HtmlNode materialsNode = document.DocumentNode.SelectSingleNode($"//div[@class='materials']");
                    string materialsConverted = materialsNode?.InnerText ?? "";

                    // Categories array
                    string[] categories = { "Staal", "RVS", "Gietijzer", "Aluminiumlegeringen", "Titaniumlegeringen", "Gehard staal" };

                    // Create an array to store the results
                    int[] result = new int[categories.Length];

                    // Iterate through categories and check if they exist in the text
                    for (int j = 0; j < categories.Length; j++)
                    {
                        if (materialsConverted.Contains(categories[j], StringComparison.OrdinalIgnoreCase))
                        {
                            result[j] = 1;
                        }
                        else
                        {
                            result[j] = 0;
                        }

                        // Add each result to the list
                        resultList.Add(result[j].ToString());
                    }

                    break;

                default:
                    if (node != null)
                    {
                        // Use XPath to select the second child node
                        HtmlNode secondChild = node.ChildNodes.ElementAtOrDefault(1);

                        // Check if the second child node is found before extracting text
                        resultList.Add((secondChild != null) ? secondChild.InnerText : "");
                    }
                    else
                    {
                        // Element with the specified class not found
                        resultList.Add("");
                    }
                    break;
            }
        }

        // Convert the list to an array before returning
        return resultList.ToArray();
    }

    static string ExtractTextAfter(string originalText, string variableValue)
    {
        // Find the position of the variable's value
        int variableIndex = originalText.IndexOf(variableValue);

        // Check if the variable's value is found
        if (variableIndex != -1)
        {
            // Use Substring to get the text after the variable's value
            return originalText.Substring(variableIndex + variableValue.Length).Trim();
        }
        else
        {
            // Variable's value not found, return an empty string or handle accordingly
            return "";
        }
    }

    static void InsertDataIntoAccessDB(string connectionString, string tableName, List<string> columnNames, object[] dataArray)
    {
        if (columnNames.Count != dataArray.Length)
        {
            throw new ArgumentException("Number of column names must match the length of the data array.");
        }

        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            connection.Open();

            // Build the SQL command for a single row
            string query = $"INSERT INTO {tableName} ({string.Join(", ", columnNames)}) VALUES ({string.Join(", ", dataArray.Select(value => $"{value}"))})";

            using (OleDbCommand command = new OleDbCommand(query, connection))
            {
                // Add parameters for all values
                for (int i = 0; i < dataArray.Length; i++)
                {
                    command.Parameters.AddWithValue($"@{dataArray[i]}", string.IsNullOrEmpty((string)dataArray[i]) ? DBNull.Value : dataArray[i]);
                }

                // Execute the command
                command.ExecuteNonQuery();
            }
        }
    }

}
