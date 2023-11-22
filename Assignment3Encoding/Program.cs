using System.Data;
using System.Xml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using Microsoft.VisualBasic.FileIO;

using (HttpClient client = new HttpClient())
{
    string apiUrl = "https://jsonplaceholder.typicode.com/users";

    HttpResponseMessage response = await client.GetAsync(apiUrl);

    if (response.IsSuccessStatusCode)
    {
        string responseData = await response.Content.ReadAsStringAsync();

        //Console.WriteLine(responseData);
        string jsonData = responseData;

        List<User> users = JsonConvert.DeserializeObject<List<User>>(jsonData);

        displayData(users);

        storeInExcel(users);

        readXMLFiles();

        readCSVFiles();

        readJSONFiles();
    }
}

void displayData(List<User>? users)
{
    Console.WriteLine("ID\t| Name\t\t\t| Username\t\t\t| Email\t\t\t\t\t\t\t| Address");
    Console.WriteLine(new string('-', 130));

    foreach (User user in users)
    {
        string address = $"{user.address.street}, {user.address.suite}, {user.address.city}, {user.address.zipcode}";
        Console.WriteLine($"{user.id}\t| {user.name.PadRight(20)}\t| {user.username.PadRight(20)}\t| {user.email.PadRight(20)}\t| {address}");
        Console.WriteLine(new string('-', 130));
    }
}

void storeInExcel(List<User>? users)
{
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    using (var package = new ExcelPackage())
    {
        var worksheet = package.Workbook.Worksheets.Add("Users");

        worksheet.Cells[1, 1].Value = "ID";
        worksheet.Cells[1, 2].Value = "Name";
        worksheet.Cells[1, 3].Value = "Username";
        worksheet.Cells[1, 4].Value = "Email";
        worksheet.Cells[1, 5].Value = "Address";

        int row = 2;
        foreach (var user in users)
        {
            worksheet.Cells[row, 1].Value = user.id;
            worksheet.Cells[row, 2].Value = user.name;
            worksheet.Cells[row, 3].Value = user.username;
            worksheet.Cells[row, 4].Value = user.email;

            string address = $"{user.address.street}, {user.address.suite}, {user.address.city}, {user.address.zipcode}";
            worksheet.Cells[row, 5].Value = address;

            row++;
        }

        package.SaveAs(new System.IO.FileInfo("UsersData.xlsx"));
    }

    Console.WriteLine("Excel file created successfully.");

}

void readXMLFiles()
{
    Console.WriteLine("======Reading XML FILE======");
    XmlDocument doc = new XmlDocument();
    doc.Load("books.xml");

    XmlNodeList bookNodes = doc.DocumentElement.SelectNodes("/books/book");

    Console.WriteLine("Category".PadRight(15) + "Title".PadRight(25) + "Author(s)".PadRight(30) + "Year".PadRight(10) + "Price");
    Console.WriteLine("--------------------------------------------------------------------------------------------------------");

    foreach (XmlNode node in bookNodes)
    {
        string category = node.Attributes["category"].Value;
        string title = node.SelectSingleNode("title").InnerText;
        string authors = GetAuthors(node.SelectNodes("author"));
        string year = node.SelectSingleNode("year").InnerText;
        string price = node.SelectSingleNode("price").InnerText;

        Console.WriteLine($"{category.PadRight(15)}{title.PadRight(25)}{(authors.Length > 25 ? authors.Substring(0, 25) + "..." : authors).PadRight(30)}{year.PadRight(10)}{price}");
        Console.WriteLine(new string('-', 130));
    }
}

static string GetAuthors(XmlNodeList authorNodes)
{
    string authors = "";
    foreach (XmlNode authorNode in authorNodes)
    {
        authors += authorNode.InnerText + ", ";
    }
    return authors.TrimEnd(' ', ',');
}

void readCSVFiles()
{
    Console.WriteLine("==========ReadCSV==========");
    string filePath = "books.csv";

    using (TextFieldParser parser = new TextFieldParser(filePath))
    {
        parser.TextFieldType = FieldType.Delimited;
        parser.SetDelimiters(",");

        while (!parser.EndOfData)
        {
            string[] fields = parser.ReadFields();
            foreach (string field in fields)
            {
                Console.Write($"{field,-20}");
            }
            Console.WriteLine();
        }
    }
}

void readJSONFiles()
{
    Console.WriteLine("==========READJSON==========");
    string jsonText = File.ReadAllText("books.json");
      JObject bookstore = JObject.Parse(jsonText);
    JArray books = (JArray)bookstore["bookstore"]["book"];


    Console.WriteLine("Category".PadRight(15) + "Title".PadRight(25) + "Author(s)".PadRight(30) + "Year".PadRight(10) + "Price");
    Console.WriteLine(new string('-', 85));

    foreach (var book in books)
    {
        string category = book["_category"].ToString().PadRight(15);
        string title = book["title"]["__text"].ToString().PadRight(25);
        string authors = GetAuthors(book["author"]).PadRight(30);
        string year = book["year"].ToString().PadRight(10);
        string price = "$" + book["price"].ToString();

        Console.WriteLine(category + title + (authors.Length > 25 ? authors.Substring(0, 25) : authors) + "\t" + year + price);
    }
    string GetAuthors(JToken authorsToken)
    {
        if (authorsToken.Type == JTokenType.Array)
        {
            return string.Join(", ", authorsToken.Select(a => a.ToString()));
        }
        else
        {
            return authorsToken.ToString();
        }
    }
}

void storeJSONInExcel(List<Book>? books)
{
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    using (var package = new ExcelPackage())
    {
        var worksheet = package.Workbook.Worksheets.Add("Users");

        worksheet.Cells[1, 1].Value = "Category";
        worksheet.Cells[1, 2].Value = "Title";
        worksheet.Cells[1, 3].Value = "Author";
        worksheet.Cells[1, 4].Value = "Year";
        worksheet.Cells[1, 5].Value = "Price";

        int row = 2;
        foreach (var book in books)
        {
            worksheet.Cells[row, 1].Value = book.category;
            worksheet.Cells[row, 2].Value = book.title;
            worksheet.Cells[row, 3].Value = book.author;
            worksheet.Cells[row, 4].Value = book.year;
            worksheet.Cells[row, 5].Value = book.price;
            row++;
        }

        package.SaveAs(new System.IO.FileInfo("books.xlsx"));
    }

    Console.WriteLine("Excel file created successfully.");

}