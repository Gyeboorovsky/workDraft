using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Xml;

namespace Adapter;

//KOD Z ZEWNĘTRZNEJ BIBLIOTEKI
public class UsersApi
{
    public async Task<string> GetUsersXmlAsync()
    {
        var apiResponse = "<?xml version= \"1.0\" encoding= \"UTF-8\"?><users><user name=\"John\" surname=\"Doe\"/><user name=\"John\" surname=\"Wayne\"/><user name=\"John\" surname=\"Rambo\"/></users>";

        XmlDocument doc = new XmlDocument();
        doc.LoadXml(apiResponse);

        return await Task.FromResult(doc.InnerXml);
    }

    public string GetUserCsv()
    {
        char[] result;
        
        string fileName = "users.csv";
        string path = Path.Combine(Environment.CurrentDirectory, @"../../../", fileName);

        using (StreamReader reader = File.OpenText(fileName))
        {
            result = new char[reader.BaseStream.Length];
            reader.ReadAsync(result, 0, (int)reader.BaseStream.Length);
        }

        StringBuilder builder = new StringBuilder();
        foreach (char c in result)
        {
            
        }

        // using (var csvReader = new CsvReader(new StreamReader(System.IO.File.OpenRead(@"D:\CSVFolder\CSVFile.csv")), true))
        // {
        //     csvTable.Load(csvReader);
        // }
        
        // using (TextFieldParser parser = new TextFieldParser(path))
        // {
        //     parser.TextFieldType = FieldType.Delimited;
        //     parser.SetDelimiters(",");
        //     while (!parser.EndOfData) 
        //     {
        //         //Processing row
        //         string[] fields = parser.ReadFields();
        //         csvData += $"{fields[0]} {fields[1]};";
        //     }
        // }
        
        return csvData;
    }
}

public interface IUserRepository
{
    List<string> GetUserNames();
}

public class UserReposytoryCsvAdapter : IUserRepository
{
    private UsersApi _adaptee = null;

    public UserReposytoryCsvAdapter(UsersApi adaptee)
    {
        _adaptee = adaptee;
    }

    public List<string> GetUserNames()
    {
        List<string> userNames = null;
        string incompatibleApiResponse = this._adaptee.GetUserCsv();
        userNames = incompatibleApiResponse.Split(';').ToList();
        return userNames;
    }
}

public class UsersRepositoryAdapter : IUserRepository
{
    private UsersApi _adaptee = null;

    public UsersRepositoryAdapter(UsersApi adaptee)
    {
        _adaptee = adaptee;
    }

    public List<string> GetUserNames()
    {
        string incompatibleApiResponse = this._adaptee
          .GetUsersXmlAsync()
          .GetAwaiter()
          .GetResult();

        XmlDocument doc = new XmlDocument();
        doc.LoadXml(incompatibleApiResponse);

        var rootEl = doc.LastChild;

        List<string> userNames = new List<string>();

        if (rootEl.HasChildNodes)
        {
            foreach (XmlNode user in rootEl.ChildNodes)
            {
                userNames.Add(user.Attributes["name"].InnerText + user.Attributes["surname"].InnerText);
            }
        }
        return userNames;
    }
}