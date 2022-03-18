using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleApp3;

public static class ReportExtension
{
    public static void GenerateReport(WordprocessingDocument wdDoc)
    {
        Body rawBody = wdDoc.MainDocumentPart.Document.Body;
        
        var newBody = RebuildXmlTree(rawBody);
        rawBody.RemoveAllChildren();
        rawBody.FillWithChildren(newBody);
    }

    private static List<OpenXmlElement>? RebuildXmlTree(Body rawBody)
    {
        var body = rawBody.ChildElements.ToList();
        //list of tags to interpret
        var tags = new List<string>() {"Rower"};
        body.ProcessBody(tags);

        return body;
    }

    private static void FillWithChildren(this Body body, List<OpenXmlElement>? tree)
    {
        foreach (var element in tree)
        {
            body.AppendChild(element);
        }
    }

    private static void Substitute(this OpenXmlElement element, string textToSubstitute)
    {
        var sdtRun = element.ChildElements[1].ChildElements[0].ChildElements[1].ChildElements;
        Text text = new Text();
        text.Text = textToSubstitute;
        var item = (Text)sdtRun.GetItem(1);
        item.Text = textToSubstitute;
        //sdtRun.RemoveAt(1);
        
        // element.SetAttribute();
        //sdtRun.Insert(1, text);

        Console.WriteLine();
    }

    private static int FindSdtContentElement(List<OpenXmlElement> elementList)
    {
        for (int i = 0; i < elementList.Count; i++)
        {
            if (elementList[i].LocalName == "sdtContent")
                return i;
        }

        throw new Exception("Element do not contain child element sdtContent");
    }
    
    private static void ProcessBody(this List<OpenXmlElement> body, List<string> tags, string tagLocalName = "sdt")
    {
        //TODO substituting via interpreter
        var tempTextToSubstitute = "Mój przyjaciel prosiaczek";

        for (int i = 0; i < body.Count; i++)
            if (body[i].LocalName == tagLocalName && tags.Contains(body[i].InnerText))
            {
                body[i].Substitute(tempTextToSubstitute);
            }
    }
}