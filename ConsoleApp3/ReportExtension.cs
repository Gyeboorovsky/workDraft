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
        var contentIndex = element.ChildElements.ToList().FindElementIndexByLocalName("sdtContent");
        var runIndex = element.ChildElements[contentIndex].ChildElements[0].ToList().FindElementIndexByLocalName("r");

        var sdtRun = element.ChildElements[contentIndex].ChildElements[0].ChildElements[runIndex].ChildElements;
        
        Text text = new Text();
        text.Text = textToSubstitute;

        var textIndex = sdtRun.ToList().FindElementIndexByLocalName("t");
        var item = (Text)sdtRun.GetItem(textIndex);
        item.Text = textToSubstitute;
    }

    private static int FindElementIndexByLocalName(this List<OpenXmlElement> elementList, string localName)
    {
        for (int i = 0; i < elementList.Count; i++)
        {
            if (elementList[i].LocalName == localName)
                return i;
        }

        throw new Exception($"Element do not contain child element {localName}");
    }
    
    private static void ProcessBody(this List<OpenXmlElement> body, List<string> tags, string tagLocalName = "sdt")
    {
        //TODO substituting via interpreter
        var tempTextToSubstitute = "Gumowa kaczuszka";

        foreach (var t in body)
            if (t.LocalName == tagLocalName && tags.Contains(t.InnerText))
            {
                t.Substitute(tempTextToSubstitute);
            }
    }
}