using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleApp3;

public static class ReportExtension
{
    public static void GenerateReport(this WordprocessingDocument wdDoc)
    {
        Body rawBody = wdDoc.MainDocumentPart.Document.Body;

        rawBody.RebuildXmlTree();
        wdDoc.Save();
    }

    private static List<OpenXmlElement>? RebuildXmlTree(this Body rawBody)
    {
        var body = rawBody.ChildElements;
        //list of tags to interpret
        
        var tags = new List<string>() {"Rower"};

        var tagName = "holder";
        
        var placeHolders = rawBody.GetPlaceHolders(tagName);
        
        placeHolders.ProcessBody(tags);

        return new List<OpenXmlElement>();
        //return body;
    }

    private static List<OpenXmlElement> GetPlaceHolders(this OpenXmlElement elem, string tagName)
    {
        var list = new List<OpenXmlElement>();
        
        if (elem.HasTagName(tagName))
        {
            list.Add(elem);
            return list;
        }

        foreach (var childElement in elem.ChildElements)
        {
            list.AddRange(childElement.GetPlaceHolders(tagName));
        }

        return list;
    }

    private static bool HasTagName(this OpenXmlElement element, string tagName)
    {
        if (element.GetType() != typeof(SdtBlock))
            return false;
        var tag = (Tag)element.ChildElements.FirstOrDefault(x => x.GetType() == typeof(SdtProperties))?.ChildElements
            .FirstOrDefault(x => x.GetType() == typeof(Tag))!;
        return tag?.Val?.Value == tagName;
    }
    
    private static void FillWithChildren(this Body body, List<OpenXmlElement>? tree)
    {
        foreach (var element in tree)
        {
            body.AppendChild(element);
        }
    }

    private static OpenXmlElement? Find<T>(this OpenXmlElement elem)
    where T : OpenXmlElement
    {
        //TODO!!!!!!!!!!!!!!!!!!!!!!!!!!! return bo inaczej to będzie dupa -Rafał
        OpenXmlElement? searchedElement = null;
        
        foreach (var childElement in elem.ChildElements)
        {
            if (childElement.GetType() == typeof(T))
            {
                searchedElement = childElement;
            }
            else
            {
                searchedElement = childElement.Find<T>();
            }
        }
        return searchedElement;
    }
    
    private static void Substitute(this OpenXmlElement element, string textToSubstitute)
    {
        // var contentIndex = element.ChildElements.ToList().FindElementIndexByLocalName("sdtContent");
        // var runIndex = element.ChildElements[contentIndex].ChildElements[0].ToList().FindElementIndexByLocalName("r");
        
        // var sdtRun = element.ChildElements[contentIndex].ChildElements[0].ChildElements[runIndex].ChildElements;
        
        Text text = new Text();
        text.Text = textToSubstitute;

        // var textIndex = sdtRun.ToList().FindElementIndexByLocalName("t");
        // var item = (Text)sdtRun.GetItem(textIndex);

        var item = (Text) element.Find<Text>()!;
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
    
    private static void ProcessBody(this List<OpenXmlElement> placeholders, List<string> tags, string tagLocalName = "sdt")
    {
        //TODO substituting via interpreter
        var tempTextToSubstitute = "Mój przyjaciel prosiaczek";

        foreach (var t in placeholders)
            if (t.LocalName == tagLocalName && tags.Contains(t.InnerText))
            {
                t.Substitute(tempTextToSubstitute);
            }
    }
}