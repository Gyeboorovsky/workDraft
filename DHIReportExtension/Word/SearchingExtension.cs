using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DHIReportExtension
{
    public static class SearchingExtension
    {
        public static List<OpenXmlElement> GetPlaceHolders(this OpenXmlElement elem, string tagName)
        {
            var list = new List<OpenXmlElement>();
            
            if (elem.HasTagName(tagName))
            {
                list.Add(elem);
                return list;
            }

            foreach (var childElement in elem.ChildElements)
            {
                if (childElement.LocalName == "")
                {
                    
                }
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
        
        public static List<OpenXmlElement> FindAllOf<T>(this OpenXmlElement elem)
            where T : OpenXmlElement
        {
            var list = new List<OpenXmlElement>();
            
            if (elem.GetType() == typeof(T))
            {
                list.Add(elem);
                return list;
            }

            foreach (var childElement in elem.ChildElements)
            {
                list.AddRange(childElement.FindAllOf<T>());
            }

            return list;
        }

        public static OpenXmlElement? FindFirstOfDefaultOf<T>(this OpenXmlElement elem)
            where T : OpenXmlElement
        {
            foreach (var childElement in elem.ChildElements)
            {
                if (childElement.GetType() == typeof(T))
                    return childElement;
                
                var result = childElement.FindFirstOfDefaultOf<T>();
                if (result?.GetType() == typeof(T))
                    return result;
            }
            
            return null;
        }
    }
}