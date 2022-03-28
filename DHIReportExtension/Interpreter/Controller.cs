using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;

namespace DHIWordExtension.Interpreter
{
    //Table.WSNM.Analyses { Filter="Diameter gt 50", Sort="Diameter asc", Select="Id Diameter MinorLoss HeadNodeId Roughness" }
    //Count.WSNM.Pipes { Filter="Diameter gt 50" }
    
    public static class Controller
    {
        public static void Interpret(string input)
        {
            input = "Count.WSNM.Pipes { Filter:\"Diameter gt 50\" }";
            ////////////////////////////////////////////////////////////////

            
            

            ExtractJson(input);
        }

        public static object ExtractJson(string input)
        {
            var extr = input.Split(' ').ToList();
            extr.RemoveAt(0);
            string dupa = "";
            foreach (var e in extr)
            {
                dupa += e;
            }
            
            
            DynamicCośTam obj = JsonConvert.DeserializeObject<DynamicCośTam>(dupa);

            
            
            return obj;
        }
    }

    public class DynamicCośTam
    {
        public string Filter { get; set; }
    }
}