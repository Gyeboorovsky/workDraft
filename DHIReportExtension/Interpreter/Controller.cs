using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DHIWordExtension.Interpreter.Models;
using Newtonsoft.Json;

namespace DHIWordExtension.Interpreter
{
    //Table.WSNM.Analyses { Filter:\"Diameter gt 50\", Sort:\"Diameter asc\", Select:\"Id Diameter MinorLoss HeadNodeId Roughness\" }
    //Table.WSNM.Analyses { Filter:"Diameter gt 50", Sort:"Diameter asc", Select:"Id Diameter MinorLoss HeadNodeId Roughness" }
    //Count.WSNM.Pipes { Filter=\"Diameter gt 50\" }
    
    public static class Controller
    {
        public static void InterpretingTest(string input)
        {
            input = "Table.WSNM.Analyses { Filter:\"Diameter gt 50\", Sort:\"Diameter asc\", Select:\"Id Diameter MinorLoss HeadNodeId Roughness\" }";
            ////////////////////////////////////////////////////////////////

            QueryModel queryObject = QueryMapping(input);

            Console.WriteLine(queryObject);
        }

        public static QueryModel QueryMapping(string input)
        {
            var bodyPart = Regex.Match(input, @"\{.*\}").ToString();
            var pathPart = Regex.Match(input, @"^([^\{ ])+").ToString();
            var pathSubParts = pathPart.Split('.');

            QueryModel queryObject = JsonConvert.DeserializeObject<QueryModel>(bodyPart);
            queryObject.QueryType =  (QueryType) Enum.Parse(typeof(QueryType), pathSubParts[0], true);
            queryObject.DataBaseName = pathSubParts[1];
            queryObject.TableName = pathSubParts[2];

            return queryObject;
        }
    }
}