using Autodesk.Forge.DesignAutomation.Model;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.IO;

namespace Interaction
{
    /// <summary>
    /// Customizable part of Publisher class.
    /// </summary>
    internal partial class Publisher
    {
        /// <summary>
        /// Constants.
        /// </summary>
        private static class Constants
        {
            private const int EngineVersion = 2021;
            public static readonly string Engine = $"Autodesk.Inventor+{EngineVersion}";

            public const string Description = "Creates Wallshelf based on json file";

            internal static class Bundle
            {
                public static readonly string Id = "DA4ShelfBuilder";
                public const string Label = "alpha";

                public static readonly AppBundle Definition = new AppBundle
                {
                    Engine = Engine,
                    Id = Id,
                    Description = Description
                };
            }

            internal static class Activity
            {
                public static readonly string Id = Bundle.Id;
                public const string Label = Bundle.Label;
            }

            internal static class Parameters
            {
                public const string InventorDoc = nameof(InventorDoc);
                public const string OutputIam = nameof(OutputIam);
                public const string OutputPDF = nameof(OutputPDF);
                public const string InventorParams = nameof(InventorParams);
            }
        }


        /// <summary>
        /// Get command line for activity.
        /// </summary>
        private static List<string> GetActivityCommandLine()
        {
            //return new List<string> { $"$(engine.path)\\InventorCoreConsole.exe /i \"$(args[{Constants.Parameters.InventorDoc}].path)\" /al \"$(appbundles[{Constants.Activity.Id}].path)\" \"$(args[{Constants.Parameters.InventorParams}].path)\"" };
            // ovo je komandna linija koja ne šalje parametre
            return new List<string> { $"$(engine.path)\\InventorCoreConsole.exe /i \"$(args[{Constants.Parameters.InventorDoc}].path)\" /al \"$(appbundles[{Constants.Activity.Id}].path)\"" };
        }

        /// <summary>
        /// Get activity parameters.
        /// </summary>
        private static Dictionary<string, Parameter> GetActivityParams()
        {
            return new Dictionary<string, Parameter>
                {
                    {
                        Constants.Parameters.InventorDoc,
                        new Parameter
                        {
                            Verb = Verb.Get,
                            Zip = true,
                            LocalName="MyWallShelf.iam",
                            Description = "IAM file to process"
                        }
                    },
                    /* Smatram da ovo ne treba jer je json fajl poslat kroz zip
                    {
                        Constants.Parameters.InventorParams,
                        new Parameter
                        {
                            Verb=Verb.Put,
                            Description="JSON file with configuration",
                            LocalName ="params.json"
                        }
                    },*/
                    {
                        Constants.Parameters.OutputIam,
                        new Parameter
                        {
                            Verb = Verb.Put,
                            Zip = true,
                            LocalName = "Wall_shelf",
                            Description = "Resulting assembly"
                        }
                    }
                };
        }

        /// <summary>
        /// Get arguments for workitem.
        /// </summary>
        private static Dictionary<string, IArgument> GetWorkItemArgs(string bucketKey, string inputName, string paramFile, string outputName, string token)
        {
            string jsonPath = paramFile;
            JObject inputJSON = JObject.Parse(File.ReadAllText(jsonPath));
            string inputJsonString = inputJSON.ToString(Newtonsoft.Json.Formatting.None);

            return new Dictionary<string, IArgument>
            {
                {
                    Constants.Parameters.InventorDoc,
                    new XrefTreeArgument
                    {
                        Verb=Verb.Get,
                        LocalName="Wall_shelf",
                        PathInZip="MyWallShelf.iam",
                        Url = string.Format("https://developer.api.autodesk.com/oss/v2/buckets/{0}/objects/{1}", bucketKey, inputName),
                        Headers = new Dictionary<string, string>()
                        {
                            {"Authorization", "Bearer " + token }
                        }
                    }
                },
                /* Zašto da ubacujem parametre kroz liniju kad imam prosleđen taj isti fajl
                {
                    Constants.Parameters.InventorParams,
                    new XrefTreeArgument
                    {
                        Verb = Verb.Put,
                        Url = "data:application/json, " + inputJsonString
                    }
                },
                */
                {
                    Constants.Parameters.OutputIam,
                    new XrefTreeArgument
                    {
                        Url = string.Format("https://developer.api.autodesk.com/oss/v2/buckets/{0}/objects/{1}", bucketKey, outputName),
                        Verb=Verb.Put,
                        Headers=new Dictionary<string, string>()
                        {
                            { "Authorization", "Bearer " + token }
                        }
                    }
                }
            };
        }
    }
}
