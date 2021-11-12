using Inventor;
using System;
using System.IO;

namespace DebugPluginLocally
{
    class Program
    {
        static void Main()
        {
            using (var inv = new InventorConnector())
            {
                InventorServer server = inv.GetInventorServer();

                try
                {
                    Console.WriteLine("Running locally...");
                    // run the plugin
                    DebugSamplePlugin(server);
                }
                catch (Exception e)
                {
                    string message = $"Exception: {e.Message}";
                    if (e.InnerException != null)
                        message += $"{System.Environment.NewLine}    Inner exception: {e.InnerException.Message}";

                    Console.WriteLine(message);
                }
                finally
                {
                    if (System.Diagnostics.Debugger.IsAttached)
                    {
                        Console.WriteLine("Press any key to exit. All documents will be closed.");
                        Console.ReadKey();
                    }
                }
            }
        }

        /// <summary>
        /// Opens box.ipt and runs SamplePlugin
        /// </summary>
        /// <param name="app"></param>
        private static void DebugSamplePlugin(InventorServer app)
        {
            // basic part
            string componentSample = "MyWallShelf.iam";

            // get project directory
            string projectdir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;

            // get box.ipt absolute path
            string componentPath = System.IO.Path.Combine(projectdir, @"inputFiles\", componentSample);

            string componentPathCopy = System.IO.Path.Combine(projectdir, @"inputFiles\", "CopyOf" + componentSample);

            try
            {
                // delete an existing file
                System.IO.File.Delete(componentPathCopy);
            }
            catch (IOException)
            {
                Console.WriteLine("The specified file is in use. It might be open by Inventor");
                return;
            }

            // create a copy
            System.IO.File.Copy(componentPath, componentPathCopy);

            // open box.ipt by Inventor
            Document doc = app.Documents.Open(componentPath);

            // get params.json absolute path
            string paramsPath = System.IO.Path.Combine(projectdir, @"inputFiles\", "params.json");

            // create a name value map
            Inventor.NameValueMap map = app.TransientObjects.CreateNameValueMap();

            // add parameters into the map, do not change "_1". You may add more parameters "_2", "_3"...
            map.Add("_1", paramsPath);

            // add extra parameters to showcase newly supported parsing and new helper class NameValueMapHelper
            map.Add("intIndex", "1");
            map.Add("stringIndex", "test");
            map.Add("stringCollectionIndex", "str1, str2, str3");
            map.Add("intCollectionIndex", "34, 256, 9999, 500, 43");


            // create an instance of DA4ShelfBuilderPlugin
            DA4ShelfBuilderPlugin.SampleAutomation plugin = new DA4ShelfBuilderPlugin.SampleAutomation(app);

            // run the plugin
            plugin.RunWithArguments(doc, map);

        }
    }
}
