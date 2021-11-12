/////////////////////////////////////////////////////////////////////
// Copyright (c) Autodesk, Inc. All rights reserved
// Written by Forge Partner Development
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
/////////////////////////////////////////////////////////////////////

using Autodesk.Forge.DesignAutomation.Inventor.Utils;
using Autodesk.Forge.DesignAutomation.Inventor.Utils.Helpers;
using Inventor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace DA4ShelfBuilderPlugin
{
    [ComVisible(true)]
    public class SampleAutomation
    {
        private readonly InventorServer inventorApplication;

        public SampleAutomation(InventorServer inventorApp)
        {
            inventorApplication = inventorApp;
        }

        public void Run(Document doc)
        {
            LogTrace("Run me called with {0}", doc.DisplayName);
            LogTrace("asd");

            if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
            {
                using (new HeartBeat())
                {
                    // TODO: handle the Inventor part here
                }
            }
            else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject) // Assembly.
            {
                
                using (new HeartBeat())
                {
                    //GetOnDemandFile("onProgress", "", "", "{'id': 'post test', 'value': 'something'}", null);
                    
                    WallShelfCreator WSC = new WallShelfCreator(doc as AssemblyDocument, inventorApplication);
                    WSC.Entry();
                }
            }
        }

        public void RunWithArguments(Document doc, NameValueMap map)
        {
            LogTrace("Processing " + doc.FullFileName);

            try
            {
                // Using NameValueMapExtension
                if (map.HasKey("intIndex"))
                {
                    int intValue = map.AsInt("intIndex");
                    LogTrace($"Value of intIndex is: {intValue}");
                }

                if (map.HasKey("stringCollectionIndex"))
                {
                    IEnumerable<string> strCollection = map.AsStringCollection("stringCollectionIndex");

                    foreach (string strValue in strCollection)
                    {
                        LogTrace($"String value is: {strValue}");
                    }
                }

                if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    using (new HeartBeat())
                    {
                        // TODO: handle the Inventor part here
                    }
                }
                else if (doc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject) // Assembly.
                {
                    using (new HeartBeat())
                    {
                        WallShelfCreator WSC = new WallShelfCreator(doc as AssemblyDocument, inventorApplication);
                        WSC.Entry();
                    }
                }
            }
            catch (Exception e)
            {
                LogError("Processing failed. " + e.ToString());
            }
        }

        #region Logging utilities

        /// <summary>
        /// Call ACESAPI to get response from Forge server
        /// </summary>
        /// <param name="name">name of the onDemand input parameter as specified in the Activity</param>
        /// <param name="suffix">a query string - optional parameters that can be addded to the url defined in the WorkItem</param>
        /// <param name="headers">http call headers</param>
        /// <param name="data">name of file to save the response to</param>
        /// <param name="responseFile">name of file to save the response to or null</param>
        public static void GetOnDemandFile(string name, 
                                           string suffix, 
                                           string headers, 
                                           string data, 
                                           string responseFile)
        {
            LogTrace("Sending a POST request");
            LogTrace("!ACESAPI:acesHttpOperation({0},{1},{2},{3},{4})",
            name ?? "", suffix ?? "", headers ?? "", data ?? "", responseFile);
            LogTrace("Log after ACESAPI");

        }

        /// <summary>
        /// Log message with 'trace' log level.
        /// </summary>
        private static void LogTrace(string format, params object[] args)
        {
            Trace.TraceInformation(format, args);
        }

        /// <summary>
        /// Log message with 'trace' log level.
        /// </summary>
        private static void LogTrace(string message)
        {
            Trace.TraceInformation(message);
        }

        /// <summary>
        /// Log message with 'error' log level.
        /// </summary>
        private static void LogError(string format, params object[] args)
        {
            Trace.TraceError(format, args);
        }

        /// <summary>
        /// Log message with 'error' log level.
        /// </summary>
        private static void LogError(string message)
        {
            Trace.TraceError(message);
        }

        #endregion
    }
}