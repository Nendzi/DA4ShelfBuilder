using DA4ShelfBuilderPlugin.Models;
using System.IO;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Inventor;
using System;
using System.Diagnostics;

namespace DA4ShelfBuilderPlugin
{
    class WallShelfCreator
    {
        AssemblyDocument oDoc;
        InventorServer iApp;
        string projectPath;
        List<ShelfDataModel> proccessedShelfData = new List<ShelfDataModel>();

        public WallShelfCreator(AssemblyDocument oAsmDoc, InventorServer inventor)
        {
            oDoc = oAsmDoc;
            iApp = inventor;
        }
        public void Entry()
        {
            //projectPath = iApp.DesignProjectManager.ActiveDesignProject.WorkspacePath; //in local implamentation
            projectPath = System.IO.Path.GetDirectoryName(oDoc.FullFileName); // in Forge implementation
            Trace.TraceInformation($"Project folder is: {projectPath}");

            // collect all datas for shelfs
            List<ShelfDataModel> allShelfs = GetAllShelfs();

            // Analyze all geometry data. Lookup for shelfs in intersection
            // split shelfs on vertical orjentation and horizontal
            List<ShelfDataModel> verticalElement = new List<ShelfDataModel>();
            List<ShelfDataModel> horizontalElement = new List<ShelfDataModel>();

            foreach (ShelfDataModel item in allShelfs)
            {
                if (item.Orientation == "horizontal")
                {
                    horizontalElement.Add(item);
                }
                else
                {
                    verticalElement.Add(item);
                }
            }

            // for each vertical shelf lookup for intersection with every horizontal
            // there is intersection if midpoint coordinate of one shelf is between first and second points of another shelf
            int j = 0;
            double MidVX;
            double MidVY;
            double LenV;
            double MidHX;
            double MidHY;
            double LenH;

            do
            {
                for (int k = 0; k < horizontalElement.Count; k++)
                {
                    MidVX = verticalElement[j].MidPoint.X;
                    MidVY = verticalElement[j].MidPoint.Y;
                    LenV = verticalElement[j].Length;
                    MidHX = horizontalElement[k].MidPoint.X;
                    MidHY = horizontalElement[k].MidPoint.Y;
                    LenH = horizontalElement[k].Length;
                    if (IsBetween(MidVX, MidHX, LenH, false) && IsBetween(MidHY, MidVY, LenV,true))
                    {
                        // if ther is inter section, split vertical shelf on two pieces.
                        // create new element same as that shold be splited
                        // increment index of last vertical element
                        verticalElement.Add(new ShelfDataModel()
                        {
                            Depth = verticalElement[j].Depth,
                            Length = MidHY - (MidVY - LenV / 2),
                            Material = verticalElement[j].Material,
                            MidPoint = new MidPoint2DModel()
                            {
                                X = verticalElement[j].MidPoint.X,
                                Y = (MidVY + MidHY) / 2 - LenV / 4
                            },
                            Orientation = verticalElement[j].Orientation,
                            Thickness = verticalElement[j].Thickness
                        });
                        // set new Y koorinate in original (started) splited elements
                        verticalElement[j].MidPoint.Y = (MidVY + MidHY) / 2 + LenV / 4;
                        // set new lengths in original (strted) elements
                        verticalElement[j].Length = (MidVY + LenV / 2) - MidHY;
                    }
                }
                j++;
            } while (j < verticalElement.Count);

            proccessedShelfData.AddRange(horizontalElement);
            proccessedShelfData.AddRange(verticalElement);

            DetectAllConnections(horizontalElement.Count, verticalElement.Count);

            oDoc.ComponentDefinition.RepresentationsManager.DesignViewRepresentations["Default"].Activate();
            CreateShelfBasedOnData();
            CreateDrawing();
        }
        private bool IsBetween(double vx, double hx, double Length, bool ignore)
        {
            // ignore - ignoriše kontakte u krajnjim tačkama
            bool cond1;
            bool cond2;

            if (ignore)
            {
                cond1 = vx < hx + Length / 2;
                cond2 = hx - Length / 2 < vx; 
            }
            else
            {
                cond1 = vx <= hx + Length / 2;
                cond2 = hx - Length / 2 <= vx;
            }

            if (cond1 && cond2)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private void DetectAllConnections(int h, int v)
        {
            for (int i = 0; i < h; i++)
            {
                for (int j = h; j < h + v; j++)
                {
                    DetectConnection(i, j);
                }
            }
        }
        private void DetectConnection(int Hor, int Ver)
        {
            double HorMidX;
            double HorMidY;
            double HorLen;

            double VerTopY;
            double VerBottomY;
            double VerLen;

            ConnectionDataModel Data;

            HorMidX = proccessedShelfData[Hor].MidPoint.X;
            HorMidY = proccessedShelfData[Hor].MidPoint.Y;
            HorLen = proccessedShelfData[Hor].Length;

            VerLen = proccessedShelfData[Ver].Length;
            VerTopY = proccessedShelfData[Ver].MidPoint.Y + VerLen / 2;
            VerBottomY = proccessedShelfData[Ver].MidPoint.Y - VerLen / 2;

            if (VerTopY == HorMidY)
            {
                // we have connection from BOTTOM on position HorMidX
                proccessedShelfData[Ver].ConnectionOnBegin = true;
                Data = new ConnectionDataModel()
                {
                    Position = "bottom",
                    Distance = proccessedShelfData[Ver].MidPoint.X
                };
                proccessedShelfData[Hor].SetConnectionData(Data);
            }

            if (VerBottomY == HorMidY)
            {
                //we have connection from TOR on position HorMidX
                proccessedShelfData[Ver].ConnectionOnEnd = true;
                Data = new ConnectionDataModel()
                {
                    Position = "top",
                    Distance = proccessedShelfData[Ver].MidPoint.X
                };
                proccessedShelfData[Hor].SetConnectionData(Data);
            }
        }
        private List<ShelfDataModel> GetAllShelfs()
        {
            //string filePath = @"D:\TCuser_Redirect_Folders\Documents\Visual Studio 2019\Projects\Autodesk projects\InventorAddIns\DA4ShelfBuilder\DebugPluginLocally\inputFiles\";
            string paramFile = @"\params.json";
            string jsonString = System.IO.File.ReadAllText(projectPath + paramFile);
            JObject JSON = JObject.Parse(jsonString);
            JArray Shelf = (JArray)JSON["MyShelfData"];
            List<ShelfDataModel> allShelfs = Shelf.ToObject<List<ShelfDataModel>>();
            return allShelfs;
        }

        PartDocument oPart;
        PartComponentDefinition oPartDef;
        ComponentOccurrence oOcc;

        private void CreateShelfBasedOnData()
        {
            string progressString = ""; // "{'id': 'post test', 'value': 'something'}" ;
            int i = 1;
            foreach (ShelfDataModel item in proccessedShelfData)
            {
                PrepareTemplate();
                PlaceShelf(item, i);
                /*progressString += "{ 'current-progress': '";
                progressString += i.ToString();
                progressString += "', 'step': '";
                progressString += item.Orientation + "'}";
                Trace.TraceInformation(progressString);
                SampleAutomation.GetOnDemandFile("onProgress", "", "", progressString, null);*/
                Trace.TraceInformation("!ACESAPI:acesHttpOperation({0},\"\",\"\",{1},null)",
                        "onProgress",
                        "{ \"current-progress\": 30, \"step\": \"apply parameters\" }"
                        );
                i++;
            }
            oDoc.Save();
        }
        private void PrepareTemplate()
        {
            string oFileName = projectPath + @"\Shelf.ipt";

            oPart = iApp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, oFileName, false) as PartDocument;
            oPartDef = oPart.ComponentDefinition;
        }
        private void PlaceShelf(ShelfDataModel shelfElem, int rb)
        {
            const double piPola = Math.PI / 2;

            Matrix oMatrix = iApp.TransientGeometry.CreateMatrix();
            Point oPoint = iApp.TransientGeometry.CreatePoint(0, 0, 0);
            Vector oAxis = iApp.TransientGeometry.CreateVector(0, 0, 1);
            if (shelfElem.Orientation == "vertical")
            {
                oMatrix.SetToRotation(piPola, oAxis, oPoint);
            }

            oMatrix.SetTranslation(iApp.TransientGeometry.CreateVector(shelfElem.MidPoint.X / 10, shelfElem.MidPoint.Y / 10, 0));

            oOcc = oDoc.ComponentDefinition.Occurrences.AddByComponentDefinition(oPartDef as ComponentDefinition, oMatrix);

            //    add material to shelf
            AssignMaterial(oOcc, shelfElem.Material);
            CorrectDimensions(shelfElem.Length, shelfElem.Thickness, shelfElem.Depth, shelfElem.Orientation);

            //     suppress holes for front connection or make a new holes on top and bottom sides
            SuppressOrDrillHoles(shelfElem);

            PartComponentDefinition oPartComponentDefinition = oOcc.Definition as PartComponentDefinition;
            PartDocument oPartDocument = oPartComponentDefinition.Document as PartDocument;
            oPartDocument.SaveAs(projectPath + @"\Plate_" + rb.ToString() + ".ipt", false);

            oDoc.Update();
        }
        private void CorrectDimensions(double dWidth, double thk, double dDepth, string orj)
        {
            PartComponentDefinition oShelfDef = oOcc.Definition as PartComponentDefinition;
            ModelParameters oParams = oShelfDef.Parameters.ModelParameters;
            ModelParameter oParam = oParams["Length"];
            double lenValue;

            if (orj == "vertical")
            {
                lenValue = dWidth - thk;
            }
            else
            {
                lenValue = dWidth + thk;
            }
            oParam.Value = lenValue / 10;

            oParam = oParams["Thk"];
            oParam.Value = thk / 10;

            oParam = oParams["Depth"];
            oParam.Value = dDepth / 10;
        }

        // Create an object collection for the hole center points.
        ObjectCollection oHoleCentersForTop;
        ObjectCollection oHoleCentersForBottom;
        PlanarSketch oSketchTop;
        PlanarSketch oSketchBottom;

        private void SuppressOrDrillHoles(ShelfDataModel newShelfElement)
        {
            PartComponentDefinition oShelfDef = oOcc.Definition as PartComponentDefinition;

            HoleFeatures oHoleFeats = oShelfDef.Features.HoleFeatures;

            oHoleFeats["Hole1"].Suppressed = !newShelfElement.ConnectionOnBegin;
            oHoleFeats["Hole5"].Suppressed = !newShelfElement.ConnectionOnBegin;

            oHoleFeats["Hole3"].Suppressed = !newShelfElement.ConnectionOnEnd;
            oHoleFeats["Hole7"].Suppressed = !newShelfElement.ConnectionOnEnd;

            if (newShelfElement.Orientation == "horizontal")
            {
                oHoleCentersForTop = iApp.TransientObjects.CreateObjectCollection();
                oHoleCentersForBottom = iApp.TransientObjects.CreateObjectCollection();

                // Create a new sketch to contain the points that define the hole centers.
                oSketchTop = oShelfDef.Sketches.Add(oShelfDef.SurfaceBodies[1].Faces[6]);
                oSketchBottom = oShelfDef.Sketches.Add(oShelfDef.SurfaceBodies[1].Faces[1]);

                ConnectionDataModel cd = new ConnectionDataModel();
                for (int i = 0; i < newShelfElement.ConnectionList.Count; i++)
                {
                    cd = newShelfElement.ConnectionList[i];
                    if (cd.Position == "top")
                    {
                        CreateHoleCenters(-1, newShelfElement.ConnectionList[i].Distance, newShelfElement);
                    }
                    else
                    {
                        CreateHoleCenters(1, newShelfElement.ConnectionList[i].Distance, newShelfElement);
                    }

                }
                // Create the hole feature.
                if (oHoleCentersForTop.Count > 0)
                {
                    oShelfDef.Features.HoleFeatures.AddDrilledByDistanceExtent(oHoleCentersForTop, "8 mm", 0.7, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
                }
                if (oHoleCentersForBottom.Count > 0)
                {
                    oShelfDef.Features.HoleFeatures.AddDrilledByDistanceExtent(oHoleCentersForBottom, "8 mm", 0.7, PartFeatureExtentDirectionEnum.kPositiveExtentDirection);
                }
            }
            return;

            void CreateHoleCenters(int Corrector, double dist, ShelfDataModel nse)
            {
                double XLocation = (dist - nse.MidPoint.X) / 10;

                // Set a reference to the transient geometry object.
                Point2d Point1;
                Point2d Point2;

                if (Corrector == 1)
                {
                    // Add two points as hole centers.
                    Point1 = iApp.TransientGeometry.CreatePoint2d(XLocation, 4);
                    Point2 = iApp.TransientGeometry.CreatePoint2d(XLocation, (nse.Depth - 20) / 10);
                    oHoleCentersForBottom.Add(oSketchBottom.SketchPoints.Add(Point1));
                    oHoleCentersForBottom.Add(oSketchBottom.SketchPoints.Add(Point2));
                }
                else
                {
                    Point1 = iApp.TransientGeometry.CreatePoint2d(XLocation, -2);
                    Point2 = iApp.TransientGeometry.CreatePoint2d(XLocation, -(nse.Depth - 40) / 10);
                    oHoleCentersForTop.Add(oSketchTop.SketchPoints.Add(Point1));
                    oHoleCentersForTop.Add(oSketchTop.SketchPoints.Add(Point2));
                }
            }
        }
        private void AssignMaterial(ComponentOccurrence oOcc, int MaterialIndex)
        {
            Asset localAsset;
            string assetName = "Default";
            try
            {
                switch (MaterialIndex)
                {
                    case 1:
                        assetName = "English Oak"; break;
                    case 2:
                        assetName = "Maple"; break;
                    case 3:
                        assetName = "Red Birch"; break;
                    case 4:
                        assetName = "Teak"; break;
                    case 5:
                        assetName = "Wild Cherry - Java"; break;
                    case 6:
                        assetName = "Yellow Pine - Natural Polished"; break;
                    default:
                        assetName = "Default"; break;
                }

                localAsset = oDoc.Assets[assetName];
            }
            catch
            {
                AssetLibrary assetLib = iApp.AssetLibraries["Autodesk Appearance Library"];

                Asset libAsset = assetLib.AppearanceAssets[assetName];

                localAsset = libAsset.CopyTo(oDoc);
            }

            oOcc.Appearance = localAsset;
        }
        DrawingDocument oDrawDoc;
        private void CreateDrawing()
        {
            // Set a reference to the drawing document.
            // This assumes a drawing document is active.
            Trace.TraceInformation($"Project path in CreateDrawing is: {projectPath}");
            oDrawDoc = iApp.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, projectPath + @"\DrwTemp.idw",false) as DrawingDocument;
           
            // Set a reference to the BOM
            BOM oBOM = oDoc.ComponentDefinition.BOM;
            oBOM.PartsOnlyViewEnabled = true;
            oBOM.StructuredViewEnabled = true;

            //Set a reference to the "Structured" BOMView
            BOMView oBOMView = oBOM.BOMViews["Parts only"];

            PartDocument oFileName;
            double oScale;
            double oMp, oNp;

            //Set a position
            Point2d oPosition1, oPosition3;
            Sheet oSheet;

            Trace.TraceInformation($"Assembly has :{oBOMView.BOMRows.Count} parts");
            for (int i = 1; i <= oBOMView.BOMRows.Count; i++)
            {
                //set FileName
                oFileName = oBOMView.BOMRows[i].ComponentDefinitions[1].Document;
                _Document oViewModel = (_Document)oFileName;
                Point oMinPoint, oMaxPoint;
                oMinPoint = oFileName.ComponentDefinition.RangeBox.MinPoint;
                oMaxPoint = oFileName.ComponentDefinition.RangeBox.MaxPoint;

                double dX, dY, dZ;
                oScale = 0.5;
                dX = (oMaxPoint.X - oMinPoint.X) * oScale;
                dY = (oMaxPoint.Y - oMinPoint.Y) * oScale;
                dZ = (oMaxPoint.Z - oMinPoint.Z) * oScale;

                // calculate width for paper
                double paperWidth, paperHeigth;
                paperWidth = Math.Round(2 + 2 + dY + 2 + dX + 2 + 0.5, 0) + 1;
                paperHeigth = Math.Round(0.5 + 3 + 2 + dZ + 2 + dY + 2 + dZ + 2 + 0.5, 0) + 1;

                //Set a reference to the active sheet.
                //A draft view can only be created on an active sheet.
                PageOrientationTypeEnum paperOrjentation;
                if (paperWidth > paperHeigth)
                {
                    paperOrjentation = PageOrientationTypeEnum.kLandscapePageOrientation;
                }
                else
                {
                    paperOrjentation = PageOrientationTypeEnum.kPortraitPageOrientation;
                }
                oSheet = oDrawDoc.Sheets.Add(DrawingSheetSizeEnum.kCustomDrawingSheetSize, paperOrjentation, "", paperWidth, paperHeigth);

                //Set a reference to the drawing view.
                DrawingView oBaseView;
                DrawingView oProjectedView;

                oMp = paperWidth - 0.5 - 2 - dX / 2;
                oNp = paperHeigth - 0.5 - 2 - dZ / 2;

                oPosition1 = iApp.TransientGeometry.CreatePoint2d(oMp, oNp);
                oBaseView = oSheet.DrawingViews.AddBaseView(oViewModel, oPosition1, oScale, ViewOrientationTypeEnum.kTopViewOrientation, DrawingViewStyleEnum.kHiddenLineDrawingViewStyle);

                oPosition3 = iApp.TransientGeometry.CreatePoint2d(oMp - dX / 2 - 2 - dY / 2, oNp);
                oProjectedView = oSheet.DrawingViews.AddProjectedView(oBaseView, oPosition3, DrawingViewStyleEnum.kFromBaseDrawingViewStyle);
            }

            oDrawDoc.SaveAs(projectPath + @"\WallShelfDrawings.idw", false);

            PlaceHoleTable();
            PlaceDimensions();
            PlaceShelfAssemblyOnFirstSheet();

            oDrawDoc.Save();
            PublishPDF();
        }
        private void PlaceHoleTable()
        {
            Sheets oSheets = oDrawDoc.Sheets;
            Sheet oSheet;

            DrawingView oDrawingView;
            Point2d oPlacementPoint;
            GeometryIntent oDimIntent;
            Point2d oPointIntent;
            DrawingCurve oCurve;
            HoleTable oViewHoleTable;

            BorderDefinition oBorderDef = oDrawDoc.BorderDefinitions["Nena"];

            for (int i = 2; i <= oSheets.Count; i++)
            {
                oSheet = oSheets[i];
                oSheet.Activate();
                oSheet.AddBorder(oBorderDef);
                oSheet.AddTitleBlock("Nena");

                if (oSheet.DrawingViews.Count > 0)
                {
                    oDrawingView = oSheet.DrawingViews[1];
                    oCurve = BottomLine(oDrawingView.DrawingCurves);
                    oPointIntent = LeftPoint(oCurve);
                    oDimIntent = oSheet.CreateGeometryIntent(oCurve, oPointIntent);
                    if (!oDrawingView.HasOriginIndicator)
                    {
                        oDrawingView.CreateOriginIndicator(oDimIntent);
                    }
                    oDrawingView.OriginIndicator.Visible = true;
                    oPlacementPoint = iApp.TransientGeometry.CreatePoint2d(2, oSheet.Height / 2);
                    oViewHoleTable = oSheet.HoleTables.Add(oDrawingView, oPlacementPoint);
                }
            }
        }
        private DrawingCurve BottomLine(DrawingCurvesEnumerator oCurves)
        {
            int i = 0;
            double Xs;
            double Ys = 0;
            double Xe;
            double Ye = 0;

            bool cond1, cond2;

            DrawingCurve bl;
            foreach (DrawingCurve oCurve in oCurves)
            {
                if (IsCurveFromExtrude(oCurve))
                {
                    if (i == 0)
                    {
                        Xs = Math.Round(oCurve.StartPoint.X, 6);
                        Ys = Math.Round(oCurve.StartPoint.Y, 6);
                        Xe = Math.Round(oCurve.EndPoint.X, 6);
                        Ye = Math.Round(oCurve.EndPoint.Y, 6);
                        i++;
                        bl = oCurve;
                    }
                    else
                    {
                        cond1 = (Math.Round(oCurve.StartPoint.Y, 6) <= Ys);
                        cond2 = (Math.Round(oCurve.EndPoint.Y, 6) <= Ye);
                        if (cond1 && cond2)
                        {
                            bl = oCurve;
                            return bl;
                        }
                    }
                }
            }
            return null;
        }
        private Point2d LeftPoint(DrawingCurve oCurve)
        {
            if (oCurve.StartPoint.X < oCurve.EndPoint.X)
            {
                return oCurve.StartPoint;
            }
            else
            {
                return oCurve.EndPoint;
            }
        }
        private void PlaceShelfAssemblyOnFirstSheet()
        {
            Sheet oSheet = oDrawDoc.Sheets[1];
            _Document oModel = (_Document)oDoc;
            Point2d oPosition = iApp.TransientGeometry.CreatePoint2d(oSheet.Width / 2, oSheet.Height / 2);
            double oScale = 0.2;
            DrawingView oDrwView = oSheet.DrawingViews.AddBaseView(oModel, oPosition, oScale, ViewOrientationTypeEnum.kIsoTopRightViewOrientation, DrawingViewStyleEnum.kShadedDrawingViewStyle, "Default");
        }
        private void PlaceDimensions()
        {
            Sheets oSheets = oDrawDoc.Sheets;
            DrawingViews oDrwViews;
            DrawingCurvesEnumerator oDrwCurves;

            DrawingCurve[] ParalelCurvesCouples = new DrawingCurve[4];
            int i = 0;

            foreach (Sheet oSheet in oSheets)
            {
                oDrwViews = oSheet.DrawingViews;
                foreach (DrawingView oView in oDrwViews)
                {
                    oDrwCurves = oView.DrawingCurves;
                    foreach (DrawingCurve Curve in oDrwCurves)
                    {
                        if (IsCurveFromExtrude(Curve))
                        {
                            ParalelCurvesCouples[i] = Curve;
                            i++;
                        }
                    }
                    i = 0;
                    PlaceDimOnView(oSheet, oView, ParalelCurvesCouples);
                }
            }
        }
        private bool IsCurveFromExtrude(DrawingCurve Curve)
        {
            CurveTypeEnum curveType;
            Edge oEdge;
            bool output = false;

            if (Curve.ModelGeometry is Edge)
            {
                oEdge = Curve.ModelGeometry;
                curveType = oEdge.GeometryType;
                if (curveType == CurveTypeEnum.kLineSegmentCurve)
                {
                    output = true;
                }
            }

            return output;
        }
        private void PlaceDimOnView(Sheet oActiveSheet, DrawingView oDrawingView, DrawingCurve[] Curves)
        {
            DrawingCurve[] VertEdges = new DrawingCurve[2];
            DrawingCurve[] HorEdges = new DrawingCurve[2];

            // Create the x-axis vector
            Vector2d oXAxis = iApp.TransientGeometry.CreateVector2d(1, 0);

            Vector2d oCurveVector;

            int j = 0;
            int k = 0;

            for (int i = 0; i < 4; i++)
            {
                oCurveVector = Curves[i].StartPoint.VectorTo(Curves[i].EndPoint);
                if (oCurveVector.IsParallelTo(oXAxis))
                {
                    HorEdges[j] = Curves[i];
                    j++;
                }
                else
                {
                    VertEdges[k] = Curves[i];
                    k++;
                }
            }

            GeometryIntent oDimIntent1;
            GeometryIntent oDimIntent2;
            DimensionTypeEnum DimType;
            GeneralDimension oGeneralVerDimension;
            Point2d oTextOrigin1;

            // Set a reference to the general dimensions collection.
            GeneralDimensions oGeneralDimensions = oActiveSheet.DrawingDimensions.GeneralDimensions;

            // Create the horizontal dimension.
            oDimIntent1 = oActiveSheet.CreateGeometryIntent(VertEdges[0], PointIntentEnum.kStartPointIntent);
            oDimIntent2 = oActiveSheet.CreateGeometryIntent(VertEdges[1], PointIntentEnum.kEndPointIntent);

            oTextOrigin1 = iApp.TransientGeometry.CreatePoint2d(HorEdges[0].MidPoint.X, oDrawingView.Top + 1);

            DimType = DimensionTypeEnum.kHorizontalDimensionType;

            oGeneralVerDimension = oGeneralDimensions.AddLinear(oTextOrigin1, oDimIntent1, oDimIntent2, DimType) as GeneralDimension;

            // Create the vertikal dimension.
            oDimIntent1 = oActiveSheet.CreateGeometryIntent(HorEdges[0], PointIntentEnum.kStartPointIntent);
            oDimIntent2 = oActiveSheet.CreateGeometryIntent(HorEdges[1], PointIntentEnum.kEndPointIntent);

            oTextOrigin1 = iApp.TransientGeometry.CreatePoint2d(oDrawingView.Left - 1, VertEdges[0].MidPoint.Y);

            DimType = DimensionTypeEnum.kVerticalDimensionType;

            oGeneralVerDimension = oGeneralDimensions.AddLinear(oTextOrigin1, oDimIntent1, oDimIntent2, DimType) as GeneralDimension;
        }
        private void PublishPDF()
        {
            // Get the PDF translator Add-In.
            TranslatorAddIn PDFAddIn = iApp.ApplicationAddIns.ItemById["{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}"] as TranslatorAddIn;

            TranslationContext oContext = iApp.TransientObjects.CreateTranslationContext();
            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;

            // Create a NameValueMap object
            NameValueMap oOptions = iApp.TransientObjects.CreateNameValueMap();

            // Create a DataMedium object
            DataMedium oDataMedium = iApp.TransientObjects.CreateDataMedium();

            // Check whether the translator has //SaveCopyAs// options
            if (PDFAddIn.HasSaveCopyAsOptions[oDrawDoc, oContext, oOptions])
            {
                // Options for drawings...

                oOptions.Value["All_Color_AS_Black"] = 0;
                oOptions.Value["Remove_Line_Weights"] = 0;
                oOptions.Value["Vector_Resolution"] = 400;
                oOptions.Value["Sheet_Range"] = PrintRangeEnum.kPrintAllSheets;
                //oOptions.Value["Custom_Begin_Sheet"] = 2;
                //oOptions.Value["Custom_End_Sheet"] = 4;
            }

            //Set the destination file name
            oDataMedium.FileName = projectPath + @"\test.pdf";

            //Publish document.
            PDFAddIn.SaveCopyAs(oDrawDoc, oContext, oOptions, oDataMedium);
        }
    }
}

