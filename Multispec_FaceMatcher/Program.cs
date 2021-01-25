using System;
using Neurotec.Biometrics;
using Neurotec.Biometrics.Client;
using Neurotec.Licensing;
using System.IO;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

namespace Multispec_FaceMatcher
{
    class Program
    {
        const int Port = 5000;
        const string Address = "/local";
        static string[] faceLicenseComponents = { "Biometrics.FaceMatchingFast", "Biometrics.FaceSegmentsDetection" }; // "Biometrics.FaceExtractionFast"
        static NBiometricClient client;

        const string baseTempPath = @"C:\Users\John\Desktop\C#_Projects-Solutions\MultispecFiles\Templates\Sets\";
        const string savePath = @"C:\Users\John\Desktop\C#_Projects-Solutions\MultispecFiles\Data\";
        static int maxScore;

        static string[] galleryTemplates;

        static void Main(string[] args)
        {
            // Activate Licenses
            StartClient();
            
            Console.ReadLine(); // Pause for testing

            // Load in all gallery templates
            galleryTemplates = Directory.GetFiles(@"DIRECTORY PATH TO GALLERY TEMPLATES","*.dat",SearchOption.TopDirectoryOnly);


            // Read in all sets
            string[] dirSets = Directory.GetDirectories(baseTempPath);

            // Loop through sets
            foreach(string set in dirSets)
            {
                maxScore = 0;

                PerformMatching(set);
            }

            // Close out client
            EndClient();

            Console.ReadLine(); // Pause for testing
        }

        private static void StartClient()
        {
            foreach(string license in faceLicenseComponents)
            {
                if(NLicense.ObtainComponents(Address, Port, license))
                {
                    Console.WriteLine(string.Format("License was obtained: {0}",license));
                }
                else
                {
                    Console.WriteLine(string.Format("License was not obtained: {0}",license));
                }
            }

            Console.WriteLine();
            Console.Write("Initializing Client . . . . ");

            client = new NBiometricClient();
            client.BiometricTypes = NBiometricType.Face;
            client.Initialize();

            Console.WriteLine("Client Initialized!");
            Console.WriteLine();
        }

        private static void EndClient()
        {
            Console.WriteLine();

            foreach(string license in faceLicenseComponents)
            {
                Console.WriteLine(string.Format("Releasing license: {0}",license));
                NLicense.ReleaseComponents(license);
            }

            Console.WriteLine();
            Console.Write("Disposing Client . . . . ");

            client.Dispose();

            Console.WriteLine("Client Disposed!");
        }

        private static void PerformMatching(string probeDirectoryPath)
        {
            Dictionary<int,List<int>> genuineList = new Dictionary<int, List<int>>();
            Dictionary<int,List<int>> impostorList = new Dictionary<int, List<int>>();
            
            // Counter for the number of matches performed
            int matchCount = 0;

            /* VERSION 1 */
            /* Performing matching for entire set */

            /*
            // Get all scenarios and loop through them
            string[] dirScenarios = Directory.GetDirectories(probeDirectoryPath);
            foreach(string scenario in dirScenarios)
            {
                // Get all templates within scenario
                string[] probeTemplates = Directory.GetFiles(scenario,"*.dat",SearchOption.TopDirectoryOnly);

                // Loop through probes
                foreach(string probeFile in probeTemplates)
                {
                    string probeID = Path.GetFileNameWithoutExtension(probeFile);

                    // Create new subject from probe template
                    using(NSubject probe = NSubject.FromFile(probeFile))
                    {
                        // Loop through all gallery subjects for each probe template
                        foreach(string galleryFile in galleryTemplates)
                        {
                            string galleryID = Path.GetFileNameWithoutExtension(galleryFile);

                            // Create new gallery subject
                            using(NSubject gallery = NSubject.FromFile(galleryFile))
                            {
                                NBiometricStatus status = client.Verify(probe,gallery);
                                if(status == NBiometricStatus.Ok || status == NBiometricStatus.MatchNotFound)
                                {
                                    matchCount++;
                                    int score = probe.MatchingResults[0].Score;

                                    // Check maximum score
                                    if(score > maxScore)
                                        maxScore = score;

                                    // Temp value, would need to find quality score for the probe image
                                    int qualityScore = 0;
                                    List<int> qualityScoreList;

                                    // Genuine
                                    if(CheckGenuine(probeID,galleryID))
                                    {
                                        if(!genuineList.TryGetValue(score, out qualityScoreList))
                                        {
                                            // Could not get list for the given score
                                            qualityScoreList = new List<int>(); // Create new list
                                            genuineList.Add(score,qualityScoreList); // Add to genuine dictionary
                                        }

                                        qualityScoreList.Add(qualityScore);
                                    }

                                    // Impostor
                                    else
                                    {
                                        if(!impostorList.TryGetValue(score, out qualityScoreList))
                                        {
                                            // Could not get list for the given score
                                            qualityScoreList = new List<int>(); // Create new list
                                            impostorList.Add(score,qualityScoreList); // Add to genuine dictionary
                                        }

                                        qualityScoreList.Add(qualityScore);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Matching Error (Probe = {0}) (Gallery = {1})", probeID, galleryID);
                                }
                            }
                        }
                    }
                }
            }
            //*/

            /* VERSION 2 */
            /* All matching with respect to each individual scenario */

            //*
            // Get all templates within scenario
            string[] probeTemplates = Directory.GetFiles(probeDirectoryPath,"*.dat",SearchOption.TopDirectoryOnly);

            // Loop through probes
            foreach(string probeFile in probeTemplates)
            {
                string probeID = Path.GetFileNameWithoutExtension(probeFile); // GET THE ID FOR THE PROBE

                // Create new subject from probe template
                using(NSubject probe = NSubject.FromFile(probeFile))
                {
                    // Loop through all gallery subjects for each probe template
                    foreach(string galleryFile in galleryTemplates)
                    {
                        string galleryID = Path.GetFileNameWithoutExtension(galleryFile); // GET THE ID FOR THE GALLERY SUBJECT

                        // Create new gallery subject
                        using(NSubject gallery = NSubject.FromFile(galleryFile))
                        {
                            NBiometricStatus status = client.Verify(probe,gallery);
                            if(status == NBiometricStatus.Ok || status == NBiometricStatus.MatchNotFound)
                            {
                                matchCount++;
                                int score = probe.MatchingResults[0].Score;

                                // Check maximum score
                                if(score > maxScore)
                                    maxScore = score;

                                // Temp value, would need to find quality score for the probe image
                                //int qualityScore = 0;
                                List<int> qualityScoreList;

                                // Genuine
                                if(CheckGenuine(probeID,galleryID))
                                {
                                    // Try to get list from score
                                    if(!genuineList.TryGetValue(score, out qualityScoreList))
                                    {
                                        // Could not get list for the given score
                                        qualityScoreList = new List<int>(); // Create new list
                                        genuineList.Add(score,qualityScoreList); // Add to genuine dictionary
                                    }

                                    qualityScoreList.Add(score); // Add quality score to list (using match score now since we don't have quality scores yet)
                                }

                                // Impostor
                                else
                                {
                                    // Try to get list from score
                                    if(!impostorList.TryGetValue(score, out qualityScoreList))
                                    {
                                        // Could not get list for the given score
                                        qualityScoreList = new List<int>(); // Create new list
                                        impostorList.Add(score,qualityScoreList); // Add to genuine dictionary
                                    }

                                    qualityScoreList.Add(score); // Add quality score to list (using match score now since we don't have quality scores yet)
                                }
                            }
                            else
                            {
                                Console.WriteLine("Matching Error (Probe = {0}) (Gallery = {1})", probeID, galleryID);
                            }
                        }
                    }
                }
            }
            //*/

            /* SAVE EXCEL DATA */

            SaveData(matchCount, genuineList, impostorList, new DirectoryInfo(probeDirectoryPath).Name);

        }

        private static void SaveData(int count, Dictionary<int,List<int>> genuine, Dictionary<int,List<int>> impostor, string fileName)
        {
            /* EVALUATE DATA INTO 2D ARRAYS (better for fast Excel saving) */

            /*
            Row 1: Match Score
            Row 2: Genuine score probability
            Row 3: Impostor score probability
            Row 4: Quality score mean
            Row 5: Quality score median
            Row 6: Quality score standard deviation
            Row 7: Number of quality scores for that score
            */
            double[,] data = new double[3,maxScore+1]; //new double[7,maxScore+1];

            /* ********************************************** */

            // Loop through entire range of scores
            for(int i = 0; i <= maxScore; i++)
            {
                data[0,i] = i;

                List<int> matchList;

                // Try to get number of impostor matches
                if(impostor.TryGetValue(i, out matchList))
                    data[2,i] = matchList.Count / (double)count; // Impostor probability
                else
                    data[2,i] = 0;

                // Try to get number of genuine matches
                if(genuine.TryGetValue(i, out matchList))
                {
                    data[1,i] = matchList.Count / (double)count; // Genuine Probability

                    /*
                    double mean = GetMean(matchList);
                    int median = GetMedian(matchList);
                    double stdDev = GetStdDev(mean,matchList);

                    
                    data[3,i] = mean;
                    data[4,i] = median;
                    data[5,i] = stdDev;
                    data[6,i] = matchList.Count;
                    */
                }
                else
                {
                    data[1,i] = 0;

                    /*
                    data[3,i] = 0;
                    data[4,i] = 0;
                    data[5,i] = 0;
                    data[6,i] = 0;
                    */
                }

            }

            /* ********************************************** */

            /* CREATE AND SAVE EXCEL SHEETS WITH PROCESSED DATA */

            /* INITIALIZE EXCEL WORK */
            Excel.Application xlApp = new Excel.Application();

            if(xlApp == null)
            {
                Console.WriteLine("EXCEL ERROR ENCOUNTERED ! ! ! !");
            }
            else
            {
                Excel.Workbook xlWorkbook;
                Excel.Worksheet pdfWorksheet;
                //Excel.Worksheet qualityWorksheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkbook = xlApp.Workbooks.Add(misValue);

                /* *************************** */
                
                pdfWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.Item[1];
                pdfWorksheet.Name = "Multispec Data";

                pdfWorksheet.Cells[1,1] = "Match Score";
                pdfWorksheet.Cells[2,1] = "Genuine Probability";
                pdfWorksheet.Cells[3,1] = "Impostor Probability";
                /*
                pdfWorksheet.Cells[4,1] = "QltyScr Mean";
                pdfWorksheet.Cells[5,1] = "QltyScr Median";
                pdfWorksheet.Cells[6,1] = "QltyScr Std Dev";
                pdfWorksheet.Cells[7,1] = "QltyScr N";
                */

                Excel.Range pdfRange = (Excel.Range) pdfWorksheet.Cells[1,2];
                pdfRange = pdfRange.Resize[3,maxScore + 1]; //pdfRange.Resize[7,maxScore + 1]
                pdfRange.Value[Excel.XlRangeValueDataType.xlRangeValueDefault] = data;

                /* *************************** */
                /*
                qualityWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.Add(misValue, pdfWorksheet, misValue, misValue);
                qualityWorksheet.Name = "Quality Analysis";

                qualityWorksheet.Cells[1,1] = "Match Score";
                qualityWorksheet.Cells[2,1] = "Mean";
                qualityWorksheet.Cells[3,1] = "Median";
                qualityWorksheet.Cells[4,1] = "Std Dev";
                qualityWorksheet.Cells[5,1] = "N";

                Excel.Range qualityRange = (Excel.Range) qualityWorksheet.Cells[1,2];
                qualityRange = qualityRange.Resize[5,maxScore + 1];
                qualityRange.Value[Excel.XlRangeValueDataType.xlRangeValueDefault] = QualityAnalysis;
                //*/
                /* *************************** */

                // SAVE EXCEL FILE

                xlWorkbook.SaveAs(savePath + fileName + ".xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, true, misValue, misValue, misValue);
                xlWorkbook.Close(true, misValue, misValue);
                xlApp.Quit();

                /* *************************** */

                // Release all Excel objects
                //Marshal.ReleaseComObject(qualityWorksheet);
                Marshal.ReleaseComObject(pdfWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);
                Marshal.ReleaseComObject(xlApp);

                Console.WriteLine("The {0} Excel file was successfully written");
            }
            
            Console.WriteLine();
        }

        private static bool CheckGenuine(string probeFile, string galleryFile)
        {
            // TRUE  -> Genuine
            // FALSE -> Impostor

            // Note: file naming scheme follows: "RID_Date_22_Session#_MSFC_5DSR_Wavelength_..."
            string probeID = probeFile.Split('_')[0];
            string galleryID = galleryFile.Split('_')[0];

            return string.Equals(probeID, galleryID);
        }

        private static double GetMean(List<int> scoreList)
        {
            int sum = 0;
            foreach(int score in scoreList)
                sum += score;

            return sum / (double)scoreList.Count;
        }

        private static int GetMedian(List<int> scoreList)
        {
            scoreList.Sort();

            int index = (int)Math.Ceiling(scoreList.Count / (double)2) - 1;

            return scoreList[index];
        }

        private static double GetStdDev(double mean, List<int> scoreList)
        {
            double stdDev = 0;

            if(scoreList.Count != 1)
            {
                double sum = 0;

                foreach(int score in scoreList)
                    sum += Math.Pow(score - mean,2);

                stdDev = Math.Sqrt(sum / (scoreList.Count - 1));
            }

            return stdDev;
        }

        private static double GetQualityScore(string subjectID)
        {
            // Start at folder location of quality scores, if not incorporated into the filename

            // Read in file with matching filename/ID as subject

            // Read file using specified file format

            // Convert into a double

            // return double
            return 0;
        }
    }
}
