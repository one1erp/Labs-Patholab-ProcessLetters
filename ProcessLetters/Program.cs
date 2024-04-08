
using FAXCOMLib;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using Oracle.ManagedDataAccess.Client;
using Patholab_Common;
using Patholab_DAL_V1;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
using SendAssutaResponse;
using SendClalitFinalResponse;
using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using Exception = System.Exception;
//using Application = Microsoft.Office.Interop.Word.Application;
using oApplication = Microsoft.Office.Interop.Outlook.Application;

//for logging

namespace ProcessLetters
{

    class Program
    {
        // private static string PrintDir;
        private static string PDFDir;
        private static string XmlDir;
        private static string HM_Dir;
        private static string Maligant_Dir;
        private static string Outlook_Template;

        private static OracleConnection _connection;

        // C# doesn't have optional arguments so we'll need a dummy value
        private static object oMissing = System.Reflection.Missing.Value;
        //private static Microsoft.Office.Interop.Word.Application word = null;
        private static Microsoft.Office.Interop.Outlook.Application outlookApp = null;
        private static DataLayer dal;
        private static int maxPrintAttempts;
        private static FAXCOMLib.FaxServer faxServer = new FaxServer();
        private static string _faxServerName;
        private static string _adobeAcrobatLocation;
        private static OracleCommand _cmd;
        private static SendAssutaResponse.SendResponse assutaResponse;
        //    private static Boolean debug = false;
        private static bool UseFax;



        static void Main(string[] args)
        {

            //   if (debug)
            //   { Console.WriteLine("In debug mode"); }
            string machineName = "";
            dal = new DataLayer();
            string connectionStrings = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;
            maxPrintAttempts = int.Parse(ConfigurationManager.AppSettings["maxPrintAttempts"]);
            UseFax = (ConfigurationManager.AppSettings["UseFax"].ToString()).ToUpper() == "TRUE";
            if (UseFax)
            {
                _faxServerName = ConfigurationManager.AppSettings["faxServerName"];
                machineName = _faxServerName;// "VM-RR";
                if (machineName == "THIS-PC") machineName = Environment.MachineName;
            }
            Log("Finding Adobe Acrobt PDF Reader");
            //var adobe = Registry.LocalMachine.OpenSubKey("Software").OpenSubKey("Microsoft").OpenSubKey("Windows").OpenSubKey("CurrentVersion").OpenSubKey("App Paths").OpenSubKey("AcroRd32.exe");
            //var path = adobe.GetValue("");
            //if (path.ToString() != "")
            //{
            //    _adobeAcrobatLocation = path.ToString();
            //}
            //else
            //{
            //    var adobeOtherWay = Registry.LocalMachine.OpenSubKey("Software").OpenSubKey("Classes").OpenSubKey("acrobat").OpenSubKey("shell").OpenSubKey("open").OpenSubKey("command");
            //    var pathOtherWay = adobeOtherWay.GetValue("");
            //    if (pathOtherWay.ToString() != "")
            //    {
            //        _adobeAcrobatLocation = pathOtherWay.ToString();
            //    }
            //    else
            //    {
            //        Log("!!!! CRITICAL ERROR !!!! Can't find Adobe Acrobt PDF Reader. Exiting program!");
            //        Console.ReadKey();

            //        return;
            //    }
            //}
            if (UseFax)
            {


                Log("Connecting To '" + machineName + "' Fax Server");
                try
                {
                    faxServer.Connect(machineName);
                    Log("Connected to Fax Server");
                }
                catch (Exception ex)
                {

                    Log("!!!! CRITICAL ERROR !!!! Can't connect to fax server. Exiting program!");
                    Console.ReadKey();

                    return;
                }
            }
            //  var connectionStrings = "Data Source=PATHOLAB;User ID=lims_sys;Password=lims_sys";
            // this patrt works from app.config
            Log("Connecting to DB");
            dal.MockConnect();//ZMANI instal avigail
            Log("Connected to DB DAL");

            try
            {
                string connectionStringOracle = ConfigurationManager.ConnectionStrings["connectionStringOracle"].ConnectionString;

                _connection = new OracleConnection(connectionStringOracle);
                _connection.Open();
            }
            catch (Exception EXP)
            {
                Log("!!!! CRITICAL ERROR !!!! Can't connect to DB with oracle. Exiting program!");
                Log(EXP);
                Console.ReadKey();
                return;
            }

            Log("Connected to DB Oracle");
            try
            {

                assutaResponse = new SendResponse(dal);
                Log("Connected to Assuta rsponse maker");
            }
            catch (Exception exp1)
            {
                Log("!!!! CRITICAL ERROR !!!! Can't start assuta response. Exiting program!");
                Log(exp1);
                Console.ReadKey();
                return;
            }
            //       PrintDir = @"p:\patholab\print\";
            bool pdfDirExistsInParameters = false;
            try
            {
                PHRASE_HEADER SystemParams = dal.GetPhraseByName("System Parameters");
                pdfDirExistsInParameters =
                   SystemParams.PhraseEntriesDictonary.TryGetValue("PDF Directory", out PDFDir) &&
                SystemParams.PhraseEntriesDictonary.TryGetValue("XML Directory", out XmlDir);
                if (!PDFDir.EndsWith(@"\")) PDFDir += @"\";
                if (!XmlDir.EndsWith(@"\")) XmlDir += @"\";

                if (SystemParams.PhraseEntriesDictonary.TryGetValue("Ministry of Health Directory", out HM_Dir))
                    if (!HM_Dir.EndsWith(@"\")) HM_Dir += @"\";

                if (SystemParams.PhraseEntriesDictonary.TryGetValue("CopyMalignantDir", out Maligant_Dir))
                    if (!Maligant_Dir.EndsWith(@"\"))
                        Maligant_Dir += @"\";

                if (SystemParams.PhraseEntriesDictonary.TryGetValue("Outlook Response Template", out Outlook_Template))
                    if (!Outlook_Template.EndsWith(@"\"))
                    {
                        //Error   Outlook_Template += @"\"; 7/9/20
                    }



            }
            catch (Exception ex)
            {
                Log("Error in getting system params\r\n");
                Log(ex);
                if (args.Count() != 0)
                {
                    Console.ReadKey();
                }
                return;

            }
            if (!pdfDirExistsInParameters)
            {
                Exception ex;
                Log("Error in getting system params\r\n");
                Log(ex = new Exception(@"Error: Could not find entries  ""PDF Directory"" or ""XML Directory"" in Phrase ""System Parameters"""));

                if (args.Count() != 0)
                {
                    Console.ReadKey();
                }
                return;
            }

            try
            {
                Log("Processing PDF directory:" + PDFDir + "\r\nScaning every second\r\n--------------");
                while (true)
                {
                    ProcessPDFDirectory(PDFDir);
                    Console.Write(".");
                    Thread.Sleep(1000);
                }
            }
            catch (Exception e)
            {
                Log(e);
            }
            finally
            {
                //   ((Microsoft.Office.Interop.Outlook._Application)outlookApp).Quit();
            }
            outlookApp = null;
            //for debug

            Console.ReadKey();
        }

        public static void ProcessPDFDirectory(string dirPath)
        {
            if (dirPath == null)
            {
                Log(new Exception("Error! PDF directory is null"));
                return;
            }

            //רוץ על כל הקבצים בספרי הנוכחית. בספריה ראשית שם המחשב ריק
            string[] fileEntries = Directory.GetFiles(dirPath);
            foreach (string fileName in fileEntries)
                if (Path.GetExtension(fileName).ToUpper() == ".PDF"
                    && Path.GetFileName(fileName).Substring(0, 1) != "$"
                    && Path.GetFileName(fileName).Substring(0, 1) != "~")
                {
                    ProcessPDFFile(fileName);
                }

        }

        public static void ProcessPDFFile(string path)
        {
            Log("\r\nProcessing file '" + path + "'.");
            FileStream stream = null;
            bool isOpen = false;
            string typeCSV;
            try
            {
                stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                isOpen = true;
                Console.WriteLine("File '" + path + "' is locked for editing and will not be processed");
                if (stream != null)
                    stream.Close();
                return;
                //Show your prompt here.
            }
            finally
            {
                if (stream != null)
                    stream.Close();

            }

            //FileInfo pdfFile = new FileInfo(path);
            //Object filename = (Object)pdfFile.FullName;
            /// get the wreport and destinations for pdf
            /// there is an option for multiple wreports.

            string[] typeArray;

            string[] saveAsArray;
            string[] copiesArray;
            string[] wrDestinationIdArray;
            long sdgId = 0;
            int workflowNodeId;
            try
            {
                string fileName = Path.GetFileNameWithoutExtension(path);
                if (!long.TryParse(fileName.Split('_')[0], out sdgId))
                {
                    Log("Cannot open file '" + path + "'. The name should be the 'SDG ID_Workflow node id'.");
                    return;
                }
                workflowNodeId = int.Parse(fileName.Split('_')[1]);
            }
            catch (Exception cof)
            {

                Log("Cannot open file '" + path + "'. The name should be the 'SDG ID_Workflow node id'.");
                Log(cof);
                return;
            }

            SDG sdg;
            string sdgStatus;
            DateTime? sdgAuthorisedOn = null;
            try
            {
                sdg = dal.FindBy<SDG>(d => d.SDG_ID == sdgId).FirstOrDefault();
                Console.WriteLine("Opened SDG name '" + sdg.NAME + "'" + "Patholab Name :" +
                                  sdg.SDG_USER.U_PATHOLAB_NUMBER ?? "");
            }
            catch (Exception cof1)
            {
                Log(cof1);
                Log("Cannot open file '" + path + "'. The SDG was not found.");
                return;
            }
            //sdg.STATUS != "A" ||
            if (sdg.SDG_USER.U_PDF_PATH != null)
            {
                //delete authorised and printed pdf, they should not be here
                try
                {
                    Console.WriteLine("The SDG has a 'PDF_PATH'. The file will be deleted. it was not supposed to be here at all.");
                    File.Delete(path);
                }
                catch (Exception ex)
                {

                    Log(ex);
                }
                return;
            }
            WORKFLOW_NODE extentionNode = dal.FindBy<WORKFLOW_NODE>(wn => wn.WORKFLOW_NODE_ID == workflowNodeId).SingleOrDefault();
            sdgStatus = sdg.STATUS;
            string wreportId;
            U_WREPORT[] wreports =
               dal.FindBy<U_WREPORT>(wr => wr.U_WREPORT_USER.U_WORKFLOW_EVENT == extentionNode.PARENT_NODE.NAME ||
                 wr.U_WREPORT_USER.U_WORKFLOW_EVENT == "Print PDF Letter").ToArray(); //roy fort debug
            if (extentionNode.PARENT_NODE.NAME == "Authorised")
            // (extentionNode.PARENT_NODE.NAME == "A" || extentionNode.PARENT_NODE.NAME == "ToAuthorise")
            {
                wreports =
                dal.FindBy<U_WREPORT>(wr => wr.U_WREPORT_USER.U_WORKFLOW_EVENT == "Print PDF Letter").ToArray();
                if (sdgStatus != "A") sdgStatus = "A";
                sdgAuthorisedOn = sdg.AUTHORISED_ON ?? dal.GetSysdate();

            }
            bool isError = false;

            foreach (U_WREPORT wreport in wreports)
            {
                if (
                    !(";" + wreport.U_WREPORT_USER.U_WORKFLOW_NAME + ";").Contains(";" + extentionNode.WORKFLOW.NAME +
                                                                                   ";"))
                    continue;
                try
                {
                    if (wreport.U_WRDESTINATION_USER == null)
                    {
                        Log("Error:Can not find the print/save Destination for wreport '" +
                                          wreport.NAME + "' ");
                        continue;
                    }
                    wreportId = wreport.U_WREPORT_ID.ToString();
                }
                catch (Exception ex)
                {
                    Log(ex);
                    Log("Error:Can not find the process Destination for wreport '" +
                                        wreport.NAME + "' ");
                    continue;
                }


                SDG_USER sdgUser = sdg.SDG_USER;
                string pathToLastPDFWithWatermark = null;
                try
                {
                    //deal with revisions
                    try
                    {
                        if (sdgUser.U_IS_LAST_UPDATE == "T")
                        {
                            //find the last revision
                            SDG sDGlastRevision =
                                dal.FindBy<SDG>(
                                    d =>
                                    d.NAME.Substring(0, 10) == sdgUser.SDG.NAME &&
                                    d.SDG_USER.U_IS_LAST_UPDATE == "F" && d.SDG_USER.U_PDF_PATH != null)
                                   .OrderByDescending(d => d.SDG_ID).FirstOrDefault();
                            if (sDGlastRevision != null && File.Exists(sDGlastRevision.SDG_USER.U_PDF_PATH))
                            {
                                string pathToLastPDF = sDGlastRevision.SDG_USER.U_PDF_PATH;
                                pathToLastPDFWithWatermark = Path.GetDirectoryName(pathToLastPDF) + @"\" + Path.GetFileNameWithoutExtension(pathToLastPDF) +
                                                             @"_WM.pdf";
                                File.Delete(pathToLastPDFWithWatermark);
                                if (AddWatermark(pathToLastPDF, pathToLastPDFWithWatermark))
                                {

                                }
                                else
                                {

                                    File.Delete(pathToLastPDFWithWatermark);
                                    pathToLastPDFWithWatermark = null;
                                }
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        Log(ex);
                        Log("Erorr Adding Revision PDf:" + ex.ToString());
                    }
                    //compress

                    if (pathToLastPDFWithWatermark != null)
                    {
                        System.Diagnostics.Process process = new System.Diagnostics.Process();
                        System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                        startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;

                        //maybe i can overwright the file
                        string tempPdf = path + "F";
                        startInfo.FileName = ConfigurationManager.AppSettings["ghostscriptgswin32cFullPath"];
                        startInfo.Arguments = @"-sOutputFile=""" + tempPdf + "\" " +
                                              ConfigurationManager.AppSettings["ghostscriptArguments"] +
                                              " -f\"" + (string)path + "\"";

                        startInfo.Arguments += " \"" + pathToLastPDFWithWatermark + "\"";

                        process.StartInfo = startInfo;
                        process.Start();
                        process.WaitForExit();
                        try
                        {
                            File.Delete(path);
                            File.Move(tempPdf, path);
                            //File.Delete(tempPdf);
                        }
                        catch (Exception ex)
                        {

                            Log(ex);
                        }

                    }
                    if (pathToLastPDFWithWatermark != null)
                    {
                        File.Delete(pathToLastPDFWithWatermark);
                        pathToLastPDFWithWatermark = null;
                    }
                }
                catch (Exception ex)
                {
                    Log("error compressing pdf or adding revision pdf");
                    Log(ex);
                    //File.Delete((string)tempPdf);
                    continue;
                }
                Log("Initial PDF Created");
                dal.InsertToSdgLog(sdgId, "PDF.D", 0, wreport.U_WREPORT_ID.ToString());



                // רושמים ללוג
                //רושמים לSDG את המיקום 
                // כאשר נשלח פקס או מייל או הודפס רושמים לסדג לוג
                //במידה ויש שגיאה רושמים ללוג ושולחים מייללמנהל
                //במידה והודפס או נשלח מעדכנים את השדות של היעד בסדג
                if (wreport.U_WRDESTINATION_USER == null)
                {
                    Console.WriteLine("Cannot process Wreport '" + wreport.NAME + "'. No destination Connected to Wreport.");
                    continue;
                }
                U_WRDESTINATION_USER[] destinations = wreport.U_WRDESTINATION_USER.ToArray();
                string currentReportType = "";
                foreach (U_WRDESTINATION_USER destination in destinations)
                {
                    try
                    {
                        currentReportType = ExecuteOrGetString(destination.U_TYPE, sdg, sdgStatus);



                        string deviceName = ExecuteOrGetString(destination.U_DEVICE_NAME, sdg, sdgStatus);
                        if (currentReportType == "" || deviceName == "") continue;
                        string saveAs = deviceName;



                        string outputFileName = "";
                        Log("Processing " + currentReportType.ToUpper() + " Command");
                        switch (currentReportType.ToUpper())
                        {
                            case "PDF": //pdf 
                                outputFileName = saveAs;
                                //in case of an error making the filename, the name will be sdg id.PDF. deal with an errror by sending it to the error dir
                                if (outputFileName.ToString().ToUpper() == sdgId.ToString() + ".PDF")
                                {
                                    CopyPdf2Error(path, sdgUser, currentReportType);
                                    deviceName = "";
                                }
                                else //Is good
                                {
                                    DateTime timeForFolder;
                                    //if sdgUser.SDG.AUTHORISED_ON = null use created on or NOW
                                    timeForFolder = sdgAuthorisedOn ?? sdgUser.SDG.CREATED_ON ?? DateTime.Now;
                                    string newPdfForAuthorisedSdg = PDFDir + timeForFolder.ToString(@"yyyy\\MM\\") +
                                                                    outputFileName;
                                    CopyPdfToNewLocation(path, newPdfForAuthorisedSdg);
                                    sdgUser.U_PDF_PATH = newPdfForAuthorisedSdg;
                                    dal.SaveChanges();

                                    CopyToHealthMinistry(path, sdg.SDG_ID, saveAs);

                                    currentReportType = "";
                                    try
                                    {
                                        // find the private folder for customer 
                                        //i don't deal with null exceptions...
                                        string customerGroup =
                                            sdgUser.U_ORDER.U_ORDER_USER.U_CUSTOMER1.U_CUSTOMER_USER.U_CUSTOMER_GROUP;
                                        string privateFolder =
                                            sdgUser.U_ORDER.U_ORDER_USER.U_CUSTOMER1.U_CUSTOMER_USER.U_PRIVATE_LIBRARY;
                                        // ?? PDFDir;
                                        if (customerGroup != null)//Has customer group
                                            try
                                            {
                                                //Get private library by customer GRP                               
                                                privateFolder = GetLibrary4CustomerGRP(customerGroup) ?? privateFolder;
                                            }
                                            catch (Exception ex)
                                            {
                                                Logger.WriteLogFile(ex);
                                                Log(ex);
                                            }

                                        //if no private folder dont copy pdf to it

                                        if (!string.IsNullOrWhiteSpace(privateFolder))
                                        {
                                            privateFolder += privateFolder.EndsWith("\\") ? "" : "\\";
                                            
                                            //Ashi 12/5/21  Send Maccabi XML like Assuta
                                            if (sdgUser.U_ORDER.U_ORDER_USER.U_CUSTOMER1.U_CUSTOMER_USER.U_LETTER_B == "TM")
                                            {
                                                string macFn = sdg.SDG_ID.ToString();
                                                CopyPdfToNewLocation(path, Path.Combine(privateFolder, macFn + ".PDF"));

                                                assutaResponse.Send(sdg, privateFolder, macFn);
                                            }




                                            else if (sdgUser.U_ORDER.U_ORDER_USER.U_CUSTOMER1.U_CUSTOMER_USER.U_CLALIT == null ||
                                                 sdgUser.U_ORDER.U_ORDER_USER.U_CUSTOMER1.U_CUSTOMER_USER.U_CLALIT != "T")
                                            {
                                                //send normal letter
                                                Log("Copying '" + outputFileName + "' to Private Folder: '" +
                                                                  privateFolder + "'");
                                                CopyPdfToNewLocation(path, privateFolder + outputFileName);
                                                currentReportType = "";
                                            }
                                            else //Is Clalit 
                                            {
                                                Console.Write("Creating XML ");
                                                //send clalit Final Report 
                                                try
                                                {
                                                    string outputXmlName = privateFolder +
                                                                           Regex.Replace(outputFileName.ToString(), ".PDF",
                                                                                         ".XML", RegexOptions.IgnoreCase);
                                                    Console.WriteLine("named : '" + outputXmlName + "'");
                                                    var xmlReport = new FinalResponse(dal, sdgUser.SDG, path,
                                                                                      outputXmlName);
                                                    xmlReport.GenerateFile();
                                                    //:should we delete pdf file after encapsulation, or move to a yearmonth backup folder?
                                                    dal.InsertToSdgLog(sdgId, "PDF.XML", 0, destination.U_WRDESTINATION_ID.ToString());
                                                    sdgUser.U_FAX_EMAIL_SENT_ON = DateTime.Now;
                                                    dal.SaveChanges();
                                                    currentReportType = "";
                                                    Console.WriteLine("XML Created");
                                                }
                                                catch (Exception ex)
                                                {
                                                    Log(ex);

                                                    Console.WriteLine("Error creating XML\r\n" + ex.ToString());
                                                    CopyPdfToNewLocation(path,
                                                                         PDFDir + "Error\\" +
                                                                         sdgUser.SDG.NAME.Replace("\\", "-") + ".XML");
                                                    currentReportType = "";
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Log(ex);

                                        Console.WriteLine("Error in " + currentReportType.ToUpper() + " Command.\r\n" +
                                                          ex.ToString());

                                    }

                                }


                                break;
                            case "CLINIC_PDF": //pdf 
                                outputFileName = saveAs;
                                //in case of an error making the filename, the name will be sdg id.PDF. deal with an errror by sending it to the error dir
                                if (outputFileName.ToString().ToUpper() == sdgId.ToString() + ".PDF")
                                {
                                    CopyPdfToNewLocation(path,
                                                         PDFDir + "Error\\" + sdgUser.SDG.NAME.Replace("\\", "-") + ".PDF");
                                    Console.WriteLine("Error in " + currentReportType.ToUpper() +
                                                      " Command. Moving to Error Folder");
                                    currentReportType = "";
                                }
                                else
                                {
                                    if (sdgUser.IMPLEMENTING_CLINIC == null)
                                    {
                                        currentReportType = "";
                                        break;
                                    }
                                    if (sdgUser.IMPLEMENTING_CLINIC.U_CLINIC_USER == null)
                                    {
                                        currentReportType = "";
                                        break;
                                    }
                                    if (sdgUser.IMPLEMENTING_CLINIC.U_CLINIC_USER.U_PRIVATE_LIBRARY == null)
                                    {
                                        currentReportType = "";
                                        break;
                                    }
                                    string privateFolder = null;
                                    try
                                    {

                                        privateFolder =
                                            sdgUser.IMPLEMENTING_CLINIC.U_CLINIC_USER.U_PRIVATE_LIBRARY;
                                    }
                                    catch (Exception ex)
                                    {

                                        Log(ex);

                                        Console.WriteLine(ex);
                                        currentReportType = "";
                                        break;
                                    }

                                    if (!privateFolder.EndsWith(@"\")) privateFolder += @"\";
                                    DateTime timeForFolder;
                                    //if sdgUser.SDG.AUTHORISED_ON = null use created on or NOW
                                    timeForFolder = sdgAuthorisedOn ?? sdgUser.SDG.CREATED_ON ?? DateTime.Now;
                                    //string newPdfForAuthorisedSdg = privateFolder + timeForFolder.ToString(@"yyyy\\MM\\") +
                                    //                                outputFileName;
                                    string newPdfForAuthorisedSdg = privateFolder + outputFileName;
                                    //do the assuta sending bit in the clinic part
                                    if (sdg.SDG_USER.IMPLEMENTING_CLINIC.U_CLINIC_USER.U_ASSUTA_CLINIC_CODE != null &&
                                        sdg.SDG_USER.IMPLEMENTING_CLINIC.U_CLINIC_USER.U_PRIVATE_LIBRARY != null)
                                    {
                                        // genrate the file name
                                        U_CLINIC_USER clinicUser = sdg.SDG_USER.IMPLEMENTING_CLINIC.U_CLINIC_USER;
                                        string outputPath = clinicUser.U_PRIVATE_LIBRARY;
                                        Console.WriteLine("creating assuta response");
                                        if (!outputPath.EndsWith(@"\")) outputPath += @"\";
                                        //Ashi 3/12/19 Change file name to sdg ID
                                        string filename = sdg.SDG_ID.ToString();// sdg.SDG_USER.U_PATHOLAB_NUMBER.Replace ( '\\', '_' ).Replace ( '/', '_' ).Replace ( '-', '_' );
                                        CopyPdfToNewLocation(path, outputPath + filename + ".PDF");
                                        assutaResponse.Send(sdg, outputPath, filename);
                                        Console.WriteLine("assuta response created:" + outputPath + filename);
                                    }
                                    else
                                    {
                                        CopyPdfToNewLocation(path, newPdfForAuthorisedSdg);
                                    }
                                    currentReportType = "";
                                }


                                break;


                            case "FAX": //Fax

                                //  doc.SendFax(saveAsArray[i]);

                                if (Fax((string)path, saveAs))
                                {
                                    sdgUser.U_FAX_EMAIL_SENT_ON = DateTime.Now;
                                    dal.SaveChanges();
                                    currentReportType = "";
                                }
                                break;

                            case "EMAIL": //Send EMail                                                             

                                //And then check other results
                                string sql = string.Format("select LIMS.Get_Coded_Pdf('{0}') as res from  dual", saveAs);
                                Logger.WriteLogFile(sql);
                                Console.WriteLine(sql);
                                var res = dal.GetDynamicStr(sql);
                                Logger.WriteLogFile("result is - " + res);
                                Console.WriteLine(res);
                                string attacedPath = path;
                                if (res != "No")
                                {
                                    attacedPath = SaveProtectedPdf(path);
                                }


                                SentFromOutlook = false;
                                string oAppName = null; //check if app still active and not closed manually
                                try
                                {
                                    oAppName = outlookApp.Name;
                                }
                                catch (Exception em)
                                {
                                    Log(em);

                                }
                                if (outlookApp == null || oAppName == null)
                                    try
                                    {
                                        System.Diagnostics.Process[] processes =
                                            System.Diagnostics.Process.GetProcessesByName("OUTLOOK");

                                        int collCount = processes.Length;

                                        if (collCount != 0)
                                        {

                                            // Outlook already running, hook into the Outlook instance

                                            outlookApp = Marshal.GetActiveObject("Outlook.Application") as oApplication;
                                            Console.Write("Connected to an existing Outlook instence.");
                                        }
                                        else
                                        {
                                            outlookApp = new oApplication();
                                            Console.Write("Starting a new Outlook instence.");
                                        }

                                    }
                                    catch (Exception em1)
                                    {
                                        Log(em1);

                                        Console.Write("Starting a new Outlook process.");
                                        outlookApp = new oApplication();
                                    }

                                Console.WriteLine("Instence Started.");
                                Microsoft.Office.Interop.Outlook.MailItem mailItem =
                                    //  outlookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                                    outlookApp.CreateItemFromTemplate(Outlook_Template);//Ashi - get template from phrase 18/08/20
                                //@"\\vm-nautilus\nautilus-share\Extensions\Patholab-uni.msg");

                                mailItem.Subject = "דיווח תוצאה פתולאב #" + sdgUser.U_PATHOLAB_NUMBER ?? "";
                                mailItem.To = saveAs.Replace(',', ';');
                                //Ashi 17/3/19 Send protected pdf
                                mailItem.Attachments.Add(attacedPath); //path
                                mailItem.Display(false);
                                bool cancel = false;
                                mailItem.Send();


                                //outlookApp.Session.SendAndReceive(true);
                                //Thread.Sleep(3000);

                                // LinkedResource image = new LinkedResource(@"\\vm-nautilus\nautilus-share\Extensions\image001.jpg");
                                // //image.ContentId="01D1FBC3.88724640"
                                // image.ContentId = Guid.NewGuid().ToString();
                                //  byte[] imaeg = File.ReadAllBytes(@"\\vm-nautilus\nautilus-share\Extensions\patholab.jpg");

                                //  mailItem.HTMLBody =@"<span lang=HE style='font-family:""Times New Roman"",""serif""' dir=RTL>מצורפת בזו תוצאת ה   בדיקה שהתקבלה בחברתנו<o:p></o:p></span></p><p class=MsoNormal dir=RTL><span lang=HE style='font-family:""Times New Roman"",""serif""'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal dir=RTL><b><span lang=HE style='font-size:14.0pt;font-family:""Times New Roman"",""serif"";color:navy'>פתו-לאב דיאגנוסטיקה בע&quot;מ</span></b><b><span lang=HE style='font-size:12.0pt;font-family:""Times New Roman"",""serif"";color:blue'><br>טל':08-9407319/131 נייד: 0545846848<br>פקס:08-9409485 מייל: <a href=""mailto:sysadmin@patho-lab.com"" title=""mailto:sysadmin@patho-lab.com""><span lang=EN-US dir=LTR style='color:blue'>sysadmin@patho-lab.com</span></a></span></b><b><span dir=LTR style='font-size:12.0pt;font-family:""Times New Roman"",""serif"";color:#1F497D'><o:p></o:p></span></b></p><p class=MsoNormal dir=RTL><span lang=HE style='font-family:""Times New Roman"",""serif""'>"+
                                //    @"<o:p>&nbsp;</o:p></span></p><p class=MsoNormal dir=RTL>"+
                                //    @"<span dir=LTR><img border=0 width=128 height=35 id=""תמונה_x0020_3"" src=""data:image/jpeg;base64,"+ System.Convert.ToBase64String(imaeg)+@""" alt=""logo- Patho-Lab (2) (800x217)""></span><span dir=LTR>";
                                // mailItem.Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceHigh;

                                sdgUser.U_FAX_EMAIL_SENT_ON = DateTime.Now;
                                dal.SaveChanges();
                                currentReportType = "";
                                break;

                            case "PRINT": //print
                                currentReportType = "";
                                break;
                                object copies = ExecuteOrGetString(destination.U_COPIES, sdg, sdgStatus);
                                if (saveAs != "")
                                    //print to printer name or use server default
                                    //{
                                    //    int printAttempts = 0;
                                    //    while (saveAsArray[i] != "DEFAULT" && GetActivePrinter(doc) != saveAsArray[i] && printAttempts < maxPrintAttempts)
                                    //    {
                                    //        doc.Application.ActivePrinter = saveAsArray[i];
                                    //        Thread.Sleep(TimeSpan.FromSeconds(5)); // emperical testing found this to be sufficient for our system
                                    //        printAttempts++;
                                    //    }
                                    //}
                                    //doc.PrintOut(ref oTrue, ref oFalse, ref range, ref oMissing, ref oMissing, ref oMissing,
                                    //    ref items, ref copies, ref pages, ref pageType, ref oFalse, ref oTrue,
                                    //    ref oMissing, ref oFalse, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                                    try
                                    {
                                        //System.Diagnostics.Process process = new System.Diagnostics.Process();
                                        //System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                                        //startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;


                                        //startInfo.FileName = _adobeAcrobatLocation;
                                        //startInfo.Arguments = @" /t """ + defaultPdf + @""" """ + saveAsArray[i] + @"""";
                                        //Log("Running " + _adobeAcrobatLocation + @" /t """ + defaultPdf + @""" """ +
                                        //    saveAsArray[i] + @"""");
                                        //process.StartInfo = startInfo;
                                        //process.Start();
                                        //process.WaitForExit();

                                        currentReportType = "";
                                    }


                                    catch (Exception ex)
                                    {
                                        Logger.WriteLogFile(ex);
                                    }

                                break;

                            default:

                                break;
                        }


                        dal.InsertToSdgLog(sdgId, "PDF." + currentReportType, 0, destination.U_WRDESTINATION_ID.ToString());
                        Console.WriteLine("Finished " + currentReportType.ToUpper() + " Command");
                        if (currentReportType != "") isError = true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error " + currentReportType ?? "" + "\r\n" + ex.ToString());
                        //error processing
                        //dal.InsertToSdgLog(sdgId, "PDF." + typeArray[i], 0,"error"+ wrDestinationIdArray[i]);
                        //  File.Delete((string)tempPdf);
                        Log(ex);
                    }
                }
            }

            //   File.Delete((string)tempPdf);

            // Close the Word document, but leave the Word application open.
            // doc has to be cast to type _Document so that it will find the
            // correct Close method.                
            //(_Document)
            if (isError == false)
            {

                Console.WriteLine("All Done with file for sdg ID =" + sdgId.ToString() + " Deleting File " + path +
                                  "\r\n-----------------");
                try
                {
                    File.Delete(path);
                }
                catch (Exception e)
                {
                    Log(e);
                }
            }
            else
            {
                try
                {
                    Console.WriteLine("Error, Moving file to Error Dir :" + Path.GetDirectoryName(path) + @"\Error\");
                    Directory.CreateDirectory(Path.GetDirectoryName(path) + @"\Error\"
                                              );
                    MoveWithReplace(path,
                              Path.GetDirectoryName(path) + @"\Error\" + @"\" + Path.GetFileName(path));
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error on moving:" + path);
                    Console.WriteLine("To:" + Path.GetDirectoryName(path) + @"\Error\" + @"\" +
                             Path.GetFileName(path));
                }
            }

            // word has to be cast to type _Application so that it will find
            // the correct Quit method.

        }


        private static void CopyToHealthMinistry(string path, long sdgId, string saveAs)
        {
            string sql = string.Format("select LIMS.Send_Copy_Health_min ('{0}') as res from  dual", sdgId);
            Console.WriteLine(sql);
            var res = dal.GetDynamicStr(sql);
            Logger.WriteLogFile(sql + "\n" + "result is - " + res);
            Console.WriteLine(res);
            string attacedPath = path;
            if (res == "T")
            {
                CopyPdfToNewLocation(path, Path.Combine(HM_Dir, saveAs));
            }


            sql = string.Format("select lims.is_malignant_HM('{0}') from dual", sdgId);
            res = dal.GetDynamicStr(sql);
            Logger.WriteLogFile(sql + "\n" + "result is - " + res);
            if (!string.IsNullOrEmpty( res))
            {
                CopyPdfToNewLocation(path, Path.Combine(Maligant_Dir, res));
            }
        }

        private static string GetLibrary4CustomerGRP(string customerGroup)
        {
            return dal.FindBy<U_CUSTOMER_USER>(cu => cu.U_CUSTOMER_CODE == customerGroup).FirstOrDefault().U_PRIVATE_LIBRARY;
        }

        private static void CopyPdf2Error(string path, SDG_USER sdgUser, string currentReportType)
        {
            CopyPdfToNewLocation(path,
                                 PDFDir + "Error\\" + sdgUser.SDG.NAME.Replace("\\", "-") + ".PDF");
            Log("Error in " + currentReportType.ToUpper() +
                              " Command. Moving to Error Folder");
        }

        private static bool AddWatermark(string inputFile, string outputFile)
        {
            try
            {


                // create a new instance of GhostscriptProcessor

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;

                // add watermark
                startInfo.FileName = ConfigurationManager.AppSettings["ghostscriptgswin32cFullPath"];
                startInfo.Arguments = @"-sOutputFile=""" + outputFile + "\" "
                                      + ConfigurationManager.AppSettings["ghostscriptWatermarkArguments"] + " "
                                      + " -f \"" + ConfigurationManager.AppSettings["ghostscriptWatermarkFilePath"] + "\" \"" + (string)inputFile + "\"";
                process.StartInfo = startInfo;
                process.Start();
                process.WaitForExit();

                return true;
                // show new pdf
                //Process.Start(outputFile);
            }
            catch (Exception ex)
            {
                Logger.MyLog(ex.ToString());
                Console.WriteLine(ex.ToString());
                return false;
            }
        }
        public static void MoveWithReplace(string sourceFileName, string destFileName)
        {
            try
            {

                //first, delete target file if exists, as File.Move() does not support overwrite
                if (File.Exists(destFileName))
                {
                    File.Delete(destFileName);
                }

                File.Move(sourceFileName, destFileName);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        private static void CopyPdfToNewLocation(object defaultPdf, object outputFileName)
        {
            if (Path.GetFileName(outputFileName.ToString()) == "")
            {
                // // make sure the output file is a file path
                outputFileName += (outputFileName.ToString().EndsWith(@"\") ? "" : @"\");
                outputFileName += Path.GetFileName(defaultPdf.ToString());
            }

            // Save document into PDF Format
            Directory.CreateDirectory(Path.GetDirectoryName(outputFileName.ToString()));
            File.Copy(defaultPdf.ToString(), outputFileName.ToString(), true);
            string S = string.Format("Copied from {0} to {1}", defaultPdf.ToString(), outputFileName.ToString());
            Console.WriteLine(S);
            Logger.WriteLogFile(S);

        }
        private static string SaveProtectedPdf(string sourcePath)
        {
            try
            {

                Console.Write("Creating protected PDF. ");


                PHRASE_HEADER ProtectedPDF = dal.GetPhraseByName("Protected PDF");



                var dic = ProtectedPDF.PhraseEntriesDictonary;

                if (ProtectedPDF == null || dic == null || !dic.ContainsKey("User") || !dic.ContainsKey("Password"))
                {
                    Logger.WriteLogFile("No User Or Password********" + sourcePath);
                    return null;
                }

                string dir = Path.GetDirectoryName(sourcePath);
                string fn = Path.GetFileName(sourcePath);
                Directory.CreateDirectory(Path.Combine(dir, "Protected"));

                string newPath = Path.Combine(dir, "Protected", fn);



                //New name for protected file
                string protectedPath = newPath.Replace(".PDF", "_Protected.PDF");

                //Copy
                File.Copy(sourcePath, protectedPath, true);

                // Open an existing document. Providing an unrequired password is ignored.
                PdfDocument document = PdfReader.Open(protectedPath, "some text");

                PdfSecuritySettings securitySettings = document.SecuritySettings;

                // Setting one of the passwords automatically sets the security level to
                // PdfDocumentSecurityLevel.Encrypted128Bit.
                securitySettings.UserPassword = dic["User"];
                securitySettings.OwnerPassword = dic["Password"];

                securitySettings.PermitAccessibilityExtractContent = false;
                securitySettings.PermitAnnotations = false;
                securitySettings.PermitAssembleDocument = false;
                securitySettings.PermitExtractContent = false;
                securitySettings.PermitFormsFill = true;
                securitySettings.PermitFullQualityPrint = false;
                securitySettings.PermitModifyDocument = true;
                securitySettings.PermitPrint = false;

                // Save the document...
                document.Save(protectedPath);
                //// ...and start a viewer.
                //Process.Start(filenameDest);

                Console.Write("protected PDF saved in . " + protectedPath);

                return protectedPath;
            }
            catch (Exception ex)
            {
                Patholab_Common.Logger.WriteLogFile("Err on Saving Protected Pdf " + ex.Message + " Nautilus - Final Letter");
                return null;

            }

        }
        private static bool Fax(string defaultPdf, string faxnumber)
        {
            if (!UseFax)
            {
                Log("Fax is set to not being used");
                return false;
            }

            //string scanFile = Regex.Replace(defaultPdf, ".PDF", ".TIFF", RegexOptions.IgnoreCase);
            //Ashi - Change tiff to tif
            string scanFile = Regex.Replace(defaultPdf, ".PDF", ".TIF", RegexOptions.IgnoreCase);
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            // makes tiff
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = ConfigurationManager.AppSettings["ghostscriptgswin32cFullPath"];
            startInfo.Arguments = @"-o""" + scanFile + "\" " + ConfigurationManager.AppSettings["ghostscriptTiffArguments"] +
                                 " \"" + defaultPdf + "\"";
            process.StartInfo = startInfo;
            process.Start();
            process.WaitForExit();

            //use a global fax server
            try
            {
                //FaxServer faxServer = new FaxServer();
                FaxDoc faxDoc = null;

                faxDoc = faxServer.CreateDocument(scanFile);


                faxDoc.FaxNumber = faxnumber.Replace("-", "");
                faxDoc.RecipientName = faxnumber;
                faxDoc.DisplayName = faxnumber;

                //faxDoc.CoverpageName ="COVER";
                //faxDoc.SendCoverpage = 1;
                ////faxDoc.ServerCoverpage = 2;
                //faxDoc.CoverpageSubject = "Container:" + Path.GetFileNameWithoutExtension(scanFile).Replace("_", "/");

                //faxDoc.CoverpageNote = (container.U_RECEIVED_ON ?? DateTime.Now).ToString(@"dd\/MM\/yyyy");
                faxDoc.Send();
                Log("scanFile:'" + scanFile + @"->" + faxnumber + "' was sent to fax server.");

                try
                {
                    File.Delete(scanFile);
                }
                catch (Exception ex)
                {
                    Log(ex);
                }
                // container.U_FAX_SEND_ON = dal.GetSysdate();
                // dal.SaveChanges();
                return true;


                //var jobId = faxDoc.ConnectedSubmit(faxServer);
            }
            catch (System.Exception e)
            {

                Log(e);
                Log("Error faxing scanFile:'" + scanFile + @"->" + faxnumber + "' !");
                try
                {
                    File.Delete(scanFile);
                }
                catch (Exception ex)
                {
                    Log(ex);
                }
                return false;

            }

        }

        private static void Log(Exception ex)
        {
            Console.WriteLine(ex.ToString());
            Logger.WriteLogFile(ex);
        }
        private static void Log(string ex)
        {
            Console.WriteLine(ex);
            Logger.WriteLogFile(ex);
        }
        private static string GetActivePrinter(Document doc)
        {
            string activePrinter = doc.Application.ActivePrinter;
            int onIndex = activePrinter.LastIndexOf(" on ");
            if (onIndex >= 0)
            {
                activePrinter = activePrinter.Substring(0, onIndex);
            }
            return activePrinter;

        }

        public static bool SentFromOutlook;


        static void MailService_Send(ref bool Cancel)
        {
            SentFromOutlook = true;
        }



        static void ThisAddIn_Close(ref bool Cancel)
        {

        }

        public static bool KillProcess(string name)
        {
            //here we're going to get a list of all running processes on
            //the computer
            bool processsFound = false;
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (Process.GetCurrentProcess().Id == clsProcess.Id)
                    continue;
                //now we're going to see if any of the running processes
                //match the currently running processes. Be sure to not
                //add the .exe to the name you provide, i.e: NOTEPAD,
                //not NOTEPAD.EXE or false is always returned even if
                //notepad is running.
                //Remember, if you have the process running more than once, 
                //say IE open 4 times the loop thr way it is now will close all 4,
                //if you want it to just close the first one it finds
                //then add a return; after the Kill
                if (clsProcess.ProcessName.Contains(name))
                {
                    clsProcess.Kill();
                    processsFound = true;
                }
            }
            //otherwise we return a false
            return processsFound;
        }
        private static string ExecuteOrGetString(string queryOrString, SDG sdg, string sdgStatus)
        {
            string result;
            result = "";

            if (queryOrString == null)
            {
                //default to don`t send/ 1 copy
                result = "";
            }
            //if ther is no select, return the string
            string query = Regex.Replace(queryOrString ?? "", "#SDG_ID#", sdg.SDG_ID.ToString(), RegexOptions.IgnoreCase);
            query = Regex.Replace(query, "#SDG_STATUS#", "'" + sdgStatus.ToUpper() + "'", RegexOptions.IgnoreCase);



            if (query.IndexOf("select", StringComparison.OrdinalIgnoreCase) < 0)
            {
                result = query;
            }
            else
            {

                OracleDataReader reader = RunQuery(query);
                // Run query in queryString, 
                if (reader == null || !reader.HasRows)
                {
                    //if no resulst
                    //return;
                    result = "";
                }
                else
                {
                    result = reader.GetValue(0).ToString();

                }
                if (reader != null) reader.Dispose();
            }

            return result;
        }
        private static OracleDataReader RunQuery(string queryString)
        {
            _cmd = new OracleCommand(queryString, _connection);
            OracleDataReader reader;
            try
            {
                reader = _cmd.ExecuteReader();
                reader.Read();
            }
            catch (Exception ex)
            {
                Log(ex);
                Log("reader is null " + queryString);
                reader = null;
                //continue loop 
            }

            return reader;
        }


    }
}
