using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private   const string POSTSCRIPT_APPEND_WATERMARK_TO_EACH_PAGE = @"
          -c  /watermarkText { (CANCELED) } def
            /watermarkFont { /Helvetica-Bold 72 selectfont } def
            /watermarkColor { .75 setgray } def
            /watermarkAngle { 45 } def

            /pageWidth { currentpagedevice /PageSize get 0 get } def
            /pageHeight { currentpagedevice /PageSize get 1 get } def
   
            <<
             /EndPage {
              2 eq { pop false } 
              {
               gsave
               watermarkFont
               watermarkColor
               pageWidth .5 mul pageHeight .5 mul translate
               0 0 moveto
               watermarkText false charpath flattenpath pathbbox
               4 2 roll pop pop
               0 0 moveto
               watermarkAngle rotate
               -.5 mul exch -.5 mul exch
               rmoveto
               watermarkText show
               grestore 
               true 
              } ifelse
             } bind
            >> setpagedevice";
        public MainWindow()
        {
            InitializeComponent();

            string pathToLastPDF = @"c:\work\aa00022-17V1-361.pdf";
            string pathToLastPDFWithWatermark = @"c:\work\aa00022-17V1-361_WM.pdf";
            File.Delete(pathToLastPDFWithWatermark);
            AddWatermark(pathToLastPDF, pathToLastPDFWithWatermark);
        }
        private static bool AddWatermark(string inputFile, string outputFile)
        {


            // get the last installed Ghostscript reference
            //GhostscriptVersionInfo gsversion =
            //    GhostscriptVersionInfo.GetLastInstalledVersion(
            //                GhostscriptLicense.GPL | GhostscriptLicense.AFPL,
            //                GhostscriptLicense.GPL);
            try
            {

                //List<string> switches = new List<string>();
                //switches.Add(string.Empty);

                // set required switches
                //switches.Add("-dBATCH");
                //switches.Add("-dNOPAUSE");
                //switches.Add("-dNOPAUSE");
                //switches.Add("-sDEVICE=pdfwrite");
                //switches.Add("-sOutputFile=" + outputFile);
                //switches.Add("-c");
                //switches.Add(POSTSCRIPT_APPEND_WATERMARK_TO_EACH_PAGE);
                //switches.Add("-f");
                //switches.Add(inputFile);

                // create a new instance of GhostscriptProcessor
                //using (GhostscriptProcessor processor = new GhostscriptProcessor(gsversion, false))
                //{
                //    // start processing pdf file
                //    processor.StartProcessing(switches.ToArray(), null);
                //}

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
                //startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;


                startInfo.FileName = ConfigurationManager.AppSettings["ghostscriptgswin32cFullPath"];
                startInfo.Arguments = @"-sOutputFile=""" + outputFile + "\" "
                                      + ConfigurationManager.AppSettings["ghostscriptArguments"] + " "
                                    + "-sDEVICE=pdfwrite "
                                      + POSTSCRIPT_APPEND_WATERMARK_TO_EACH_PAGE
                                      + " -f\"" + (string)inputFile + "\"";
                process.StartInfo = startInfo;
                process.Start();
               // process.WaitForExit();

                return true;
                // show new pdf
                //Process.Start(outputFile);
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    }
}
