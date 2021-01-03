using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

//using Microsoft.Office.Interop.Excel;

namespace PPT2Image
{
    class Program
    {
        //static string imageBase = @"H:\output";
        static string imageBase = "."; 
        static string exeBase = "."; 
        //static DBConnect db; 
        static void Main(string[] args)
        {

            //Console.WriteLine(args[0]); 
        
            string prefix = "";
            string pptfile = "";

            /*  GET PATH OF EXE  */
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            //To get the location the assembly normally resides on disk or the install directory
            //string path = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            //once you have the path you get the directory with:
            var directory = System.IO.Path.GetDirectoryName(path);
            exeBase = directory.ToString();
            /* END PATH */
            imageBase = exeBase;

            /* 
             * For commandline only run: 
             * $> PPT2Image sample.pptx
             * OR
             * 
            */

            if (args.Length != 0)  // while using commandline <this>.exe <filename>.pptx
            {                
                pptfile = exeBase + @"\" + args[0].ToString().Trim();
                prefix = args[0].ToString().Trim();
            }

            else // while running tests
            {
                //String pptfile = @"G:\LS_CAT_Apps_Weekly_7-17-15_Final.pptx";
                //string pptfile = @"H:\output\FULL_Q1FY17_LS-SWIFT_DivisionReview_Draft_Final.pptx";
                pptfile = exeBase + @"\" + @"Addis_Zero_Layer_KLARF_Problem_Statement.pptx";
                
            }


            // remove when standalone application 
            //string[] filePaths = System.IO.Directory.GetFiles(imagebase + @"\", "*.pptx");
            //pptfile = filePaths[0];
            // END 



            //db = new DBConnect();
            //db.Insert("insert into weekly (filename, hashtext, imgthumb, imglarge) values ('LS_WEEKLY5', 'PO1 ET', '/img/weekly7/thumb.png', '/img/weekly7/HD.png') "); 


            
            Console.WriteLine("Exe Base Dir = " + exeBase);
            Console.WriteLine("Image Base Dir = " + imageBase);
            
            ppt2Image(pptfile, prefix);

            //readPPTText(pptfile);

           

        }

        /*
         * http://www.free-power-point-templates.com/articles/c-code-to-convert-powerpoint-to-image/         
         */

        static void ppt2Image(string pptfile, string prefix)
        {
            Console.WriteLine("PPT File Location:" + pptfile);
            Microsoft.Office.Interop.PowerPoint.Application pptApplication = new Microsoft.Office.Interop.PowerPoint.Application();
            Presentation pptPresentation = pptApplication.Presentations
            .Open(pptfile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

            int slide_count = pptPresentation.Slides.Count;
            Console.Write("count=" + slide_count);

            for (int i = 1; i <= slide_count; ++i)
            {
                /* full HD*/                
                pptPresentation.Slides[i].Export(imageBase + @"\"+ prefix + "slide" + i + ".png", "png", 800, 600);
                
                /* Thumbnail*/
                pptPresentation.Slides[i].Export(imageBase + @"\" + prefix + "thumb.slide" + i + ".png", "png", 320, 240);
                
            }

        }

        /* 
         * http://mantascode.com/c-get-text-content-from-microsoft-powerpoint-file/ 
         */
        static void readPPTText(string pptfile)
        {
            Microsoft.Office.Interop.PowerPoint.Application PowerPoint_App = new Microsoft.Office.Interop.PowerPoint.Application();
            Microsoft.Office.Interop.PowerPoint.Presentations multi_presentations = PowerPoint_App.Presentations;
            Microsoft.Office.Interop.PowerPoint.Presentation presentation = multi_presentations.Open(pptfile, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
            /*  MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse 
             required to not open the file in a separate process.
             */
            string presentation_text = "";
            string fulltext = "";
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                //fulltext += "\n\n";
                //fulltext += "Slide:" + (i + 1) + " || ";
                foreach (var item in presentation.Slides[i + 1].Shapes)
                {                    
                    
                    var shape = (Microsoft.Office.Interop.PowerPoint.Shape)item;

                    /* if shape object is a group of shapes
                     http://www.pptfaq.com/FAQ00600_Changing_shapes_within_groups_-without_ungrouping-.htm                     
                     */
                    if (shape.Type == MsoShapeType.msoGroup)
                    {
                        Console.WriteLine(shape.GroupItems.Count);
                        for (int j = 1; j <= shape.GroupItems.Count; ++j)
                        {
                            string temp = getShapeText((Microsoft.Office.Interop.PowerPoint.Shape)shape.GroupItems[j]);                            
                            fulltext += temp;
                            presentation_text += temp;
                        }
                    }
                    /* else if shape object is NOT group of shapes and just a Shape*/
                    else
                    {
                        string temp = getShapeText(shape);
                        fulltext += temp;
                        presentation_text += temp;
                    }                        
                    
                }
                presentation_text= presentation_text.Replace("\"", " ").Replace("\'"," ").Replace("\n"," ").Replace("\r", " ");
                //presentation_text= presentation_text.Replace("\'", " ");
                Console.WriteLine("insert into weekly (filename, hashtext, imgthumb, imglarge) values ('" + pptfile + "', '" + presentation_text + "', '/img/weekly/" + pptfile + "thumb_slide" + (i + 1) + ".png', '/img/weekly/" + pptfile + "slide" + (i + 1) + ".png') ");
                //db.Insert("insert into weekly (filename, hashtext, imgthumb, imglarge) values ('" + pptfile + "', '" + presentation_text + "', '/img/weekly/" + pptfile + "_slide" + (i + 1) + ".png', '/img/weekly/" + pptfile + "_slide" + (i + 1) + ".png') "); 
                //Console.Write("Slide:" + (i + 1)+": " + presentation_text);
                Console.Write("\n ");
                presentation_text = "";


            }
            //System.IO.File.WriteAllText("G:\\text.txt",fulltext);
            System.IO.File.WriteAllText("text.txt", fulltext); 
            PowerPoint_App.Quit();
            Console.WriteLine(presentation_text);
            Console.ReadLine();
        }

        public static string getShapeText(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            string presentation_text = "";
            string textString = "";
            
                    
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                        {

                            var textFrame = shape.TextFrame;
                            var textRange = textFrame.TextRange;
                            var paragraphs = textRange.Paragraphs(-1, -1);
                            foreach (Microsoft.Office.Interop.PowerPoint.TextRange paragraph in paragraphs)
                            {
                                var text = paragraph.Text;
                                text = text.Replace("\r", "");
                                text = text.Replace("\n", " ");
                                presentation_text += text + " ";
                                textString += text + " ";
                            }
                        }
                    }
                    if (shape.HasTable == MsoTriState.msoTrue)
                    {
                        var t = shape.Table;

                        for (int j= 1; j <= t.Rows.Count; ++j  )
                            for (int k = 1; k <= t.Columns.Count; ++k)
                            {
                                //if (shape.HasTextFrame == MsoTriState.msoTrue)
                                //{
                                //    if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                                //    {

                                        var textFrame = t.Cell(j, k).Shape.TextFrame;
                                        var textRange = textFrame.TextRange;

                                        presentation_text += textRange.Text+" ";
                                        textString += textRange.Text+" ";



                                //    }
                                //}
                            }
                    }
                    
            
            if (shape.HasChart == MsoTriState.msoTrue)
            {
                Console.WriteLine("Has Chart: True");
                Microsoft.Office.Interop.PowerPoint.Chart t = shape.Chart;
                
                
                if (t.HasTitle) { Console.WriteLine("Title:" + t.ChartTitle.Text.ToString()); textString += t.ChartTitle.Text.ToString()+" "; }
                

                Microsoft.Office.Interop.PowerPoint.SeriesCollection tmp = (Microsoft.Office.Interop.PowerPoint.SeriesCollection)t.SeriesCollection();
                Console.WriteLine("Series Count:" + tmp.Count);
                

                for (int j = 1; j <= tmp.Count; ++j)
                {
                    Microsoft.Office.Interop.PowerPoint.Series aSeries = (Microsoft.Office.Interop.PowerPoint.Series)tmp.Item(j);

                    foreach (object v in (Array)aSeries.XValues)
                    {
                        if (v != null) { Console.WriteLine(v.ToString()); textString += v.ToString() +" "; }
                    }
                    foreach (object v in (Array)aSeries.Values)
                    {
                        if (v != null) { Console.WriteLine(v.ToString()); textString += v.ToString()+ " "; }
                    }
                }               
            }
            textString = textString.Replace("\r", "");
            textString = textString.Replace("\n", " ");
            return textString;

        }

    }

}