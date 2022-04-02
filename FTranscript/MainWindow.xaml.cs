using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Media;
using Microsoft.Win32;
using System.Threading.Tasks;
using System.Threading;

namespace FTranscript
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<string> droppedFiles = new List<string>();
        private static int topMargin = 60; //mm
        private int numOftables = 3;
        private static string formattedPDFFilesPath = "";
        private string formattedHtmlFilePath = "";
        private static int totalPages = 0;
        private static int pageNumber = 0;
        private static int footerHeight = 10;
        private static int headerHeight = 10;
        private static int bodyHeight = 10;
        private static int totalBodyLines = 0;
        private static bool isFirstPage = false;
        public int bodyScaler = 2;

        public MainWindow()
        {
            InitializeComponent();
            formattedPDFFilesPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Generated-Transcripts\\";
            formattedHtmlFilePath = Path.GetFullPath("wkhtmltopdf/temp/");

            footerHeight = (int)footerHeightSld.Value;
            headerHeight = (int)headerHeightSld.Value;

        }

        public int TotalBodyLines
        {
            get { return totalBodyLines = ((int)(bodyHeightSld.Value - footerHeight - headerHeight)); }
            //set { totalBodyLines = TotalBodyLines; }
        }


        private void Lab_Drop(object sender, DragEventArgs e)
        {
            try
            {
                TextBlock tblk = e.Source as TextBlock;
                tblk.Text = "";
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var temp = e.Data.GetData(DataFormats.FileDrop, true) as string[];
                    droppedFiles.AddRange(temp);
                }

                if ((droppedFiles.Count <= 0) || (!droppedFiles.Any())) { return; }
                tblk.FontSize = 14;
                tblk.TextAlignment = TextAlignment.Left;

                int i = 0;
                foreach (string s in droppedFiles)
                {
                    var file = s.Split('\\').Last();
                    if (!file.ToLower().Contains(".pdf"))
                    {
                        continue;
                    }
                    i++;
                    tblk.Text += "  " + i + ". " + file + "\n";

                }

                if (string.IsNullOrEmpty(tblk.Text))
                {
                    tblk.Text = "Please drop transcript files here Or Click to select files";
                    tblk.FontSize = 16;
                    tblk.TextAlignment = TextAlignment.Center;
                }
            }
            catch (Exception er) { Console.WriteLine(er.Message); }

        }
        private void DropLabel_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                TextBlock tblk = e.Source as TextBlock;
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "PDf files(*.pdf)|*.pdf";
                openFileDialog.Multiselect = true;
                if (openFileDialog.ShowDialog() == true && openFileDialog.FileNames.Length > 0)
                {
                    tblk.Text = "";
                    droppedFiles.AddRange(openFileDialog.FileNames);
                    tblk.FontSize = 14;
                    tblk.TextAlignment = TextAlignment.Left;

                    int i = 0;
                    foreach (string s in droppedFiles)
                    {
                        var file = s.Split('\\').Last();
                        if (!file.ToLower().Contains(".pdf"))
                        {
                            continue;
                        }
                        i++;
                        tblk.Text += "  " + i + ". " + file + "\n";

                    }


                }

            }
            catch (Exception er) { Console.WriteLine(er.Message); }
        }


        public void ProcessInputTranscripts(List<string> transcripts)
        {

            List<string> textLines = new List<string>();
            var errorMsg = "";
            try
            {
                foreach (var file in transcripts)
                {
                    pageNumber = 0;

                    ProcessStartInfo info = new ProcessStartInfo
                    {
                        FileName = "pdftohtml.exe",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        Arguments = "\"" + file + "\" Temp/Transcript/",
                        CreateNoWindow = true
                    };
                    Process pdfProc = new Process
                    {
                        StartInfo = info
                    };

                    pdfProc.Start();
                    string output = pdfProc.StandardOutput.ReadToEnd();
                    string error = pdfProc.StandardError.ReadToEnd();

                    pdfProc.WaitForExit();

                    Console.WriteLine(output);
                    Console.WriteLine(error);

                    int sentinel = 0, pageEndTracker = 0;
                    bool pageEndFound = false;
                    using (StreamReader pdfTextR = new StreamReader("Temp/Transcript/s.html"))
                    {
                        while (!pdfTextR.EndOfStream)
                        {
                            //Jump over the first 6 lines
                            if (sentinel < 6)
                            {
                                pdfTextR.ReadLine();
                                sentinel++;
                                continue;
                            }
                            sentinel++;
                            //Remove remaing html tags 
                            var temp = pdfTextR.ReadLine().Replace("<br>", " ")
                                .Replace("<b>", "").Replace("<i>", "").Replace("</b>", "")
                                .Replace("</i>", "").Replace("<A name=2></a>", "")
                                .Replace("<A name=3></a>", "").Replace("<hr>", "");
                            Console.WriteLine(temp);
                            //Check if a page end is hit
                            if (temp.Trim().CompareTo("Controller of Examinations") == 0)
                            {
                                //Page end found
                                pageEndFound = true;
                                continue;
                            }
                            //Skip the ending lines of the page
                            if (pageEndFound && pageEndTracker < 24)
                            {
                                pageEndTracker++;
                                continue;
                            }
                            //
                            pageEndTracker = 0;
                            pageEndFound = false;

                            if (temp.Contains("Registrar"))
                            {
                                break;
                            };

                            textLines.Add(temp);
                        }
                    }

                    //Generating the pdf
                    List<string> header = new List<string>();
                    List<List<string>> courses = null;
                    decompose(textLines, out header, out courses);
                    List<List<List<string>>> sem;
                    organizeIntoSem(courses, out sem);
                    //Display
                    PrintToConsole(sem);

                    if (convertHtmlToPDF(sem, header, TotalBodyLines, 0))
                    {
                        //Add the file name to the view 
                        var fileName = header[1].Trim();
                        var item = new ListViewItem { Content = (formattedFiles.Items.Count + 1) + ". " + fileName };
                        item.MouseDoubleClick += openFormattedFile_MouseDoubleClick;
                        item.Foreground = new SolidColorBrush(Colors.Green);
                        formattedFiles.Items.Add(item);
                    }
                    else
                    {
                        //Add the file name to the view 
                        var fileName = header[1].Trim();
                        var item = new ListViewItem { Content = (formattedFiles.Items.Count + 1) + ". " + fileName };
                        item.MouseDoubleClick += openFormattedFile_MouseDoubleClick;
                        item.Foreground = new SolidColorBrush(Colors.Red);
                        formattedFiles.Items.Add(item);
                    }

                    //ManagePDF.CreatePdfFromHtml(sem, header);

                    textLines.Clear();


                    try
                    {
                        //File.Delete("Temp/Transcript/s.html");
                        //File.Delete("Temp/Transcript/.html");
                        //File.Delete("Temp/Transcript/_ind.html");
                    }
                    catch (Exception) { }

                    statuslbl.Text = "Status: Processing file for" + header[1].Trim();

                }
            }


            catch (FileNotFoundException e)
            {
                errorMsg = "File is not found";

                Console.WriteLine(e.Message);
            }
            catch (FileFormatException e)
            {
                errorMsg = "File must be PDF";

                Console.WriteLine(e.Message);

            }
            catch (InsufficientMemoryException e)
            {
                errorMsg = "Your system is out of memory";

                Console.WriteLine(e.Message);

            }
            catch (PathTooLongException e)
            {
                errorMsg = "The specified path to Some files are the file is too long\n Copy the file to My Document";
                Console.WriteLine(e.Message);

            }
            catch (UnauthorizedAccessException e)
            {
                errorMsg = "You don't have the necessary right to run this program\nPlease run as adminstrator";
                Console.WriteLine(e.Message);

            }
            catch (Exception e)
            {
                errorMsg = "Problem occured scanning some files";
                Console.WriteLine("Pdf Scanning failed");
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                // errorMsg = e.Message + e.StackTrace;
            }
            if (string.IsNullOrEmpty(errorMsg))
            {
                MessageBox.Show("Formatting process completed");
                statuslbl.Text = "Status: Formatting process completed";
            }
            else
            {
                MessageBox.Show("Error: " + errorMsg + "\n\nFormatting process completed");
                statuslbl.Text = "Status: Formatting process completed";

            }


        }
        private void openFormattedFile_MouseDoubleClick(object sender, RoutedEventArgs e)
        {
            var file = "";
            try
            {
                ListViewItem item = sender as ListViewItem;
                file = item.Content.ToString().Split('.')[1].Trim().Replace(" ", "_") + ".pdf";
                Process.Start(formattedPDFFilesPath + file);
            }
            catch (Exception)
            {
                MessageBox.Show(file + " not found!");
            }
        }
        private void viewBtn_Click(object sender, RoutedEventArgs e)
        {
            openInExplorer(formattedPDFFilesPath);

        }


        private void DeletePdf_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MessageBoxResult d = MessageBox.Show("Are you sure you want to delete the files", "Delete", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (d == MessageBoxResult.Yes)
                {
                    Array.ForEach(Directory.GetFiles(formattedPDFFilesPath),
                    delegate (string path) { File.Delete(path); });
                    Array.ForEach(Directory.GetFiles(formattedHtmlFilePath),
                    delegate (string path) { File.Delete(path); });

                    MessageBox.Show("Done !");
                }
            }
            catch (Exception)
            {

            }

        }
        private void openInExplorer(string folderPath)
        {
            if (Directory.Exists(folderPath))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    Arguments = folderPath,
                    FileName = "explorer.exe"
                };

                Process.Start(startInfo);
            }
            else
            {
                MessageBox.Show(string.Format("{0} Directory does not exist!", folderPath));
            }
        }

        private void genBtn_Click(object sender, RoutedEventArgs e)
        {
            statuslbl.Text = "Status: Processing files...";

            genBtn.Cursor = Cursors.Wait;

            if (droppedFiles.Count <= 0)
            {
                MessageBox.Show("Please select the PDF files");
                statuslbl.Text = "Status: No operation";

                genBtn.Cursor = Cursors.Arrow;
                return;
            }
            formattedFiles.Items.Clear();
            ProcessInputTranscripts(droppedFiles);
            genBtn.Cursor = Cursors.Arrow;

            //Remove all generated html files
            try
            {
                Directory.Delete(formattedHtmlFilePath, true);
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        private void decompose(List<string> raw, out List<string> header, out List<List<string>> courses)
        {
            header = new List<string>();
            courses = new List<List<string>>();
            var course = new List<string>();
            int index = 0, trackHeader = 0, trackCourse = 0;
            bool routineDelay = false;
            try
            {
                foreach (var str in raw)
                {
                    index++;

                    //Resetting the delayed routine
                    if (routineDelay)
                    {
                        routineDelay = false;
                        trackCourse = 0;
                        continue;
                    }


                    //Console.Write(index + ". ");
                    //fill the header information
                    //Read the first 16 lines into the header
                    if (trackHeader < 16)
                    {
                        header.Add(str);
                        trackHeader++;
                        continue;
                    }

                    //Start reading courses;
                    //NB: Course will start with the semester (int)
                    //18th line starting from 1 will be the starting of the courses
                    //Each course is spread over 8 lines
                    if (trackCourse < 8)
                    {
                        //Add course information to the course object line by line
                        course.Add(str);

                        //Reset trackCourse and start recording new information about 
                        //a course
                        if (trackCourse == 7)
                        {
                            //Handling descrepancies in long course names which 
                            //extends to the next line
                            //NB: index points to semesters
                            var temp = "";
                            try
                            {
                                temp = raw[index];
                            }
                            catch (Exception) { }

                            int rtemp = 0;
                            if (temp.Length > 1 && !int.TryParse(temp, out rtemp))
                            {
                                //The next string is course name continuation
                                //Add it to the course name of the current object
                                course[2] = course[2] + temp;
                            }

                            //Add completed course list to the courses list

                            courses.Add(course);

                            //Clear information in the course list
                            course = new List<string>();

                            //reset course tracker index
                            trackCourse = 0;

                            //Long course names check
                            if (temp.Length > 1 && !int.TryParse(temp, out rtemp))
                            {
                                //Delay the routine by 1 step
                                trackCourse = 7;
                                routineDelay = true;
                            }

                            continue;
                        }

                        trackCourse++;
                    }





                }
            }
            catch (Exception)
            {
                throw new IndexOutOfRangeException("Semester and courses are not initialized");
            }



        }

        private void organizeIntoSem(List<List<string>> courses, out List<List<List<string>>> sem)
        {

            sem = new List<List<List<string>>>();
            var semTemp = new List<List<List<string>>>();
            int j = 0;
            while (j < 12)
            {
                sem.Add(new List<List<string>>());
                j++;
            }
            try
            {
                foreach (var course in courses)
                {
                    switch (course[0].Trim())
                    {
                        case "FS":
                            course[0] = "1";
                            sem[0].Add(course);
                            break;
                        case "1":
                            sem[0].Add(course);
                            break;
                        case "2":
                            sem[1].Add(course);
                            break;
                        case "3":
                            sem[2].Add(course);
                            break;
                        case "4":
                            sem[3].Add(course);
                            break;
                        case "5":
                            sem[4].Add(course);
                            break;
                        case "6":
                            sem[5].Add(course);
                            break;
                        case "7":
                            sem[6].Add(course);
                            break;
                        case "8":
                            sem[7].Add(course);
                            break;
                        case "9":
                            sem[8].Add(course);
                            break;
                        case "10":
                            sem[9].Add(course);
                            break;
                        case "11":
                            sem[10].Add(course);
                            break;
                        case "12":
                            sem[11].Add(course);
                            break;
                    }
                }
                int i = 0;



                for (var x = sem.Count - 1; x >= 0; x--)
                {
                    if (sem[x].Count > 0)
                    {
                        i = x;
                        i++;
                        break;
                    }
                }

                sem.RemoveRange(i, sem.Count - (i + 1));
                var temp = sem[i];
                sem.Remove(temp);


            }
            catch (ArgumentException)
            {
                //throw new Exception("Some empty semesters were skiped");

            }
            catch (InvalidOperationException)
            {
                //throw new Exception("Some empty semesters were skiped");
            }

        }

        private static double calculateGPA(List<List<string>> sem)
        {
            try
            {
                double totalCredits = 0.0;
                double valueByCredit = 0.0;

                foreach (var s in sem)
                {
                    totalCredits += int.Parse(s[4]);
                    valueByCredit += int.Parse(s[4]) * double.Parse(s[6]);
                }
                return Math.Round(valueByCredit / totalCredits, 2);
            }
            catch (DivideByZeroException)
            {
                Console.WriteLine("Zero division occured in the calculateCGP Function");
                return 0.0;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return 0.0;
            }
        }

        private static double calculateCGPA(int sem, List<List<List<string>>> semCourses)
        {
            List<List<string>> allSemCourses = new List<List<string>>();

            for (int i = 0; i < sem; i++)
            {
                allSemCourses.AddRange(semCourses[i]);
            }

            return Math.Round(calculateGPA(allSemCourses), 2);


        }

        /// <summary>
        /// Print the generated pdf content to the console
        /// Arg is array of semesters with their course 
        /// </summary>
        /// <param name="sem"></param>
        private void PrintToConsole(List<List<List<string>>> allCoursesSemester)
        {
            int i = 0;
            foreach (var s in allCoursesSemester)
            {
                i++;
                foreach (var course in s)
                {
                    foreach (var c in course)
                    {
                        Console.Write(c + " ");
                    }
                    Console.Write("\n");
                }
                Console.WriteLine("GPA: {0}", calculateGPA(s));
                Console.WriteLine("CGPA: {0}", calculateCGPA(i, allCoursesSemester));
                Console.WriteLine();
                Console.WriteLine("*******End of semester " + i + "*******");
                Console.WriteLine();
            }

        }

        /// <summary>
        /// Check if a file is still in used
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        /// 
        /// //wait for file writing to finish
        ///       while (IsFileLocked(new FileInfo(Server.MapPath("~/Temp/doc." + type.Split('/')[1])))) ;style='width:100%;margin-top:50px;position:relative'

        private static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;
            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                if (stream != null)
                {
                    stream.Close();
                }

            }
            return false;
        }

        public static string Addheader(List<string> header)
        {
            try
            {
                return string.Format(@"<header class='container'>
    <section class='school'>
            <h4 style='padding-top:60px; padding-bottom:15px; text-align:center;text-decoration:underline'> TRANSCRIPT</h4>
      <div class='dleft'>
       <p class='transfer-credits '>  Index No.: <span>{0}</span>
       <p class='cgpa '>  CGPA: <span> {1} </span></p>       
         <p class='total-credits '>  Total Credits: <span>  {2} </span></p>

       </div>
       <div class='dright'>
          <p class='student-name '>  Student Name: <span>{3}</span></p>
       <p class='course'>  Course: <span>{4} </span></p>
        <p class=''>  Class Obtained: <span>{5}</span></p>
       
       </div>
     </section>
   </header> 

   <div class='container body'>",
    header[0].Split(':')[1],
    header[3].Split(':')[1],
    header[5].Split(':')[1],
    header[1],
    header[2].Split(':')[1],
    header[4].Split(':')[1]
    );
            }
            catch (Exception)
            {
                MessageBox.Show("Pdf scanning error for some files");
                return "";
            }
        }


        public static string AddFooter(int totalpage = 1, int pageNumber = 1, bool main = true)
        {
            if (main)
                return @"<section  class='footer'  >
<style>

#trans tr{
		  background: white;
		}
		#trans table tr td{
		  font-size:11px;
		  padding-left: 3%;
		}
		#trans table tr th{
		 padding-left: 3%;
padding-bottom:20px;
		}
#trans table{
		 width: 48%;
		 float: left;
		}
.footer{
    height:230px;
}
</style>
<br/>
<h5 style='text-align:center; width:100%;'> KEY</h5>
        <p id='trans'>
        <table>
		 <tr><th>Letter</th><th>Grade</th><th>Score</th><th>Interpretation</th></tr>
		  <tr><td>A</td><td>4.00</td><td>90-100</td><td>Excellent</td></tr>
		  <tr><td>A-</td><td>3.75</td><td>85-89</td><td>Very Good</td></tr>
		  <tr><td>B+</td><td>3.50</td><td>80-84</td><td>Good</td></tr>
		  <tr><td>B</td><td>3.25</td><td>75-79</td><td>Above Average</td></tr>
		  <tr><td>B-</td><td>3.00</td><td>70-74</td><td>Average</td></tr>
		  <tr><td>C+</td><td>2.75</td><td>65-69</td><td>Pass</td></tr>
		  <tr><td>C</td><td>2.50</td><td>60-64</td><td>Pass</td></tr>
		  <tr><td>C-</td><td>2.25</td><td>55-59</td><td>Pass</td></tr>
		</table>
		<table style='margin-left:4%'>
		 <tr><th>Letter</th><th>Grade</th><th>Score</th><th>Interpretation</th></tr>
		  <tr><td>D+</td><td>2.00</td><td>50-54</td><td>Pass</td></tr>
		  <tr><td>D</td><td>1.75</td><td>45-49</td><td>Fail</td></tr>
		  <tr><td>D-</td><td>1.50</td><td>40-44</td><td>Fail</td></tr>
		  <tr><td>F</td><td>1.00</td><td>40</td><td>Fail</td></tr>
		  <tr><td>I</td><td>-</td><td>-</td><td>Incomplete</td></tr>
		  <tr><td>Y</td><td>-</td><td>-</td><td>Complete</td></tr>
		  <tr><td>Z</td><td>-</td><td>-</td><td>Disqualified</td></tr>
		  <tr><td>-</td><td>-</td><td>-</td><td>-</td></tr>
		</table>
        </p>
         <br>
         <br>
        <br>
 <div style='background:white; width: 100%;margin-top: 27%;'>
          <div class='left'><span><hr></span>Controller Of Examinations:</div>
          <div class='right'><span><hr></span>Registrar:</div>
        </div>
         <p style='width:100%;'><br>All Nations University College Official Transcript,  &nbsp&nbsp&nbsp&nbsp&nbsp      Date Generated: " +
             DateTime.Now.ToShortDateString() + @"   &nbsp&nbsp&nbsp&nbsp&nbsp  Time Generated: " + DateTime.Now.ToLocalTime().ToShortTimeString() + @"
</p><br/><p style='text-align:right; width:100%'> Page  " + pageNumber + "  of  " + totalpage + "</p><br></section></div>";
            else return @"<section  class='footer'>
         <p >All Nations University College Official Transcript,  &nbsp&nbsp&nbsp&nbsp&nbsp      Date Generated: " +
            DateTime.Now.ToShortDateString() + @"   &nbsp&nbsp&nbsp&nbsp&nbsp  Time Generated: " + DateTime.Now.ToLocalTime().ToShortTimeString() + @"</p>
         <h5>Page  " + pageNumber + "  of  " + totalpage + "</h5></section></div>";
        }

        public static int PrintPage(List<List<List<string>>> semesters, List<string> header, int TotalBodyLines, int start = 0, bool isScan = false)
        {
            if (start == 0) isFirstPage = true;
            if (!isFirstPage) TotalBodyLines += headerHeight / 20;
            List<int> tableLines;
            var tables = BuildTable(semesters, out tableLines/*paddinglines = 3*/);

            var i = start;
            var body = "<html><head><link rel='stylesheet' type='text/css' href='output.css'/></head><body class='container' style='width:100%'>" + Addheader(header);
            bool isOilnGas = false;

            if (tableLines[2] >= 13) isOilnGas = true;
            for (i = start; i < tables.Count; i++)
            {
                if (tableLines[i] < TotalBodyLines)
                {
                    TotalBodyLines -= tableLines[i];
                    if (tableLines[i] <= 5)
                    {
                        //correcting 12semesters
                        TotalBodyLines -= 2;
                    }
                    if (isOilnGas) TotalBodyLines += 3;

                    body += tables[i];

                }
                else
                {

                    break;
                }
            }


            //Move footer to correct position
            while (TotalBodyLines > 0)
            {
                //print empty lines
                body += printEmptyLine();
                TotalBodyLines--;
                if (TotalBodyLines < 5) break;
            }



            //Check if its last page 
            if (i + 1 > tables.Count)
            {

                body += AddFooter(totalPages, pageNumber, true);
            }
            else
            {

                body += AddFooter(totalPages, pageNumber, false);
            }


            //just return if we want to just calculate the page;


            ///Write to pdf file
            var fileName = "";
            try
            {
                fileName = "page" + pageNumber + ".pdf\"";

                var directoryPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Generated-Transcripts\\" + header[1].Trim().Replace(" ", "_") + "\\";

                //Create output folder if it does not exist
                if (!isScan)
                {
                    if (!Directory.Exists(directoryPath)) Directory.CreateDirectory(directoryPath);
                    if (!Directory.Exists(Path.GetFullPath("wkhtmltopdf/temp/" + header[1].Trim().Replace(" ", "_")))) Directory.CreateDirectory("wkhtmltopdf\\temp\\" + header[1].Trim().Replace(" ", "_"));

                    File.WriteAllText(Path.GetFullPath("wkhtmltopdf/temp/" + header[1].Trim().Replace(" ", "_") + "/" + header[1].Trim().Replace(" ", "_") + pageNumber + ".html"), body);

                    while (IsFileLocked(new FileInfo(Path.GetFullPath("wkhtmltopdf/temp/" + header[1].Trim().Replace(" ", "_") + "/" + header[1].Trim().Replace(" ", "_") + pageNumber + ".html")))) ;



                    //Reduce the header offset for subsequent pages other than the first page
                    var hOffset = topMargin;
                    if (!isFirstPage) hOffset = 10;


                    ProcessStartInfo info = new ProcessStartInfo
                    {
                        FileName = "wkhtmltopdf//bin//wkhtmltopdf.exe",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        Arguments = " --disable-external-links --enable-local-file-access --margin-top " + hOffset + "mm --load-error-handling ignore " +
                       "  \"" + Path.GetFullPath("wkhtmltopdf/temp/" + header[1].Trim().Replace(" ", "_")) + "/" + header[1].Trim().Replace(" ", "_") + pageNumber + ".html\"" + "   \"" + directoryPath + fileName.Replace(" ", "_"),
                        WorkingDirectory = "wkhtmltopdf",
                        CreateNoWindow = true
                    };
                    Process pdfProc = new Process
                    {
                        StartInfo = info
                    };

                    pdfProc.Start();

                    string output = pdfProc.StandardOutput.ReadToEnd();
                    string error = pdfProc.StandardError.ReadToEnd();

                    pdfProc.WaitForExit();

                    Console.WriteLine(output);
                    Console.WriteLine(error);

                    if (!error.Contains("Done"))
                    {
                        throw new IOException(fileName.Replace(" ", "_").ToLower() + " is already in used by another application ");
                    }

                }

                if (i + 1 > tables.Count) return -1;

                pageNumber++;
                if (isScan)
                {
                    totalPages = pageNumber;
                };

            }
            catch (IOException e)
            {
                MessageBox.Show(e.Message);
                return -1;
            }
            catch (Exception e)
            {
                MessageBox.Show("Pdf Conversion failed for file: " + fileName);
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                return -1;
            }
            isFirstPage = false;
            return i;
        }

        public static bool convertHtmlToPDF(List<List<List<string>>> semesters, List<string> header, int totalBodyLines, int start = 0)
        {

            if (!Directory.Exists(Path.GetFullPath("wkhtmltopdf/temp/" + header[1].Trim().Replace(" ", "_")))) Directory.CreateDirectory("wkhtmltopdf\\temp\\" + header[1].Trim().Replace(" ", "_"));
            //Create css file
            CreateCSS("wkhtmltopdf\\temp\\" + header[1].Trim().Replace(" ", "_") + "\\");
            var next = start;
            pageNumber = 1;
            //Scan to get number of pages to print;
            while ((next = PrintPage(semesters, header, totalBodyLines, next, true)) >= 0) ;


            next = start;
            pageNumber = 1;
            while ((next = PrintPage(semesters, header, totalBodyLines, next)) >= 0) ;


            //merge the page into one pdf
            var srcDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Generated-Transcripts\\" + header[1].Trim().Replace(" ", "_");
            var destFile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Generated-Transcripts\\" + header[1].Trim().Replace(" ", "_") + ".pdf";

            //Start new thread;
            MergePages(srcDir, destFile);

            return true;
        }

        /// <summary>
        /// Merges the pdf files together into one pdf files
        /// 
        /// </summary>
        /// <param name="dir">path to the pdf files</param>
        public static void MergePages(string src, string dest)
        {
            try
            {
                File.Delete(dest);
            }
            catch (Exception e) { Console.WriteLine(e.Message); }

            try
            {

                var byteList = new List<byte[]>();
                var pdfPages = Directory.GetFiles(src);

                foreach (var page in pdfPages)
                {
                    byteList.Add(System.IO.File.ReadAllBytes(page));
                }

                var pdfBytes = PdfMerger.MergeFiles(byteList);

                //Output to the destination directory
                File.WriteAllBytes(dest, pdfBytes);

                while (File.OpenRead(dest).Length < 2) ;
                //Perform some clean up
                //Delete the pages
                Array.ForEach(Directory.GetDirectories(formattedPDFFilesPath),
                    delegate (string path)
                    {
                        Array.ForEach(Directory.GetFiles(path), delegate (string innerPath)
                        {
                            File.Delete(innerPath);
                        });
                        //remove the directory
                        Directory.Delete(path);


                    });
            }
            catch (Exception)
            {
                MessageBox.Show("Some transient files were not able to be removed");
            }


        }

        //Delete all 
        /// <summary>
        /// 
        /// </summary>
        /// <param name="allCoursesSemester"></param>
        /// <param name="linesPerTable">how many lines made up the table</param>
        /// <param name="paddingLines">to specified offset of each table</param>
        /// <returns></returns>
        public static List<string> BuildTable(List<List<List<string>>> allCoursesSemester, out List<int> linesPerTable, int paddingLines = 0)
        {
            linesPerTable = new List<int>();
            int lines = 0;

            var html = "";
            List<string> tables = new List<string>();
            var tablerow = "<tr>";
            int i = 0;
            try
            {
                foreach (var sem in allCoursesSemester)
                {
                    i++;
                    lines = paddingLines;
                    html = "<section class='cat'><div class='cat-header'><div class=''>Semester: <span>" + (i)
                        + "</span></div></div><table class='container-fluid table-striped table-hover'><tr class='cat-f-row'><th >Subject Code</th><th class='td-course'>Subject Name</th><th>Exams</th><th>Credit</th><th>Grade</th><th>Grade Point</th><th>Result</th></tr>";
                    var temp = 0;
                    foreach (var course in sem)
                    {
                        lines++;//increase line per row;
                        bool skip = true;
                        foreach (var s in course)
                        {
                            temp++;
                            if (skip) { skip = false; continue; }

                            if (temp == 3)
                                tablerow += "<td class='td-course'>" + s + "</td>";
                            else tablerow += "<td>" + s + "</td>";
                        }//Row end
                        html += tablerow + "</tr>";
                        //reset for next row;
                        tablerow = "<tr>";
                        temp = 0;
                    }//All course prented for the semester
                    linesPerTable.Add(lines);
                    //GPA and CGPA 
                    var gpa = calculateGPA(allCoursesSemester[i - 1]);
                    var floorgpa = Math.Floor(gpa);
                    var decimalgpa = gpa - Math.Floor(gpa);

                    var stringGpa = floorgpa + ".";
                    if ((decimalgpa * 10).ToString().Split('.').Length < 2)
                    {
                        if (decimalgpa > 0)
                        {
                            var tmp = (decimalgpa.ToString()).Split('.')[1];
                            stringGpa += tmp + "0";
                        }
                        else
                        {
                            stringGpa += "00";
                        }
                    }
                    else
                    {
                        stringGpa = gpa.ToString();
                    }

                    var cgpa = calculateCGPA(i, allCoursesSemester);
                    var floorcgpa = Math.Floor(cgpa);
                    var decimalcgpa = cgpa - Math.Floor(cgpa);

                    var stringCgpa = floorcgpa + ".";
                    if ((decimalcgpa * 10).ToString().Split('.').Length < 2)
                    {
                        if (decimalcgpa > 0)
                        {
                            var tmp = (decimalcgpa.ToString()).Split('.')[1];
                            stringCgpa += tmp + "0";
                        }
                        else
                        {
                            stringCgpa += "00";
                        }

                    }
                    else
                    {
                        stringCgpa = cgpa.ToString();
                    }

                    html +=
                        $"<tr class='last-row'> <td> </td> <td> </td> <td></td> <td> GPA </td> <td>{stringGpa}</td> <td> CGPA </td> <td>{stringCgpa}</td></tr></table></section >";
                    tables.Add(html);

                }
            }
            catch (IndexOutOfRangeException e)
            {
                //MessageBox.Show("Problem occured during formatting");
                Console.WriteLine(e.Message);
            }

            return tables;
        }

        public static void CreateCSS(string dir)
        {
            var css = @"
body{
color:#000;
width:auto;
font-family: ' Tahoma, Geneva, Verdana, sans - serif';
margin: 0;

        }
        header{
	margin-top: 0px
    }
     header .school p
    {
        font-weight: bold;
        font-size: 16px;
        padding: 0px;
        margin: 0;
        margin-top:5px;
		   
    }
    header .school p span{
	font-weight: normal;
	font-style: italic;
}
.body{
	height: auto;
	padding-bottom: 0px;
	width: 100%;
    padding-top:15px;
    margin-left:5px;
    margin-right:5px;
}
.cat-header{
	background-color: #deeaf5;
	padding-left:0;
	padding-right: 0;
	font-weight: bold;
    font-size:16px;
}
.cat-header span
{
    font-weight: normal;
}
.cat{
    border: solid thin;
border-color: #d2d1d6;
	margin-top:30px;	
	    padding: 0px;
		margin-bottom: 10px;
		width:98%;
}
th{
	padding: 0;
	margin:0;
	font-weight: bold;
	font-size:14px;
	text-align:left;
}

td{
	text-align:left;
	font-size: 13px;
    width:11%;
}
.td-course{
    width: 34%;
}
table{
	width:100%;
	cursor: pointer;
	margin-top: 10px;
	margin-bottom: 10px;
}
.table-striped>tbody>tr:nth-of-type(odd):hover{
 background-color: lavender;
}
.table-striped>tbody>tr:nth-of-type(even):hover{
 background-color: lavender;
}
.footer{
    background: #deeaf5;
	font-weight:bold;
	padding-left:5px;
	padding-right:5px;
	padding-bottom:1px;
	    width: 97%;
}

.footer p
{
    font-weight: bold;
    font-style:italic;
    text-align: center;
    width:80%;
    font-size: 14px;
    padding:0;
    margin:0;
float: left;
  
}


.dright{
 width:40%;
 float:left;
 position: relative;
 margin-left:20%;
font-size:16px;
}

.dleft{
width: 40%;
float:right;
position:relative;
 margin-right:0%;
font-size:16px;
}
.body{
	position:relative;
	float:left;
	
}

.last-row td
{
    font-weight: bold;
    padding-top: 10px;
}
.left{
	width:55%;
	float:left;
	position:relative;
}
.right{
	width:40%;
	float:right;
	position:relative;
}
span hr
{
    margin-top:-1px;
    width:99%;
    border-style:solid;
    border-width:1px;
    border-color:  #3a46a5;
	border-bottom:none;
}

.footer #trans{
	font-weight: normal;
	text-align: left;
	width:99%;
	padding-bottom: 5px;
	margin-top:5px;
}

.footer h5
{
    width:99%;
    text-align: right;
    margin: 0;
    padding:0;
}
.table-striped>tbody>tr:nth-of-type(even)
{
    background - color: #ebfdfb;
}
.table-striped>tbody>tr:nth-of-type(odd)
{
    background - color: #fff;
}
.table-striped>tbody>tr:nth-child(1)
{
    background - color: lavender;
}
.line{
   height:1px;
   width:100%
}

/********4 tables pe page**********/
header{margin-top:0px;}.cat{margin-top: 0px;font-size: 11px;}
.footer {
	margin-top:10px;
}
.footer #trans{
	margin-top:5px;
}
table{padding: 0;margin:0px;}



/****Compact 8 per page***
.cat{width:100%;float:right;margin-bottom:0px;margin-top: 0px;font-size: 5px;height: auto;
}
th,p,td,tr,table,div,header,body,.dleft,.dright{margin:0;padding:0;width:auto;	}
.footer{position:relative;
	float:left;
	margin-top: 5px;
	width:98.5%;
	font-size: 6px;
	
}
.footer p{
	text-align: left;
	font-size: 7px;
	margin-top:5px;
}
.last-row td{
	font-size: 7px;
}
td{
	font-size:8px;
}
header .school p{
	font-size: 10px;font-weight: bold;
}
.line{
   height:18px;
   width:100%
}
th{font-size: 7px}

.dleft p,.dright p{display: inline-block;width: auto;margin:0;padding:0}
.dleft,.dright{width: auto; float:left;margin-bottom: 5px;}
table{
	width: 100%;
}

*/";
            File.WriteAllText(dir + "//output.css", css);
        }

        public static string printEmptyLine()
        {
            return "<p class='line'></p>";
        }

        private void headerHeightSld_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            try
            {
                var src = sender as Slider;
                headerHeight = (int)src.Value;
                headerHeightLbl.Content = headerHeightLbl.Content.ToString().Split(':')[0] + ": " + headerHeight;
            }
            catch (Exception)
            {

            }
        }

        private void footerHeightSld_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            try
            {
                var src = sender as Slider;
                footerHeight = (int)src.Value;
                footerHeightLbl.Content = footerHeightLbl.Content.ToString().Split(':')[0] + ": " + footerHeight;
            }
            catch (Exception)
            {

            }
        }

        private void bodyHeightSld_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            try
            {
                var src = sender as Slider;
                bodyHeight = (int)src.Value;
                bodyHeightLbl.Content = bodyHeightLbl.Content.ToString().Split(':')[0] + ": " + bodyHeight;
            }
            catch (Exception)
            {

            }
        }

        private void resetSld_Click(object sender, RoutedEventArgs e)
        {
            headerHeightSld.Value = 100;
            bodyHeightSld.Value = 150;
            footerHeightSld.Value = 20;
        }

        private void clrViewBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DropLabel.Text = "Please drop transcript files here Or Click to select files";
                droppedFiles.Clear();
            }
            catch (Exception)
            {

            }
        }


    }
}
