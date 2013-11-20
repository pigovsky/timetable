/*
 * Програма для розбору екселівських файлів з розкладами ТНЕУ.
 * 2013 (c) Піговський Ю.Р. 
 * Програма поширюється на основі ліцензії GNU.
 * 
 */

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Threading;
using System.Runtime.InteropServices;
using System.IO;

namespace ParseTimetableFromExcel
{
    
    

    public partial class MainForm : Form
    {
        List<Lesson> lessons = new List<Lesson>();                


        /*
        private void addLessonToGoogleCalendar()
        {
            Event event = new Event()
            {
                Summary = "Appointment",
                Location = "Somewhere",
                Start = new EventDateTime() {
                    DateTime = new DateTime("2011-06-03T10:00:00.000:-07:00")
                    TimeZone = "America/Los_Angeles"
                },
                End = new EventDateTime() {
                    DateTime = new DateTime("2011-06-03T10:25:00.000:-07:00")
                    TimeZone = "America/Los_Angeles"
                },
                Recurrence = new String[] {
                    "RRULE:FREQ=WEEKLY;UNTIL=20110701T100000-07:00"
                },
                Attendees = new List<EventAttendee>()
                {
                    new EventAttendee() { Email: "attendeeEmail" },
                    // ...
                }
            };

            Event recurringEvent = service.Events.Insert(event, "primary").Fetch();

            //Console.WriteLine(recurringEvent.Id);
        }*/


        public MainForm()
        {
            InitializeComponent();
        }

        const int numberOfRecordsPerLesson = 6;

        

        Workbook wb;
        string fileNameForImport;


        private void importFromMsExcel(object sender, EventArgs e)
        {
            
            
            /*for (int j = 1; j <= valueArray.GetLength(1); j++)
                dataGridView1.Columns.Add("c" + j, "v" + j);*/

            
            FileDialog fd = new OpenFileDialog();
            fd.FileName = "*.xls";

            if (fd.ShowDialog() != DialogResult.OK)
                return;

            fileNameForImport = fd.FileName;
            

            Thread t = new Thread(new ThreadStart(importFromMsExcelThreadProc));
            
            progressForm = new ProgressForm();
            progressForm.ShowDialog(this);          
            
            t.Start();
        }

        private ProgressForm progressForm;

        private long[] _cumSumOfIterationsPerSheet;

        private void importFromMsExcelThreadProc()
        {
            log = File.CreateText("UnMergeText.log");
            Microsoft.Office.Interop.Excel.Application app 
                = new Microsoft.Office.Interop.Excel.Application();

            

            wb = app.Workbooks.Open(fileNameForImport, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            workbookSheetRawDataGrid.Rows.Clear();
            workbookSheetRawDataGrid.Columns.Clear();
            workbookSheetRawDataGrid.Rows.Clear();


            int numSheets = wb.Sheets.Count;
            _cumSumOfIterationsPerSheet = new long[numSheets];

            progressForm.totalNumberOfIterations = 0;
            for (int sheetNum = 1; sheetNum <= numSheets; sheetNum++)
            {
                Worksheet sheet = (Worksheet)wb.Sheets[sheetNum];
                excelRange = sheet.UsedRange;
                progressForm.totalNumberOfIterations +=
                    excelRange.Rows.Count * excelRange.Columns.Count;
                _cumSumOfIterationsPerSheet[sheetNum - 1] = 
                    progressForm.totalNumberOfIterations;
            }

            ExcelScanIntenal(wb);

            wb.Close(false, fileNameForImport, null);
            Marshal.ReleaseComObject(wb);

            app.Quit();
            log.Close();

            progressForm.HideProgressForm(setDatasourceForLessons);
        }

        private void setDatasourceForLessons()
        {
            dataGridView2.DataSource = lessons;
        }

        Range excelRange;
        object[,] valueArray;

        private void ExcelScanIntenal(Workbook workBookIn)
        {
            // Висловлюю спеціальну подяку пану Сему Аллену (Sam Allen)
            // за чудову і зрозумілу статтю http://www.dotnetperls.com/excel 
            // про отримання даних з екселівських файлів
            //
            // Get sheet Count and store the number of sheets.
            //
            int numSheets = workBookIn.Sheets.Count;            

            progressForm.currentNumberOfIterationsPass = 0;

            //
            // Iterate through the sheets. They are indexed starting at 1.
            //
            for (int sheetNum = 1; sheetNum <= numSheets; sheetNum++)
            {
                Worksheet sheet = (Worksheet)workBookIn.Sheets[sheetNum];

                //
                // Take the used range of the sheet. Finally, get an object array of all
                // of the cells in the sheet (their values). You can do things with those
                // values. See notes about compatibility.
                //
                excelRange = sheet.UsedRange;
                
                valueArray = (object[,])excelRange.get_Value(
                    XlRangeValueDataType.xlRangeValueDefault);
                                   
                //
                // Do something with the data in the array with a custom method.
                //
                processData();
                
                

                progressForm.currentNumberOfIterationsPass =
                    _cumSumOfIterationsPerSheet[sheetNum-1];

                
                
            }
        }

        

        StreamWriter log;

        private void processData()
        {            
            int groupRowIndex,groupColIndex;
            // Розшукуємо на екселівському листі слово "група"
            // від якого будемо розшукувати всю іншу інформацію.
            IJ ij = searchText("день", valueArray);
            if (ij != null)
            {
                groupRowIndex = ij.i;
                groupColIndex = ij.j+1;
            }
            else
            {
                ij = searchText("група", valueArray);
                if (ij == null)
                    return;
                groupRowIndex = ij.i + 1;
                groupColIndex = ij.j;
            }
            

            int lessonsBeginRowIndex = groupRowIndex + 1;

            for (int j = groupColIndex; j <= valueArray.GetLength(1); j++)
            {
                string groupTitle = "";
                object groupTitleObj = valueArray[groupRowIndex, j];
                if (groupTitleObj == null)
                    continue;
                else
                {
                    groupTitle = groupTitleObj.ToString().Trim();
                    if ("".Equals(groupTitle))
                        continue;
                }

                int day=1; // Week starts from 1 -- Monday                                
                
                int i = lessonsBeginRowIndex;

                while (i <= valueArray.GetLength(0) - (numberOfRecordsPerLesson-1))
                {
                    if (valueArray[i, 1] != null)
                    {
                        string dayString = valueArray[i, 1].ToString().ToLower().Trim();
                        if (string.IsNullOrWhiteSpace(dayString) ||
                            "день".Equals(dayString))
                        {
                            i++;
                            continue;
                        }
                        else
                            day++; // Parse next day of week
                    }
                                                          
                    
                    object time = excelRange[i, 2];                                       
                    
                   
                    foreach (string week in new string[] {"Odd","Even"})
                    {
                        string subject = emptyForNull (GetValueFromMergedCell(i++, j));                        
                        string teacher = emptyForNull( valueArray[i++, j]);
                        string room = emptyForNull( valueArray[i++, j]);

                        if (!string.IsNullOrWhiteSpace(subject))
                        {
                            var timeRoomTuple = getTimeFromRoomRecord(room);
                            if (timeRoomTuple != null)
                            {
                                time = timeRoomTuple.Item1;
                                room = timeRoomTuple.Item2;
                            }

                            // Розділяємо список аудиторій по пробілах, ігноруючи букви
                            var rooms = getRoomsArray(room);
                            var teachers = extractTeacherList(teacher);

                            for (int k = 0; k < rooms.Length; k++)
                            {
                                var lesson = new Lesson()
                                {
                                    day = day,
                                    group = emptyForNull(groupTitle),
                                    room = rooms[k],
                                    subject = subject,
                                    teacher = k>=teachers.Count?"":teachers[k],
                                    time = emptyForNull(time),
                                    week = week
                                };

                                lessons.Add(lesson);
                            }
                        }
                        progressForm.currentNumberOfIterationsPass += numberOfRecordsPerLesson;
                    }                    
                    
                }
            }

            for (int j = 0; j < valueArray.GetLength(1) - workbookSheetRawDataGrid.Columns.Count; j++)
            {
                int colIndex = workbookSheetRawDataGrid.Columns.Count+1;
                AddColToWorkbookSheetRawData("c" + colIndex);
            }

            for (int i = 1; i <= valueArray.GetLength(0); i++)
                AddRowToWorkbookSheetRawData(GetRow(valueArray, i));
        }

        private static string[] getRoomsArray(string text)
        {
            List<string> rooms = new List<string>();            
            
            foreach (Match m in Regex.Matches(text, @"(\d+)"))
            {
                rooms.Add(m.Value);                
            }

            return rooms.ToArray<string>();
        }

        private static Tuple<string,string> getTimeFromRoomRecord(string room)
        {           
            string pat = @"(\d+)\s*[:\.]\s*(\d+)";

            // Instantiate the regular expression object.
            Regex r = new Regex(pat, RegexOptions.IgnoreCase);

            // Match the regular expression pattern against a text string.
            Match m = r.Match(room);

            if (m.Success)
                return new Tuple<string,string>(
                    m.Groups[1] + ":" + m.Groups[2], // time 
                    Regex.Replace(room, pat, ""));   // and the room string without time

            return null;
        }

        void tryUnMergeString(string s)
        {
            try
            {
                var t1 = UnMergerString(s);
                if (t1 != null)
                    log.WriteLine(s + " is recognized as " + t1.Item1 + " and " + t1.Item2);
            }
            catch (Exception e) 
            {
                log.WriteLine(s + " gives exception " + e.Message);
            }
        }

        private static string addSpaceBeforeAllUppercaseLetters(string source)
        {
            string result = "";


            // Ставимо пробіли перед кожною великою літерою,
            // після якої слідує небуквенний символ.
            // Це треба робити, щоб обробити ситуацію, коли ініціали 
            // пишуть одразу після прізвища без пробіла перед ними,
            // наприклад, "доц.МарценюкЄ.О." замість "доц.Марценюк Є.О."
            for (int i = 0; i < source.Length; i++)
            {
                if (Char.IsUpper(source[i]) && 
                    (i==source.Length-1 || !Char.IsLetter(source[i+1])) )
                    result += " ";
                result += source[i];
            }            

            return result;
        }

        static Tuple<string, string> UnmergeTeacherString(string teacherString)
        {
            string first="", 
                   second="";

            // Замінити в стрічці лапки на апострофи та 
            // розбити її на масив непорожніх стрічок за
            // пробілами та розділовими знаками
            var teacherStringTokens = Regex.Split(
                addSpaceBeforeAllUppercaseLetters(teacherString.Replace('"','\'')),
                @"[\s\.\,]").
                Where(s => s != String.Empty).ToArray<string>();            

            int tokenIndex = 0;
            string token;

            for (; tokenIndex < teacherStringTokens.Length; tokenIndex++)
            {
                token = teacherStringTokens[tokenIndex];

                // Якщо токен починається з малої літери і 
                // має довжину більшу за одну літеру,
                // то це викл., доц., проф. чи ст. викл.
                if (Char.IsLower(token[0]) && token.Length > 1)
                {
                    first += token + ". ";
                }
                else
                    break;
            }

            bool surnameFound = false;

            for (; tokenIndex < teacherStringTokens.Length; tokenIndex++ )
            {
                token = teacherStringTokens[tokenIndex];
                              
                // Якщо токен складається з більше ніж двох літер,
                // то це прізвище або знову викл., доц. чи проф.
                if (token.Length > 2)
                {
                    if (surnameFound)
                        break;
                    first += Char.ToUpper(token[0])+token.Substring(1)+" ";
                    surnameFound = true;
                }
                else
                {
                    first+=token.ToUpper()+".";
                }
            }

            for (; tokenIndex < teacherStringTokens.Length; tokenIndex++)
            {
                token = teacherStringTokens[tokenIndex];
                second += token +" ";
            }

            return new Tuple<string,string>(first, second);
        }

        private static List<string> extractTeacherList(string teacherString)
        {
            var res = new List<string>();

            var currentString = teacherString;

            while (!string.IsNullOrWhiteSpace(currentString))
            {
                var unmergedTuple = UnmergeTeacherString(currentString);
                res.Add(unmergedTuple.Item1);
                currentString = unmergedTuple.Item2;
            }

            return res;
        }


        /**
         * This function was developed by Julia Gordiyevych (c) 18.10.2013
         * 
         */ 
        static Tuple<String, String> UnMergerString(String entity)
        {
            String first = "";
            String second = "";
            entity = entity.Trim();
            int kk = 0, kvb = 0, kp = 0, z = 0, t = 0, p = 0;
            for (int i = 0; i < entity.Length; i++)
            {
                if (kvb > 3)
                {
                    if (p == 0)
                    {
                        p = 1;
                        for (int j = t; j < i; j++) second += entity[j];
                        first = first.Remove(t);
                    }
                    if (entity[i] == ' ')
                    {
                        if (z == 0) { second += entity[i]; kp++; }
                        z = 0;
                    }
                    else { second += entity[i]; z = 1; }
                }
                else
                {
                    if (Char.IsUpper(entity[i])) { z = 1; kvb++; if (kvb == 3) t = i + 2; }
                    if (entity[i] == '.') { z = 1; kk++; }
                    if (entity[i] == ' ')
                    {
                        if (z == 0) { first += entity[i]; kp++; }
                        z = 0;
                    }
                    else { first += entity[i]; z = 1; }
                }
            }
            first = first.Trim();
            second = second.Trim();
            if (second == "") return null;
            else return new Tuple<string, string>(first, second);
        }


        delegate void AddRowToWorkbookSheetRawDataCallback(object[] val);

        private void AddRowToWorkbookSheetRawData(object[] val)
        {
            // InvokeRequired required compares the thread ID of the 
            // calling thread to the thread ID of the creating thread. 
            // If these threads are different, it returns true. 
            if (this.workbookSheetRawDataGrid.InvokeRequired)
            {
                this.Invoke(new
                    AddRowToWorkbookSheetRawDataCallback(AddRowToWorkbookSheetRawData),
                    new object[] { val });
            }
            else
            {
                this.workbookSheetRawDataGrid.Rows.Add(val);
            }
        }

        delegate void AddColToWorkbookSheetRawDataCallback(string val);

        private void AddColToWorkbookSheetRawData(string val)
        {
            // InvokeRequired required compares the thread ID of the 
            // calling thread to the thread ID of the creating thread. 
            // If these threads are different, it returns true. 
            if (this.workbookSheetRawDataGrid.InvokeRequired)
            {                
                this.Invoke(new
                    AddColToWorkbookSheetRawDataCallback(AddColToWorkbookSheetRawData), 
                    new object[] { val });
            }
            else
            {
                this.workbookSheetRawDataGrid.Columns.Add(val, val);
            }
        }

        private static string emptyForNull(object o)
        {
            if (o == null)
                return "";
            if (o is Range)
                return (o as Range).Text;
            return o.ToString().Trim();
        }

        private string GetValueFromMergedCell(int i, int j)
        {
            object val = valueArray[i, j];
            if (val == null)
            {
                var range = excelRange[i, j] as Range;
                if (!range.MergeCells)
                    return "";
                var mergeArea = range.MergeArea;
                val = ((Range)excelRange[mergeArea.Row, mergeArea.Column]).Text;
                valueArray[i, j] = val;
            }
            val = Regex.Replace(val.ToString().ToLower(), @"\s+", "");
            return val.ToString().Trim();
        }

        public static T[] GetRow<T>(T[,] matrix, int row)
        {
            var columns = matrix.GetLength(1);
            var array = new T[columns];
            for (int i = 0; i <= columns; ++i)
            {
                try
                {
                    array[i-1] = matrix[row, i];
                }
                catch (Exception e) 
                {
                    CentralExceptionProcessor.process(e);
                }
            }
            return array;
        }

        class IJ
        {
            public int i {get;set;}
            public int j { get; set; }           
        }

        private IJ searchText(string text, object[,] valueArray)
        {
            if (valueArray == null)
                return null;
            for (var i=0; i<=valueArray.GetLength(0);++i)
                for(var j=0; j<=valueArray.GetLength(1);++j)
                {
                    try
                    {
                        string s = valueArray[i, j].ToString().ToLower().Trim();
                        if (text.Equals(s))
                        {
                            return new IJ { i=i, j=j };
                        }
                    }
                    catch (Exception e) 
                    {
                        CentralExceptionProcessor.process(e);
                    }
                }

            return null;
        }

        private void exportToMysqlDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Thread t = new Thread(new ThreadStart(exportToMysqlDbProc));

            progressForm = new ProgressForm()
            {
                currentNumberOfIterationsPass = 0,
                totalNumberOfIterations = lessons.Count
            };
            progressForm.ShowDialog(this);

            t.Start();            
        }

        private void exportToMysqlDbProc()
        {
            LessonTable.setUp();
            LessonTable.addLessons(lessons, progressForm);
            progressForm.HideProgressForm(exportFinished);
        }

        private void exportFinished()
        {
            MessageBox.Show(this, "Export is finished successfully");
        }
    }

    class Lesson
    {
        public string group { get; set; }
        public string week { get; set; } // Week can be odd or even
        public int day { get; set; }

        
        public string time
        {
            get 
            {
                return _time;
            }
            set 
            {    // У записі про час замінюємо крапку на двокрапку і забираємо всі пробіли          
                _time =
                    Regex.Replace(value.Replace('.', ':'),@"\s+","");            
            } 
        
        }
        public string teacher { get; set; }

        public string subject { get; set; }

        public string room
        {
            get;
            set;
        }

        private string _time;
    }

    static class LessonTable
    {
        private static MySqlConnection _Connection = null;
        private const string Connect =
            "Data Source=localhost;User Id=root;Password=;charset=cp1251";

        private const string database = "timetable";
        private const string table = "timetable";

        private static MySqlConnection Connection
        {
            get
            {
                if (_Connection == null)
                    _Connection = new MySqlConnection(Connect);
                return _Connection;
            }
        }

        public static void setUp()
        {
            Connection.Open();
            MySqlCommand c;
            try
            {
                Connection.ChangeDatabase(database);
            }
            catch (MySqlException e) 
            {
                CentralExceptionProcessor.process(e);
                c = new MySqlCommand("create database " + database, Connection);
                c.ExecuteNonQuery();
                Connection.ChangeDatabase(database);
            }
            //c = new MySqlCommand("DELETE FROM " + table, Connection);
            c = new MySqlCommand("DROP TABLE IF EXISTS " + table,Connection);
            c.ExecuteNonQuery();
            c = new MySqlCommand("CREATE TABLE  " + table +
                " (id int(10) unsigned NOT NULL AUTO_INCREMENT," +
                "st_group varchar(45) NOT NULL," +
                "week varchar(45) NOT NULL," +
                "day int(1) NOT NULL," +
                "lesson_time varchar(45) NOT NULL," +
                "teacher varchar(45)," +
                
                "subject varchar(45)," +
                
                "room varchar(45)," +
                
                "PRIMARY KEY (id))"+
            " default charset=cp1251", Connection);
            c.ExecuteNonQuery();
            Connection.Close();
        }

        public static void addLessons(List<Lesson> lessons, ProgressForm pf)
        {
            Connection.Open();
            foreach (var l in lessons)
            {
                addLesson(l);
                pf.currentNumberOfIterationsPass++;
            }
            Connection.Close();
            
        }

        public static void addLesson(Lesson lesson)
        {                       
            var c = new MySqlCommand("insert into " + table +
                " (st_group," +
                "week," +
                "day," +
                "lesson_time," +
                "teacher," +
                
                "subject," +
                                
                "room)" +
                " values (@group," +
                "@week," +
                "@day," +
                "@time," +
                "@teacher," +
                
                "@subject," +
                
                
                "@room)", Connection);
            c.Parameters.AddWithValue("@group", lesson.group);
            c.Parameters.AddWithValue("@week", lesson.week);
            c.Parameters.AddWithValue("@day",  lesson.day);
            c.Parameters.AddWithValue("@time",  lesson.time);
            c.Parameters.AddWithValue("@teacher", lesson.teacher);
            
            c.Parameters.AddWithValue("@subject", lesson.subject);
            
            c.Parameters.AddWithValue("@room", lesson.room);
            
            
            c.ExecuteNonQuery();             
        }
    }

}
