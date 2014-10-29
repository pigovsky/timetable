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
using ParseTimetableFromExcel.DataAccessLayer;

// (c) Yuriy Pigovsky, Yulia Gordiyevych, Nazariy Yuzvin, Yuriy Pidkova

namespace ParseTimetableFromExcel
{
        
    public partial class MainForm : Form
    {
        List<Lesson> lessons = new List<Lesson>();                

        public MainForm()
        {
            InitializeComponent();
        }

        const int numberOfRecordsPerLesson = 6;

        

        Workbook wb;
        string[] fileNamesForImport;


        private void importFromMsExcel(object sender, EventArgs e)
        {
            
            
            /*for (int j = 1; j <= valueArray.GetLength(1); j++)
                dataGridView1.Columns.Add("c" + j, "v" + j);*/


            OpenFileDialog fd = new OpenFileDialog();
            fd.FileName = "*.xls";
            fd.Multiselect = true;

            if (fd.ShowDialog() != DialogResult.OK)
                return;

            fileNamesForImport = fd.FileNames;


            startImport();
        }

        private void startImport()
        {
            Thread t = new Thread(new ThreadStart(importFromMsExcelThreadProc));

            progressForm = new ProgressForm();
            progressForm.Show();

            t.Start();
        }

        private ProgressForm progressForm;

        private long[] _cumSumOfIterationsPerSheet;

        private void importFromMsExcelThreadProc()
        {            
            //log = File.CreateText("UnMergeText.log");
            Microsoft.Office.Interop.Excel.Application app
                = new Microsoft.Office.Interop.Excel.Application();

            workbookSheetRawDataGrid.Rows.Clear();
            workbookSheetRawDataGrid.Columns.Clear();
            workbookSheetRawDataGrid.Rows.Clear();

            progressForm.totalNumberOfFiles = fileNamesForImport.Length;
            progressForm.CurrentNumberOfFilesPass = 0;
            foreach (var fileNameForImport in fileNamesForImport)
            {
                Faculty = new FileInfo(fileNameForImport).Directory.Name;
                try{
                    wb = app.Workbooks.Open(fileNameForImport, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);

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

                    wb.Close(false, fileNamesForImport, null);
                    Marshal.ReleaseComObject(wb);

                    progressForm.CurrentNumberOfFilesPass++;
                }
                catch (Exception e)
                {
                    CentralExceptionProcessor.process(e);
                }
            }


            app.Quit();
            //log.Close();

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
               
        private void processData(){      
            
            int groupRowIndex,groupColIndex;
            // Розшукуємо на екселівському листі слово "група"
            // від якого будемо розшукувати всю іншу інформацію.

            TimetableProperties timetableProperties = searchText("ПОНЕДІЛОК", valueArray);
            if (timetableProperties != null)
            {
                groupRowIndex = timetableProperties.LessonsStartRow-1;
                groupColIndex = timetableProperties.LessonsStartCol+1;
            }
            else
            {
                timetableProperties = searchText("група", valueArray);
                if (timetableProperties == null)
                    return;
                groupRowIndex = timetableProperties.LessonsStartRow + 1;
                groupColIndex = timetableProperties.LessonsStartCol;
            }
            

            int lessonsBeginRowIndex = groupRowIndex + 1;
            

            for (int j = groupColIndex; j <= valueArray.GetLength(1); j++)
            {
                string groupTitle = "";
                string previousDayString = "";
                string dayString = "";
                object groupTitleObj = valueArray[groupRowIndex, j];
                if (groupTitleObj == null)
                    continue;
                else
                {
                    groupTitle = groupTitleObj.ToString().Trim();
                    if ("".Equals(groupTitle))
                        continue;
                }

                int day=0; // Week days start from 1 -- Monday                                
                
                int i = lessonsBeginRowIndex;
               

                while (i <= valueArray.GetLength(0) - (numberOfRecordsPerLesson-1))
                {
                    previousDayString = dayString;
                    dayString = GetValueFromMergedCell(i, 1).ToLower().Trim();
                    if (string.IsNullOrWhiteSpace(dayString) ||
                            "день".Equals(dayString) ||
                            groupTitle.ToLower().Trim().Equals(dayString))
                    {
                        i++;
                        continue;
                    }
                    else if (dayString != null && !dayString.Equals(previousDayString))
                    {                                                                                          
                            day++; // Parse next day of week
                    }

                    var timeObj = getTimeFromRoomRecord(emptyForNull(excelRange[i, 2]));

                    string time = timeObj != null ? timeObj.Item1 : emptyForNull(excelRange[i, 2]);
                    
                                                           
                    
                   
                    foreach (int week in new int[] {1,2})
                    {
                        // subjects, teachers and rooms can be in merged cells for 
                        // several groups simultaneously.
                        
                        string subject = Regex.Replace(
                                GetValueFromMergedCell(i++, j).ToLower(), 
                                @"\s+", "");
                        string teacher = GetValueFromMergedCell(i++, j);
                        string room    = GetValueFromMergedCell(i++, j);

                        if (!string.IsNullOrWhiteSpace(subject))
                        {
                            var timeRoomTuple = getTimeFromRoomRecord(room);
                            if (timeRoomTuple != null)
                            {
                                time = timeRoomTuple.Item1;
                                room = timeRoomTuple.Item2;
                            }

                            // Розділяємо список аудиторій по пробілах, ігноруючи букви
                            var rooms = getRoomsArray(teacher + room);

                            // If we have no room for a lesson then 
                            // use '?' character as a room title
                            if (rooms == null || rooms.Length < 1)
                                rooms = new string[] { "?" };

                            var teachers = extractTeacherList(teacher+" "+room);
                          
                            for (int k = 0; k < teachers.Count; k++)
                            {
                                var lesson = new Lesson()
                                {
                                    day = day,                                                                        
                                    room = k>=rooms.Length? rooms.Last() : rooms[k],                                                                        
                                    time = time,
                                    week = week
                                };
                                lesson.faculty.Value = Faculty;
                                lesson.group.Value = emptyForNull(groupTitle);
                                lesson.subject.Value = subject;
                                lesson.teacher.Value = teachers[k];

                                lessons.Add(lesson);
                            }
                        }
                        
                    }
                    progressForm.currentNumberOfIterationsPass += numberOfRecordsPerLesson;

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
        ///end processData


        private string removeAllSpaces(string p)
        {
            return Regex.Replace(p, @"\s+", "");
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

        private static List<DateTime> getDatesArray(string text)
        {
            List<DateTime> dates = new List<DateTime>();

            foreach (Match m in Regex.Matches(text, @"(\d+\.\d+\.\d\d+)"))
            {
                dates.Add(DateTime.Parse(m.Value));
            }

            return dates;
        }

        private static Tuple<string,string> getTimeFromRoomRecord(string room)
        {
            string pat = @"(\d+)\s*[:\.]\s*(\d\d)";

            // Instantiate the regular expression object.
            Regex r = new Regex(pat, RegexOptions.IgnoreCase);

            // Match the regular expression pattern against a text string.
            Match m = r.Match(room);

            int hour;

            if (m.Success && Int32.TryParse(m.Groups[1].Value, out hour))
                return new Tuple<string,string>(
                    String.Format("{0:d2}", hour) + ":" + m.Groups[2], // time 
                    Regex.Replace(room, pat, ""));   // and the room string without time

            return null;
        }

        /*void tryUnMergeString(string s)
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
        }*/

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
            // The method's idea is by Yuriy Pidkova
            string first="", 
                   second="";

            // Замінити в стрічці лапки на апострофи та 
            // розбити її на масив непорожніх стрічок за
            // пробілами та розділовими знаками
            var teacherStringTokens = Regex.Split(
                addSpaceBeforeAllUppercaseLetters(teacherString.Replace('"','\'')),
                @"[\s\.\,\d]").
                Where(s => s != String.Empty).ToArray<string>();            

            int tokenIndex = 0;
            string token;

            for (; tokenIndex < teacherStringTokens.Length; tokenIndex++)
            {
                token = teacherStringTokens[tokenIndex];

                string lctoken = token.ToLower();
                if (lctoken.Equals("викл") ||
                    lctoken.Equals("доц") ||
                    lctoken.Equals("проф"))
                    continue;


                // Якщо токен починається з малої літери і 
                // має довжину більшу за одну літеру,
                // то це щось на зразок викл., доц., проф. чи ст. викл.
                if (Char.IsLower(token[0]) && token.Length > 1)
                {
                    // Commenting the following statement out we
                    // get rid of "doc.", "wykl." and "prof."
                    //first += token + ". ";                    
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
                // Якщо токен з однієї літери і слідує після прізвища,
                // то це ініціал, інакше -- катзнащо
                else if (surnameFound && token.Length == 1)
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
                if (!String.IsNullOrWhiteSpace(unmergedTuple.Item1))
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
            
            return val.ToString().Trim();
        }

        public static T[] GetRow<T>(T[,] matrix, int row)
        {
            var columns = matrix.GetLength(1);
            var array = new T[columns];
            for (int i = 1; i <= columns; ++i)
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

        class TimetableProperties
        {
            public int LessonsStartRow {get;set;}
            public int LessonsStartCol { get; set; }            
        }

        private TimetableProperties searchText(string txt, object[,] valueArray)
        {
            string text = txt.ToLower();
            if (valueArray == null)
                return null;
            string faculty = "";
            for (var i=0; i<=valueArray.GetLength(0);++i)
                for(var j=0; j<=valueArray.GetLength(1);++j)
                {
                    try
                    {
                        string s = valueArray[i, j].ToString().ToLower().Trim();
                        if (s.Contains("факультет") || s.Contains("інститут")
                            || s.Contains("програм"))
                            faculty = s;
                        else if (text.Equals(s))
                        {
                            return new TimetableProperties { LessonsStartRow=i, LessonsStartCol=j };
                        }
                    }
                    catch (Exception e) 
                    {
                        //CentralExceptionProcessor.process(e);
                    }
                }

            return null;
        }

        private void exportToMysqlDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            exportToDbProc(new MySQLAdapter());
        }

        private void exportToDbProc(IDbAccess db)            
        {
            progressForm = new ProgressForm()
            {
                currentNumberOfIterationsPass = 0,
                totalNumberOfIterations = lessons.Count
            };
            Thread t = new Thread(() =>
            {
                try
                {
                    LessonTable.addLessons(db,
                        lessons, progressForm);
                    progressForm.HideProgressForm(exportFinished);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
            });
            
            progressForm.Show();

            t.Start();            
        }

        private void exportFinished()
        {
            MessageBox.Show(this, "Export is finished successfully");
        }

        private void importFromTNEUSiteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadTimetableFromTNEU ltft = new LoadTimetableFromTNEU();
            ltft.ShowDialog();
        }

        private void exportToCouchDBToolStripMenuItem_Click(object sender, EventArgs e)
        {
            exportToDbProc(new CouchDBAdapter());
        }

        private void exportToMysqlDatabaseToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            exportToDbProc(new MySQLAdapter());
        }

        private void importFromDirectoryToolStripMenuItem_Click(object sender, EventArgs e)
        {            
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "Open a folder which contains the xls output";                
                dialog.InitialDirectory = ".";
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string folder = new FileInfo(dialog.FileName).DirectoryName;

                    fileNamesForImport = Directory.GetFiles(folder, "*.xls", SearchOption.AllDirectories);
                    startImport();                    
                }
            }
        }

        public string Faculty { get; set; }
    }

    public class StringSet
    {
        private static IDictionary<String, IList<String>> values = new Dictionary<String, IList<String>>();
        private int id;
        private string table;

        public static void ExportToSQL(MySqlConnection connection)
        {
            foreach (string table in values.Keys)
                exportTableToSQL(connection, table);

        }

        private static void exportTableToSQL(MySqlConnection connection, string table)
        {
            MySqlCommand c = new MySqlCommand("DROP TABLE IF EXISTS " + table, connection);
            c.ExecuteNonQuery();
            c = new MySqlCommand("CREATE TABLE  " + table +
                " (id int(10) unsigned NOT NULL," +
                "value varchar(256) NOT NULL," +
                "PRIMARY KEY (id))" +
            " DEFAULT CHARACTER SET utf8 COLLATE utf8_general_ci", connection);
            c.ExecuteNonQuery();
            int id = 0;
            foreach (var item in values[table])
            {
                c = new MySqlCommand("insert into " + table +
                " (id, value) values (@id, @value)", connection);
                c.Parameters.AddWithValue("@id", id);
                c.Parameters.AddWithValue("@value", item);
                c.ExecuteNonQuery();
                id++;
            }
        }

        public String Value
        {
            get
            {                
                return values[table][id];
            }
            set
            {
                if (!values.ContainsKey(table))
                {
                    values.Add(table, new List<String>());
                }
                id = values[table].IndexOf(value);
                if (id < 0)
                {
                    values[table].Add(value);
                    id = values[table].Count - 1;
                }
            }
        }

        public int Id
        {
            get
            {
                return id;
            }
        }

        public static implicit operator String(StringSet value)
        {
            return value.Value;
        }        

        public StringSet(String table)
        {
            this.table = table;
        }
    }

    class Lesson
    {
        public StringSet faculty = new StringSet("faculty");
        public StringSet group = new StringSet("st_group");
        
        public int week { get; set; } // Week can be odd or even
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
                    Regex.Replace(value.Replace('.', ':'), @"\s+", "");
            }

        }
        public StringSet teacher = new StringSet("teacher");

        public StringSet subject = new StringSet("subject");        

        public string room
        {
            get;
            set;
        }

        private string _time;        
    }

    static class LessonTable
    {               
        public static void addLessons(IDbAccess db, List<Lesson> lessons, ProgressForm pf)
        {
            db.SetUp();
            db.Open();
            foreach (var l in lessons)
            {
                db.AddLesson(l);
                pf.currentNumberOfIterationsPass++;
            }
            db.Close();
        }        
    }

}
