using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ParseTimetableFromExcel
{
    public partial class LoadTimetableFromTNEU : Form
    {
        public LoadTimetableFromTNEU()
        {
            InitializeComponent();
        }

        // Thanks for johnnythawte for this function
        // thanks for http://www.forum.mista.ru/topic.php?id=487600 for Encoding.GetEncoding("windows-1251")
        private string GetPageContent(string URI)
        {
            HttpWebRequest myRequest = 
                (HttpWebRequest)WebRequest.Create(URI); 
            myRequest.Method = "GET"; 
            WebResponse myResponse = myRequest.GetResponse(); 
            StreamReader sr = new 
                StreamReader(myResponse.GetResponseStream(), 
                Encoding.GetEncoding("windows-1251")); 
            string result = sr.ReadToEnd(); 
            sr.Close(); myResponse.Close(); 
            //- See more at: http://www.tech-recipes.com/rx/1954/get_web_page_contents_in_code_with_csharp/#sthash.3qqr4cwP.dpuf
            return result;
        }

        private const string zipDir = "zips/";

        private void DownloadZip(string name, string title)
        {
            if (!Directory.Exists(zipDir))
                Directory.CreateDirectory(zipDir);

            WebClient webClient = new WebClient();
            // Thank Some1Pr0 for following headers
            webClient.Headers.Add(HttpRequestHeader.AcceptEncoding, "gzip,deflate,sdch");
            webClient.Headers.Add(HttpRequestHeader.Referer, "http://www.tneu.edu.ua/study/timetable/");
            webClient.DownloadFile("http://www.tneu.edu.ua/engine/download.php?id=" + name, 
                zipDir + title + ".zip");

            /*WebClient webClient = new WebClient();
            //webClient.QueryString.Add("id", name);
            webClient.DownloadFile(
                "http://www.tneu.edu.ua/engine/download.php?id=939", ".htm");*/
            webClient.Dispose();
            
        }

        MatchCollection allTimetables;

        private void findAllZipsWithTimetable(string page)
        {
            string pat = @"download\.php\?id=(\d+)""\s*>([^<]+)<"; //\s*>([^<]+)</a>";

            // Instantiate the regular expression object.
            Regex r = new Regex(pat, RegexOptions.IgnoreCase);

            // Match the regular expression pattern against a text string.
            allTimetables = r.Matches(page);

            foreach (Match m in allTimetables)
            {
                listBoxTimetableFiles.Items.Add(m.Groups[1]+" - "+m.Groups[2]);
            }
        }

        private void buttonLoad_Click(object sender, EventArgs e)
        {
            string page = GetPageContent("http://www.tneu.edu.ua/study/timetable/");
            /*string[] pageLines = page.Split(new string[]{"\n"}, 
                StringSplitOptions.RemoveEmptyEntries);
            foreach (string line in pageLines)
                listBoxTimetableFiles.Items.Add(line);*/
            findAllZipsWithTimetable(page);
        }

        private void buttonDownload_Click(object sender, EventArgs e)
        {
            foreach (int i in listBoxTimetableFiles.SelectedIndices)
            {
                Match m = allTimetables[i];
                DownloadZip(m.Groups[1].ToString(), m.Groups[2].ToString());
            }
        }
    
        
    }
}
