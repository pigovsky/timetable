using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ParseTimetableFromExcel.DataAccessLayer
{
    class CouchDBAdapter : IDbAccess
    {
        const string HOST = "http://127.0.0.1:5984/";
        const string DB = HOST+"timetable/";

        private static string SendJSON(string resource, string method, string json)
        {
            // Thanks to sarh (http://stackoverflow.com/questions/15091300/posting-json-to-url-via-webclient-in-c-sharp) for following json request
            using (var cli = new WebClient())
            {
                cli.Headers[HttpRequestHeader.ContentType] = "application/json";
                cli.Encoding = Encoding.UTF8;
                var res = cli.UploadString(DB + resource, method, json);
                return res;
            }
        }
        static private string GetUUID()
        {
            using (var cli = new WebClient())
            {
                string json = cli.DownloadString(HOST+"_uuids");
                return (string)JObject.Parse(json)["uuids"][0];
            }
        }
        public void AddLesson(Lesson doc)
        {
            var uuid = GetUUID();            
            string json = JsonConvert.SerializeObject(doc);
            
            SendJSON(uuid, "PUT", json);            
        }



        public void Close()
        {
            // CouchDB is stateless so it does not need a close operation
        }

        public void Open()
        {
            // CouchDB is stateless so it does not need an open operation
        }

        public void SetUp()
        {
            // CouchDB is schema-free so it does not require database
            // create statements 
        }
    }
}
