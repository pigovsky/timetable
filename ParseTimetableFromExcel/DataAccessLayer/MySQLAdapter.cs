using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParseTimetableFromExcel.DataAccessLayer
{
    class MySQLAdapter : ParseTimetableFromExcel.DataAccessLayer.IDbAccess
    {        
        private MySqlConnection _Connection = null;
        private const string Connect =
            "Data Source=localhost;User Id=root;Password=;charset=utf8";

        private const string database = "timetable";
        private const string table = "timetable";

        private MySqlConnection Connection
        {
            get
            {
                if (_Connection == null)
                    _Connection = new MySqlConnection(Connect);
                return _Connection;
            }
        }

        public void Import()
        {
            OpenDB();
            StringSet.ImportFromSQL(Connection);
        }

        public void SetUp(bool drop)
        {
            OpenDB();
            MySqlCommand c;
            //c = new MySqlCommand("DELETE FROM " + table, Connection);
            if (drop)
            {
                c = new MySqlCommand("DROP TABLE IF EXISTS " + table, Connection);
                c.ExecuteNonQuery();

                c = new MySqlCommand("CREATE TABLE  " + table +
                    " (id int(10) unsigned NOT NULL AUTO_INCREMENT," +
                    "week varchar(45) NOT NULL," +
                    "day varchar(45) NOT NULL," +
                    "lesson_time time NOT NULL," +
                    "teacher int(10)," +
                     "subject int(10)," +
                      "room varchar(45)," +
                    "st_group int(10) NOT NULL," +
                    "faculty int(10) NOT NULL, " +
                    "PRIMARY KEY (id))" +
                " DEFAULT CHARACTER SET utf8 COLLATE utf8_general_ci", Connection);
                c.ExecuteNonQuery();
            }            
            StringSet.ExportToSQL(Connection);
            Close();
        }

        private void OpenDB()
        {
            Close();
            Open();

            try
            {
                Connection.ChangeDatabase(database);
            }
            catch (MySqlException e)
            {
                CentralExceptionProcessor.process(e);
                MySqlCommand c = new MySqlCommand("create database " + database +
                " CHARACTER SET utf8 COLLATE utf8_general_ci;", Connection);
                c.ExecuteNonQuery();
                Connection.ChangeDatabase(database);
            }
        }

        public void AddLesson(Lesson lesson)
        {
            var c = new MySqlCommand("insert into " + table +
                // " (faculty," +
                " (week," +
                "day," +
                "lesson_time," +
                "teacher," +
                "subject," +
                "room," +
                "st_group,"+
                "faculty" +
                ")" +
                " values (@week," +
                "@day," +
                "@time," +
                "@teacher," +
                "@subject," +
                "@room," +
                "@group," +
                "@faculty"+
                ")", Connection);            
            c.Parameters.AddWithValue("@week", lesson.week);
            c.Parameters.AddWithValue("@day", lesson.day);
            c.Parameters.AddWithValue("@time", lesson.time);
            c.Parameters.AddWithValue("@teacher", lesson.teacher.Id);
            c.Parameters.AddWithValue("@subject", lesson.subject.Id);
            c.Parameters.AddWithValue("@room", lesson.room);
            c.Parameters.AddWithValue("@group", lesson.group.Id);
            c.Parameters.AddWithValue("@faculty", lesson.faculty.Id);             

            c.ExecuteNonQuery();
        }        

        public void Open()
        {
            Connection.Open();
        }

        public void Close()
        {
            Connection.Close();
        }
    }
}
