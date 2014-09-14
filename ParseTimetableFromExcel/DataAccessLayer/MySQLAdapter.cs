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

        public void SetUp()
        {
            Close();
            Open();
            MySqlCommand c;
            try
            {
                Connection.ChangeDatabase(database);
            }
            catch (MySqlException e)
            {
                CentralExceptionProcessor.process(e);
                c = new MySqlCommand("create database " + database +
                " CHARACTER SET utf8 COLLATE utf8_general_ci;", Connection);
                c.ExecuteNonQuery();
                Connection.ChangeDatabase(database);
            }
            //c = new MySqlCommand("DELETE FROM " + table, Connection);
            c = new MySqlCommand("DROP TABLE IF EXISTS " + table, Connection);
            c.ExecuteNonQuery();
            c = new MySqlCommand("CREATE TABLE  " + table +
                " (id int(10) unsigned NOT NULL AUTO_INCREMENT," +
                "week varchar(45) NOT NULL," +
                "day varchar(45) NOT NULL," +
                "lesson_time time NOT NULL," +
                "teacher varchar(45)," +
                 "subject varchar(45)," +
                  "room varchar(45)," +
                "st_group varchar(45) NOT NULL," +
                "faculty varchar(45) NOT NULL, " +
                "PRIMARY KEY (id))"+
            " DEFAULT CHARACTER SET utf8 COLLATE utf8_general_ci", Connection);
            c.ExecuteNonQuery();
            Close();
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
            c.Parameters.AddWithValue("@teacher", Win1251ToUTF8(lesson.teacher));
            c.Parameters.AddWithValue("@subject", Win1251ToUTF8(lesson.subject));
            c.Parameters.AddWithValue("@room", lesson.room);
            c.Parameters.AddWithValue("@group", Win1251ToUTF8(lesson.group));
            c.Parameters.AddWithValue("@faculty", Win1251ToUTF8(lesson.faculty));             

            c.ExecuteNonQuery();
        }

        private string Win1251ToUTF8(string source)
        {
            return source;                      
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
