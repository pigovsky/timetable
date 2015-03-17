using System;
namespace ParseTimetableFromExcel.DataAccessLayer
{
    interface IDbAccess
    {
        void AddLesson(Lesson lesson);
        void Close();
        void Open();
        void SetUp(bool drop);
    }
}
