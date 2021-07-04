using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using WindowsFormsApp1.Models;

namespace WindowsFormsApp1.Service
{
    public class PeopleService
    {
        private static string LoadConnectString(string id = "sqlDB")
        {
            return ConfigurationManager.ConnectionStrings[id].ConnectionString;
            //return "data source=mytest.db;version=3;";
        }
        public static List<PeopleModel> LoadPeople()
        {
            using (IDbConnection cnn = new SQLiteConnection(LoadConnectString()))
            {
                cnn.Open();
                var output = cnn.Query<PeopleModel>("select * from people", new DynamicParameters());
                return output.ToList();
            }
        }
        public static void SavePerson(PeopleModel people)
        {
            using (IDbConnection cnn = new SQLiteConnection(LoadConnectString()))
            {
                cnn.Execute("insert into people (MName,Age,Sex,CV,Color) values(@MName,@Age,@Sex,@CV,@Color)", people);
            }
        }
        public static void UpdatePerson(PeopleModel people, int id)
        {
            using (IDbConnection cnn = new SQLiteConnection(LoadConnectString()))
            {
                cnn.Execute("update people set MName = @MName,Age = @Age,Sex = @Sex,CV = @CV,Color = @Color where ID=" + id, people);
            }
        }
        public static List<int> GetByID(string name)
        {
            using (IDbConnection cnn = new SQLiteConnection(LoadConnectString()))
            {
                return cnn.Query<int>("select id from people where MName = @MName", new { MName = name }).ToList();
            }
        }
        public static int Delete(int id)
        {
            using (IDbConnection cnn = new SQLiteConnection(LoadConnectString()))
            {
                return cnn.Execute("delete from people where ID=" + id);
            }
        }
    }
}
