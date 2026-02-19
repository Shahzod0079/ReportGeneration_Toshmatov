using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using WpfApp1.Classes.Common;
using System.Text.RegularExpressions;
using ReportGeneration_Toshmatov.Models;

namespace ReportGeneration_Toshmatov.Classes
{
    public class GroupContext : Models.Group
    {
        public GroupContext(int Id, string Name) : base(Id, Name) { }

        public static List<GroupContext> Allgroups()
        {
            List<GroupContext> allgroups = new List<GroupContext>();
            MySqlConnection connection = Connection.OpenConnection();
            MySqlDataReader BDGroups = Connection.Query("SELECT * FROM `Group` ORDER BY `Name`", connection);
            while (BDGroups.Read())
            {
                allgroups.Add(new GroupContext(
                    BDGroups.GetInt32(0),
                    BDGroups.GetString(1)));
            }
            Connection.CloseConnection(connection);
            return allgroups;
        }
    }
}