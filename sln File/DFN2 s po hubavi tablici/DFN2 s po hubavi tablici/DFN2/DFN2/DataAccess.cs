using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.InteropServices;

namespace DFN2.Database_Elements
{
    public class DataAccess
    {
        public SqlConnection connection = new System.Data.SqlClient.SqlConnection(Helper.CnnVaL("SqlConnection"));
        public void InsertInfoSchools(string name, string neispuo, string street, string city, string community, string area, string post_code, string phone, string email, string principal_name)
        {
            using (connection)
            {
                connection.Execute($"INSERT INTO schools(name,neispuo,street,city,community,area,post_code,phone,email,principal_name) VALUES('{name}','{neispuo}','{street}','{city}','{community}','{area}','{post_code}','{phone}','{email}','{principal_name}')");
            }
        }
        public void UpdateInfoSchools(string name, string neispuo, string street, string city, string community, string area, string post_code, string phone, string email, string principal_name)
        {
            using (connection)
            {
                connection.Execute($"UPDATE schools SET name='{name}',neispuo='{neispuo}',street='{street}',city='{city}',community='{community}',area='{area}',post_code='{post_code}',phone='{phone}',email='{email}',principal_name='{principal_name}' WHERE id=1");
            }
        }
        public void InsertInfoOrders(string number, DateTime date)
        {
            using (connection)
            {
                connection.Execute($"INSERT INTO orders(number,date) VALUES('{number}',{date})");
            }
        }
        public void InsertInfoDepartments(string name)
        {
            using (connection)
            {
                connection.Execute($"INSERT INTO departments(manager_name) VALUES('{name}')");
            }
        }
        public void UpdateInfoDepartments(string name)
        {
            using (connection)
            {
                connection.Execute($"UPDATE departments SET manager_name='{name}'");
            }
        }
        public void InsertInfoMembers(string first_name, string middle_name, string last_name)
        {
            using (connection)
            {
                connection.Execute($"INSERT INTO department_members(first_name,middle_name,last_name,department_id) VALUES('{first_name}','{middle_name}','{last_name}',{1})");
            }
        }
        public bool CheckIfMemberExists(string first_name, string middle_name, string last_name)
        {
            connection.Open();
            using (connection)
            {
                int check = 0;
                SqlCommand command = new SqlCommand($"SELECT COUNT(id) FROM department_members WHERE first_name='{first_name}' AND middle_name='{middle_name}' AND last_name='{last_name}'", connection);
                check = (int)command.ExecuteScalar();
                if (check == 0)
                {
                    connection.Close();
                    return false;
                }
                else
                {
                    connection.Close();
                    return true;
                }
            }
        }
        public void DeleteInfoMembers(string first_name, string middle_name, string last_name)
        {
            using (connection)
            {

                connection.Execute($"DELETE FROM department_members WHERE first_name='{first_name}' AND middle_name='{middle_name}' AND last_name='{last_name}'");
            }
        }
        public bool CheckIfSchoolInserted()
        {
            connection.Open();
            using (connection)
            {
                int check = 0;
                SqlCommand command = new SqlCommand("SELECT COUNT(id) FROM schools WHERE id=id", connection);
                check = (int)command.ExecuteScalar();
                if (check == 0)
                {
                    connection.Close();
                    return false;
                }
                else
                {
                    connection.Close();
                    return true;
                }
            }
        }
        public bool CheckIfManagerInserted()
        {
            connection.Open();
            using (connection)
            {
                int check = 0;
                SqlCommand command = new SqlCommand("SELECT COUNT(id) FROM departments WHERE id=id", connection);
                check = (int)command.ExecuteScalar();
                if (check == 0)
                {
                    connection.Close();
                    return false;
                }
                else
                {
                    connection.Close();
                    return true;
                }
            }
        }
        public void InsertDocument(string fabric_number, string nom_number, string year, string status)
        {
            SqlCommand command = new SqlCommand($"INSERT INTO \"{nom_number}\"(fabric_number,year,status) VALUES('{fabric_number}','{year}','{status}');", connection);
            connection.Open();
            using (connection)
            {
                command.ExecuteNonQuery();
            }
            connection.Close();
        }
        public void InsertDocument(string fabric_number, string nom_number, string year, string status, string series, string origin)
        {
            SqlCommand command = new SqlCommand($"INSERT INTO \"{nom_number}\"(fabric_number,year,status,series,origin) VALUES('{fabric_number}','{year}','{status}','{series}','{origin}');", connection);
            connection.Open();
            using (connection)
            {
                command.ExecuteNonQuery();
            }
            connection.Close();
        }
        public void InsertDocument(string fabric_number, string nom_number, string year, string status,string origin)
        {
            SqlCommand command = new SqlCommand($"INSERT INTO \"{nom_number}\"(fabric_number,year,status,origin) VALUES('{fabric_number}','{year}','{status}','{origin}');", connection);
            connection.Open();
            using (connection)
            {
                command.ExecuteNonQuery();
            }
            connection.Close();
        }
        public List<string> GetAllDocsForOrigin(string nom_number, string origin)
        {
            List<string> result = new List<string>();
            connection.Open();
            using (connection)
            {
                SqlCommand command = new SqlCommand($"SELECT fabric_number FROM \"{nom_number}\" WHERE origin='{origin}'", connection);
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    result.Add((string)reader["fabric_number"]);
                }
            }
            connection.Close();
            return result;
        }
        public int GetCountOfDoc(string nom_number,string status)
        {
            int result = 0;
            connection.Open();
            using (connection)
            {
                SqlCommand command = new SqlCommand($"SELECT COUNT(status) WHERE status={status}",connection);
                result = (int)command.ExecuteScalar();
            }
            connection.Close();
            return result;
        }
        public void UpdateDoc(string nom_number,string fabric_number)
        {
            connection.Open();
            using (connection)
            {
                SqlCommand command = new SqlCommand($"UPDATE \"{nom_number}\" SET status='Предаден на друга институция' WHERE fabric_number={fabric_number}",connection);
                command.ExecuteNonQuery();
            }
            connection.Close();
        }
        public List<string> GetAllDocsForStatus(string nom_number, string status)
        {
            List<string> result = new List<string>();
            connection.Open();
            using (connection)
            {
                SqlCommand command = new SqlCommand($"SELECT fabric_number FROM \"{nom_number}\" WHERE status='{status}'", connection);
                SqlDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    result.Add((string)reader["fabric_number"]);
                }
            }
            connection.Close();
            return result;
        }
        public void InsertOrder(string number, DateTime year)
        {
            using (connection)
            {
                connection.Execute($"INSERT INTO orders(number,date) VALUES('{number}','{year.ToString()}')");
            }
        }
        
        public List<string> GetSchoolData()
        {
            List<string> result = new List<string>();
            connection.Open();
            using (connection)
            {
                SqlCommand sqlCommand = new SqlCommand("SELECT * FROM schools WHERE id=1");

            }
            return result;
        }
        public int GetCountOfDestroyedDoc(string nom_number)
        {
            int result = 0;
            connection.Open();
            using(connection)
            {
                SqlCommand sqlCommand = new SqlCommand($"SELECT COUNT(fabric_number) FROM \"{nom_number}\" WHERE status='Годен за унищожаване' OR status='Анулиран'",connection);
                result = (int)sqlCommand.ExecuteScalar();
            }
            connection.Close();
            return result;
        }
        public int GetCountDupsForYear(string nom_number,string year)
        {
            int result = 0;
            connection.Open();
            using (connection)
            {
                SqlCommand sqlCommand = new SqlCommand($"SELECT COUNT(fabric_number) FROM \"{nom_number}\" WHERE status='Наличен' AND year='{year}'", connection);
                result = (int)sqlCommand.ExecuteScalar();
            }
            connection.Close();
            return result;
        }
        public string GetSchoolName()
        {
            string result ;
            connection.Open();
            using (connection)
            {
                SqlCommand sqlCommand = new SqlCommand($"SELECT name FROM schools WHERE id=1", connection);
                result = (string)sqlCommand.ExecuteScalar();
            }
            connection.Close();
            return result;
        }
        public string GetSchoolNeispuo()
        {
            string result;
            connection.Open();
            using (connection)
            {
                SqlCommand sqlCommand = new SqlCommand($"SELECT neispuo FROM schools WHERE id=1", connection);
                result = (string)sqlCommand.ExecuteScalar();
            }
            connection.Close();
            return result;
        }
        public string GetSchoolStreet()
        {
            string result;
            connection.Open();
            using (connection)
            {
                SqlCommand sqlCommand = new SqlCommand($"SELECT street FROM schools WHERE id=1", connection);
                result = (string)sqlCommand.ExecuteScalar();
            }
            connection.Close();
            return result;
        }
        public string GetSchoolCity()
        {
            string result;
            connection.Open();
            using (connection)
            {
                SqlCommand sqlCommand = new SqlCommand($"SELECT city FROM schools WHERE id=1", connection);
                result = (string)sqlCommand.ExecuteScalar();
            }
            connection.Close();
            return result;
        }
        public string GetSchoolCommunity()
        {
            string result;
            connection.Open();
            using (connection)
            {
                SqlCommand sqlCommand = new SqlCommand($"SELECT community FROM schools WHERE id=1", connection);
                result = (string)sqlCommand.ExecuteScalar();
            }
            connection.Close();
            return result;
        }
        public string GetSchoolArea()
        {
            string result;
            connection.Open();
            using (connection)
            {
                SqlCommand sqlCommand = new SqlCommand($"SELECT area FROM schools WHERE id=1", connection);
                result = (string)sqlCommand.ExecuteScalar();
            }
            connection.Close();
            return result;
        }
        public string GetSchoolPostCode()
        {
            string result;
            connection.Open();
            using (connection)
            {
                SqlCommand sqlCommand = new SqlCommand($"SELECT post_code FROM schools WHERE id=1", connection);
                result = (string)sqlCommand.ExecuteScalar();
            }
            connection.Close();
            return result;
        }
        public string GetSchoolPhone()
        {
            string result;
            connection.Open();
            using (connection)
            {
                SqlCommand sqlCommand = new SqlCommand($"SELECT phone FROM schools WHERE id=1", connection);
                result = (string)sqlCommand.ExecuteScalar();
            }
            connection.Close();
            return result;
        }
        public string GetSchoolEmail()
        {
            string result;
            connection.Open();
            using (connection)
            {
                SqlCommand sqlCommand = new SqlCommand($"SELECT email FROM schools WHERE id=1", connection);
                result = (string)sqlCommand.ExecuteScalar();
            }
            connection.Close();
            return result;
        }
        public string GetSchoolPrincipalName()
        {
            string result;
            connection.Open();
            using (connection)
            {
                SqlCommand sqlCommand = new SqlCommand($"SELECT principal_name FROM schools WHERE id=1", connection);
                result = (string)sqlCommand.ExecuteScalar();
            }
            connection.Close();
            return result;
        }
        public string GetManager()
        {
            string result;
            connection.Open();
            using (connection)
            {
                SqlCommand sqlCommand = new SqlCommand($"SELECT manager_name FROM departments WHERE id=1", connection);
                result = (string)sqlCommand.ExecuteScalar();
            }
            connection.Close();
            return result;
        }
    }
}