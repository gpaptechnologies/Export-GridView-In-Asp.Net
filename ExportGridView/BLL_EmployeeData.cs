using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace ExportGridView
{
    public class BLL_EmployeeData
    {
        public DataTable GetEmployeeData()
        {
            SqlConnection connString = new SqlConnection(ConfigurationManager.ConnectionStrings["myConn"].ToString());

            SqlConnection connection = null;
            SqlCommand command = null;
            SqlDataAdapter sqlda = null;
            DataTable dtEmployees = null;

            using (connection = connString)
            {
                command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;
                command.CommandText = "USP_GETEMPLOYEES";
                sqlda = new SqlDataAdapter(command);
                dtEmployees = new DataTable();
                sqlda.Fill(dtEmployees);
            }
            return dtEmployees;
        }
    }
}