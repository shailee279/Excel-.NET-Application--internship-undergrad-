
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;

namespace ExcelGeneration
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void btnSubmit_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["testConnectionString"].ConnectionString);
            SqlCommand cmd = new SqlCommand("select * from usertable where username =@username and password=@password and email=@email ", con);
            cmd.Parameters.AddWithValue("@username", TUsername.Text);
            cmd.Parameters.AddWithValue("@password", TPassword.Text);
             cmd.Parameters.AddWithValue("@email", TEmail.Text);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                Response.Redirect("WebForm2.aspx");
            }
            else
            {
                ClientScript.RegisterStartupScript(Page.GetType(), "validation", "<script language='javascript'>alert('Invalid Username, Password or email')</script>");
            }
        }
    }
}

