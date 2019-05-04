using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;
using System.Configuration;

public partial class _SessionTransferForm : System.Web.UI.Page
{

    string strConn = Convert.ToString(ConfigurationManager.ConnectionStrings["HP_Connection"]);

    private void Page_Load(object sender, System.EventArgs e)
    {
        string guidSave;
        if (Request.QueryString["dir"] == "2asp")
        {
            //Add the session information to the database, and redirect to the
            //  ASP implementation of SessionTransfer.
            guidSave = AddSessionToDatabase();
            string strRootURL = "";
            if (Convert.ToString(ConfigurationManager.AppSettings["EnableDebug"]) == "true")
                strRootURL = Convert.ToString(ConfigurationManager.AppSettings["LocalHostPort"]);
            Response.Redirect(strRootURL+"/SessionTransfer.asp?dir=2asp&guid=" + guidSave +
                "&url=" + Server.UrlEncode(Request.QueryString["url"]));
        }
        else if (Request.QueryString["out"] == "1")
        {
            Session.Abandon();
            if (string.IsNullOrEmpty(Convert.ToString(Request.QueryString["url"])) == false && string.IsNullOrEmpty(Convert.ToString(Request.QueryString["referer"])) == false)
            {
                Response.Redirect(Request.QueryString["url"] + "&referer=" + Request.QueryString["referer"] );
            }
            else if (string.IsNullOrEmpty(Convert.ToString(Request.QueryString["url"])) == false)
            {
                Response.Redirect(Request.QueryString["url"]);
            }
            else
            {
                Response.Redirect("/default.asp?out=2");
            }
        }
        else
        {
            //Retreive the session information, and redirect to the URL specified
            //  by the querystring.
            GetSessionFromDatabase(Request.QueryString["guid"]);
            ClearSessionFromDatabase(Request.QueryString["guid"]);
            //Response.Redirect(Request.QueryString["url"]);
        }
    }

    //This method adds the session information to the database and returns the GUID
    //  used to identify the data.
    private string AddSessionToDatabase()
    {
        SqlConnection con = new SqlConnection(strConn);
        SqlCommand cmd = new SqlCommand();
        con.Open();
        cmd.Connection = con;
        int i = 0;
        string strSql, guidTemp = GetGuid();

        while (i < Session.Contents.Count)
        {
            strSql = "INSERT INTO ASPSessionState (GUID, SessionKey, SessionValue) " +
                "VALUES ('" + guidTemp + "', '" + Session.Contents.Keys[i].ToString() + "', '" +
                Session.Contents[i].ToString() + "')";
            cmd.CommandText = strSql;
            cmd.ExecuteNonQuery();
            i++;
        }

        con.Close();
        cmd.Dispose();
        con.Dispose();

        return guidTemp;
    }


    //This method retrieves the session information identified by the parameter
    //  guidIn from the database.
    private void GetSessionFromDatabase(string guidIn)
    {
        SqlConnection con = new SqlConnection(strConn);
        SqlCommand cmd = new SqlCommand();
        SqlDataReader dr;
        con.Open();
        cmd.Connection = con;

        string strSql, guidTemp = GetGuid();

        //Get a DataReader that contains all the Session information
        strSql = "SELECT * FROM ASPSessionState WHERE GUID = '" + guidIn + "'";
        cmd.CommandText = strSql;
        dr = cmd.ExecuteReader();

        //Iterate through the results and store them in the session object
        while (dr.Read())
        {
            Session[dr["SessionKey"].ToString()] = dr["SessionValue"].ToString();
        }

        Session.Timeout = 440;

        //Clean up database objects
        dr.Close();
        con.Close();
        cmd.Dispose();
        con.Dispose();
    }


    //This method removes all session information from the database identified by the 
    //  the GUID passed in through the parameter guidIn.
    private void ClearSessionFromDatabase(string guidIn)
    {
        SqlConnection con = new SqlConnection(strConn);
        SqlCommand cmd = new SqlCommand();
        con.Open();
        cmd.Connection = con;
        string strSql;

        strSql = "DELETE FROM ASPSessionState WHERE GUID = '" + guidIn + "'";
        cmd.CommandText = strSql;
        cmd.ExecuteNonQuery();

        con.Close();
        cmd.Dispose();
        con.Dispose();
    }

    //This method returns a new GUID as a string.
    private string GetGuid()
    {
        return System.Guid.NewGuid().ToString();
    }

    #region Web Form Designer generated code
    override protected void OnInit(EventArgs e)
    {
        //This call is required by the ASP.NET Web Form Designer.
        InitializeComponent();
        base.OnInit(e);
    }

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        this.Load += new System.EventHandler(this.Page_Load);

    }
    #endregion

}