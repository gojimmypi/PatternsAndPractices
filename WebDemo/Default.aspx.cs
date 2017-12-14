using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using Microsoft.ApplicationBlocks.Data;
namespace WebDemo
{
    public partial class _Default : Page
    {
        const string MY_SERVER = "hadev12";
        const string MY_DATABASE = "master";

        //***********************************************************************************************************************************
        static string ConnectionString(string strServerName, string strDatabaseName)
        //***********************************************************************************************************************************
        {
            // see http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpref/html/frlrfSystemDataSqlClientSqlConnectionClassConnectionStringTopic.asp
            //
            // there is some debate as to whether the Oledb provider is indeed faster than the native client!
            //  
            return "Workstation ID=myDemo;" +
                   "packet size=8192;" +
                   "Persist Security Info=false;" +
                   "Server=" + strServerName + ";" +
                   "Database=" + strDatabaseName + ";" +
                   "Trusted_Connection=true; " +
                   "Network Library=dbmssocn;" +
                   "Pooling=True; " +
                   "Enlist=True; " +
                   "Connection Lifetime=14400; " +
                   "Max Pool Size=20; Min Pool Size=0";
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            string strSQL = "select * from sys.databases";
            this.DatabaseList.DataSource = SqlHelper.ExecuteDataset(ConnectionString(MY_SERVER, MY_DATABASE), CommandType.Text, strSQL );
            this.DatabaseList.DataBind();
        }
    }
}