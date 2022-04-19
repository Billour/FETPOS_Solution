using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ToolsLibrary.Logic;
using ToolsLibrary.Entity;
using ToolsLibrary.Helper;
using log4net;

namespace ToolsWeb
{
    public partial class _Default : System.Web.UI.Page
    {
        private ILog log = log4net.LogManager.GetLogger(typeof(_Default));							


        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnCreate_Click(object sender, EventArgs e)
        {
            WordDocHelper docHelper = null;

            try
            {
                log.Info("Get Schema Map");

                OracleSchemaLogic logic = new OracleSchemaLogic();

                SchmaMap map = logic.GetOracleSchema();

                log.Info("Get Save Target");
                //int count = map.Tables.Count;
                string fileName = Server.MapPath(String.Format("~/target/{0}", String.Format("POS_Schema_{0}.doc", DateTime.Now.ToString("yyyy-MM-dd"))));

                docHelper = new WordDocHelper(Server.MapPath("~/doc/TableSchemaTemplate.doc"), Server.MapPath("~/target"), fileName);

                log.Info("Insert Word");

                docHelper.InsertSchema(map);
                
                log.Info("Save");
                
                docHelper.Save();

                log.Info("Save Word Success");
            }
            finally
            {
                log.Info("Word Close");

                docHelper.Close();
            }
           
        }
    }
}
