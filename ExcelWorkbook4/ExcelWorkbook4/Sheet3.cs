using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelWorkbook4
{
    public partial class Sheet3
    {
        private void Sheet3_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet3_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.reload.Click += new System.EventHandler(this.reload_Click);
            this.update.Click += new System.EventHandler(this.update_Click);
            this.Startup += new System.EventHandler(this.Sheet3_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet3_Shutdown);

        }

        #endregion

        private void update_Click(object sender, EventArgs e)
        {
            pub.UpdateDirectory(this);
        }

        private void reload_Click(object sender, EventArgs e)
        {
            pub.reloadDirectory(this);
        }
    }
}
