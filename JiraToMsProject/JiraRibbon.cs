using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace JiraToMsProject
{
    public partial class JiraRibbon
    {
        private void JiraRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button_import_jira_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Test");
        }
    }
}
