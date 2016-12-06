using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using JiraToMsProject.Properties;
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
            var openFileDialog = new OpenFileDialog
            {
                RestoreDirectory = true,
                Filter = Resources.OPEN_FILE_FILTER
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var filename = openFileDialog.FileName;
            }
        }
    }
}
