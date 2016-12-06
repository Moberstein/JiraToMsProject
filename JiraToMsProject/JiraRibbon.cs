using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using JiraToMsProject.Properties;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

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

                var app = new Excel.Application();
                var workbook = app.Workbooks.Open(filename);
                var worksheet = (Excel.Worksheet)workbook.Worksheets.Item[1];

                var jira = worksheet.Range["A2"].Value2;
                if (!jira.Trim().ToLower().Equals("jira"))
                {
                    MessageBox.Show(Resources.NO_JIRA_IMPORT_FILE_FOUND);
                }

                app.Quit();
            }
        }
    }
}
