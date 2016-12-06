using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using JiraToMsProject.Properties;
using Microsoft.Office.Tools.Ribbon;

using Excel = Microsoft.Office.Interop.Excel;
using MSProject = Microsoft.Office.Interop.MSProject;

namespace JiraToMsProject
{
    public partial class JiraRibbon
    {
        private string[] ganttPhases =
        {
            "Business_Modeling",
            "Requirements",
            "Analysis_Design",
            "Implementation",
            "Test",
            "Deployment",
            "Change_Management",
            "Project_Management",
            "Environment"
        };

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

                var window = (MSProject.Project)e.Control.Context;
                var project = window.Application.ActiveProject;

                var issues = new List<Issue>();
                var globalSprints = new List<string>();
                var line = 5;
                while (!string.IsNullOrEmpty(worksheet.Range["C" + line].Value2))
                {
                    string sprintString = worksheet.Range["AV" + line].Value2;
                    string[] sprints = { };
                    if (!string.IsNullOrWhiteSpace(sprintString))
                    {
                        sprints = sprintString.Split(',');
                        for (var i = 0; i < sprints.Length; i++)
                        {
                            sprints[i] = sprints[i].Trim();
                            if (!globalSprints.Contains(sprints[i]))
                            {
                                globalSprints.Add(sprints[i]);
                            }
                        }
                    }

                    string labelString = worksheet.Range["AJ" + line].Value2;
                    string[] labels = { };
                    if (!string.IsNullOrWhiteSpace(labelString))
                    {
                        labels = labelString.Split(',');
                        for (var i = 0; i < labels.Length; i++)
                        {
                            labels[i] = labels[i].Trim();
                        }
                    }

                    var item = new Issue
                    {
                        Key = worksheet.Range["B" + line].Value2,
                        Name = worksheet.Range["C" + line].Value2,
                        Type = worksheet.Range["D" + line].Value2,
                        Ressource = worksheet.Range["H" + line].Value2,
                        Created = worksheet.Range["K" + line].Value2?.ToString(),
                        Estimated = worksheet.Range["V" + line].Value2?.ToString(),
                        SubTasks = worksheet.Range["Z" + line].Value2,
                        Labels = labels,
                        Epic = worksheet.Range["AL" + line].Value2,
                        Sprints = sprints
                    };
                    issues.Add(item);

                    line++;
                }

                var outdent = false;

                var firstSprint = true;
                var outdentSprint = false;
                var outdentPhase = false;
                globalSprints.Sort();
                foreach (var globalSprint in globalSprints)
                {
                    var sprint = project.Tasks.Add(globalSprint);
                    if (firstSprint)
                    {
                        firstSprint = false;
                    }
                    else if (outdentSprint)
                    {
                        try
                        {
                            sprint.OutlineOutdent();
                        }
                        catch (Exception)
                        {
                            MessageBox.Show($"Outdent failed: {globalSprint}!");
                        }
                    }

                    if (outdentPhase) { sprint.OutlineOutdent(); }

                    var firstPhase = true;
                    outdentSprint = false;
                    foreach (var gantPhase in ganttPhases)
                    {
                        outdentSprint = true;
                        var phase = project.Tasks.Add(gantPhase);
                        if (firstPhase)
                        {
                            firstPhase = false;
                            phase.OutlineIndent();
                        }
                        else if (outdentPhase)
                        {
                            try
                            {
                                phase.OutlineOutdent();
                            }
                            catch (Exception)
                            {
                                MessageBox.Show($"Outdent failed: {globalSprint}, {gantPhase}");
                            }
                        }

                        var firstIssue = true;
                        outdentPhase = false;
                        MSProject.Task task = null;
                        foreach (var issue in issues.Where(x => x.Sprints.Contains(globalSprint) && x.Labels.Contains(gantPhase)))
                        {
                            outdentPhase = true;
                            task = project.Tasks.Add(issue.Name);
                            task.ResourceNames = issue.Ressource;
                            if (!string.IsNullOrWhiteSpace(issue.Estimated)) task.Duration = Convert.ToInt32(issue.Estimated) / 60;
                            task.Start = DateTime.FromOADate(Convert.ToDouble(issue.Created));

                            if (firstIssue)
                            {
                                firstIssue = false;
                                task.OutlineIndent();
                            }
                        }
                    }
                }

                app.Quit();
            }
        }
    }
}
