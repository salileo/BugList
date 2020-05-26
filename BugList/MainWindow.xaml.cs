﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.SourceControl.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;

namespace BugList
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private class Dev
        {
            public string Name { get; set; }
            public string NameInVSO { get; set; }

            public Dev(string name, string vso)
            {
                this.Name = name;
                this.NameInVSO = vso;
            }
        }

        private class QueryData
        {
            public string FilePath { get; set; }
            public Dev Dev { get; set; }
            public DateTime StartDate { get; set; }
            public bool LookForResolvedOnly { get; set; }
            public byte[] VSOData { get; set; }
            public bool VSOTaskDone { get; set; }
        };

        public MainWindow()
        {
            InitializeComponent();

            List<Dev> devs = new List<Dev>();
            devs.Add(new Dev("Tim", "Tim McBride (ASP.NET)"));
            devs.Add(new Dev("Guru", "Guru Kumaraguru"));
            devs.Add(new Dev("Lucas", "Lucas Stanford"));
            devs.Add(new Dev("Petre", "Petre Munteanu"));
            devs.Add(new Dev("Chris", "Christopher Scrosati"));
            devs.Add(new Dev("Julian", "Julian Dominguez"));
            devs.Add(new Dev("Rosy", "Rosy Chen"));
            devs.Add(new Dev("Sammy", "Sammy Israwi"));
            devs.Add(new Dev("Salil", "Salil Kapoor"));

            foreach (Dev dev in devs)
            {
                QueryData data = new QueryData()
                {
                    FilePath = @"C:\Users\Salilk\Desktop",
                    Dev = dev,
                    StartDate = new DateTime(2019, 5, 1),
                    LookForResolvedOnly = true
                };

                BackgroundWorker vsoTask = new BackgroundWorker();
                vsoTask.WorkerReportsProgress = true;
                vsoTask.WorkerSupportsCancellation = true;
                vsoTask.DoWork += new DoWorkEventHandler(VSOQuery);
                vsoTask.ProgressChanged += new ProgressChangedEventHandler(VSOQuery_Progress);
                vsoTask.RunWorkerCompleted += new RunWorkerCompletedEventHandler(VSOQuery_Completed);
                vsoTask.RunWorkerAsync(data);
            }
        }

        private void VSOQuery(object sender, DoWorkEventArgs workArgs)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            QueryData data = workArgs.Argument as QueryData;
            string vsoQuery = string.Format(@"
                select *
                from WorkItems
                where [System.ChangedBy] ever '{0}'
                and [System.ChangedDate] > '{1}'", data.Dev.NameInVSO, data.StartDate.ToShortDateString());

            try
            {
                using (MemoryStream stream = new MemoryStream())
                {
                    TfsTeamProjectCollection projects = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri("https://msazure.visualstudio.com"));
                    if (worker.CancellationPending)
                    {
                        workArgs.Cancel = true;
                        return;
                    }

                    WorkItemStore workItemStore = projects.GetService<WorkItemStore>();
                    if (worker.CancellationPending)
                    {
                        workArgs.Cancel = true;
                        return;
                    }

                    WorkItemCollection workItems = workItemStore.Query(vsoQuery);
                    if (worker.CancellationPending)
                    {
                        workArgs.Cancel = true;
                        return;
                    }

                    using (StreamWriter writer = new StreamWriter(stream))
                    {
                        for (int i = workItems.Count - 1; i >= 0; i--)
                        {
                            if (worker.CancellationPending)
                            {
                                workArgs.Cancel = true;
                                return;
                            }

                            worker.ReportProgress((int)(((workItems.Count - i) * 100.00) / workItems.Count), data.Dev.Name);

                            Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem workItem = workItems[i];
                            if (workItem.Revisions == null)
                            {
                                continue;
                            }

                            string assignedTo = "";
                            if (workItem.Fields.Contains("Assigned To"))
                            {
                                assignedTo = workItem.Fields["Assigned To"].Value as string;
                            }

                            string currentStatus = "";
                            if (workItem.Fields.Contains("Status"))
                            {
                                currentStatus = workItem.Fields["Status"].Value as string;
                            }
                            string currentResolution = "";
                            if (workItem.Fields.Contains("Resolution_Custom"))
                            {
                                currentResolution = workItem.Fields["Resolution_Custom"].Value as string;
                            }

                            for (int j = workItem.Revisions.Count - 1; j >= 0; j--)
                            {
                                if (worker.CancellationPending)
                                {
                                    workArgs.Cancel = true;
                                    return;
                                }

                                Revision revision = workItem.Revisions[j];
                                if (revision.Fields.Contains("Changed By") && revision.Fields.Contains("Changed Date"))
                                {
                                    DateTime changedDate = (DateTime)revision.Fields["Changed Date"].Value;
                                    if (changedDate > data.StartDate)
                                    {
                                        string changedBy = revision.Fields["Changed By"].Value as string;
                                        if (changedBy.Contains(data.Dev.NameInVSO))
                                        {
                                            if (data.LookForResolvedOnly)
                                            {
                                                Field state = revision.Fields.Contains("State") ? revision.Fields["State"] : null;
                                                Field status = revision.Fields.Contains("Status") ? revision.Fields["Status"] : null;
                                                Field resolution = revision.Fields.Contains("Resolution_Custom") ? revision.Fields["Resolution_Custom"] : null;

                                                if (state != null &&
                                                    state.OriginalValue != state.Value &&
                                                    (state.Value as string == "Done" || state.Value as string == "Removed"))
                                                {
                                                    writer.WriteLine("https://msazure.visualstudio.com/One/_workitems/edit/" + workItem.Id.ToString() + ", " + assignedTo + ", " + workItem.Title.Replace(',', ';') + ", " + changedDate.ToString() + ", " + workItem.State + ", " + currentStatus + ", " + currentResolution + ", " + workItem.Tags + ", " + "From state '" + state.OriginalValue as string + "' to '" + state.Value as string + "'");
                                                    break;
                                                }
                                                else if (status != null &&
                                                    status.OriginalValue != status.Value &&
                                                    status.Value as string != "" &&
                                                    status.Value as string != "New" &&
                                                    status.Value as string != "Blocked" &&
                                                    status.Value as string != "Delayed" &&
                                                    status.Value as string != "In Progress" &&
                                                    status.Value as string != "In Review" &&
                                                    status.Value as string != "Forecasted")
                                                {
                                                    writer.WriteLine("https://msazure.visualstudio.com/One/_workitems/edit/" + workItem.Id.ToString() + ", " + assignedTo + ", " + workItem.Title.Replace(',', ';') + ", " + changedDate.ToString() + ", " + workItem.State + ", " + currentStatus + ", " + currentResolution + ", " + workItem.Tags + ", " + "From status '" + status.OriginalValue as string + "' to '" + status.Value as string + "'");
                                                    break;
                                                }
                                                else if (resolution != null &&
                                                    resolution.OriginalValue != resolution.Value)
                                                {
                                                    writer.WriteLine("https://msazure.visualstudio.com/One/_workitems/edit/" + workItem.Id.ToString() + ", " + assignedTo + ", " + workItem.Title.Replace(',', ';') + ", " + changedDate.ToString() + ", " + workItem.State + ", " + currentStatus + ", " + currentResolution + ", " + workItem.Tags + ", " + "From resolution '" + resolution.OriginalValue as string + "' to '" + resolution.Value as string + "'");
                                                    break;
                                                } else
                                                {

                                                }
                                            }
                                            else
                                            {
                                                writer.WriteLine("https://msazure.visualstudio.com/One/_workitems/edit/" + workItem.Id.ToString() + ", " + assignedTo + ", " + workItem.Title.Replace(',', ';') + ", " + changedDate.ToString() + ", " + workItem.State + ", " + currentStatus + ", " + currentResolution + ", " + workItem.Tags);
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }

                    data.VSOData = stream.ToArray();
                }
            }
            finally
            {
                data.VSOTaskDone = true;
                workArgs.Result = data;
            }
        }

        Dictionary<string, int> vsoProgress = new Dictionary<string, int>();
        private void VSOQuery_Progress(object sender, ProgressChangedEventArgs e)
        {
            vsoProgress[e.UserState as string] = e.ProgressPercentage;

            string text = "";
            foreach (KeyValuePair<string, int> pair in vsoProgress)
            {
                text += pair.Key + ":" + pair.Value.ToString() + "; ";
            }

            c_vso.Text = text;
        }

        private void VSOQuery_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            QueryData data = e.Result as QueryData;

            string text = "";
            foreach (KeyValuePair<string, int> pair in vsoProgress)
            {
                text += pair.Key + ":" + (pair.Key == data.Dev.Name ? "Done" : pair.Value.ToString()) + "; ";
            }

            c_vso.Text = text;
            this.WriteFiles(data);
        }

        private void WriteFiles(QueryData data)
        {
            if (!data.VSOTaskDone)
            {
                return;
            }

            using (FileStream file = File.Create(System.IO.Path.Combine(data.FilePath, data.Dev.Name + ".csv")))
            {
                if (data.VSOData != null)
                    file.Write(data.VSOData, 0, data.VSOData.Length);
            }
        }
    }
}