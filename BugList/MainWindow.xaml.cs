using System;
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
            devs.Add(new Dev("Ankush", "Ankush Sharma"));
            devs.Add(new Dev("Ashish", "Ashish Mittal"));
            devs.Add(new Dev("CanHua", "CanHua Li"));
            devs.Add(new Dev("Dhanaraj", "Dhanaraj Durairaj"));
            devs.Add(new Dev("Gitansh", "Gitansh Garg"));
            devs.Add(new Dev("Harmanpreet", "Harmanpreet Singh"));
            devs.Add(new Dev("Kaustubh", "Kaustubh Choudhary"));
            devs.Add(new Dev("Muralimanohar", "Muralimanohar P"));
            devs.Add(new Dev("Nitya", "Nitya Sandadi"));
            devs.Add(new Dev("Ram", "Ram Narendar"));
            devs.Add(new Dev("Shae", "Shae Hurst"));
            devs.Add(new Dev("Vignesh", "Vignesh Sridhar"));
            //devs.Add(new Dev("Salil", "Salil Kapoor"));

            foreach (Dev dev in devs)
            {
                QueryData data = new QueryData()
                {
                    FilePath = @"D:\OneDrive\Salil Documents\Desktop",
                    Dev = dev,
                    StartDate = new DateTime(2020, 6, 1),
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
                    TfsTeamProjectCollection projects = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(new Uri("https://o365exchange.visualstudio.com"));
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
                        writer.WriteLine("URL,Assigned To,Title,Changed Date,State,Reason,Tags,Change");
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

                            string currentReason = "";
                            if (workItem.Fields.Contains("Reason"))
                            {
                                currentReason = workItem.Fields["Reason"].Value as string;
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
                                                Field reason = revision.Fields.Contains("Reason") ? revision.Fields["Reason"] : null;

                                                if (state != null &&
                                                    state.OriginalValue != state.Value &&
                                                    (state.Value as string == "Completed" || state.Value as string == "Closed" || state.Value as string == "Removed" || state.Value as string == "Resolved" || state.Value as string == "Merged to Prod"))
                                                {
                                                    this.WriteLine(writer, workItem, assignedTo, changedDate, currentReason, "From state '" + state.OriginalValue as string + "' to '" + state.Value as string + "'");
                                                    break;
                                                }
                                                else if (reason != null &&
                                                    reason.OriginalValue != reason.Value)
                                                {
                                                    this.WriteLine(writer, workItem, assignedTo, changedDate, currentReason, "From reason '" + reason.OriginalValue as string + "' to '" + reason.Value as string + "'");
                                                    break;
                                                }
                                                else
                                                {

                                                }
                                            }
                                            else
                                            {
                                                this.WriteLine(writer, workItem, assignedTo, changedDate, currentReason, "");
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

        private void WriteLine(StreamWriter writer, Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem workItem, string assignedTo, DateTime changedDate, string currentReason, string message)
        {
            writer.WriteLine("https://o365exchange.visualstudio.com/Viva%20Ally/_workitems/edit/" + workItem.Id.ToString() + "," + assignedTo + "," + workItem.Title.Replace(',', ';') + "," + changedDate.ToString() + "," + workItem.State + "," + currentReason + "," + workItem.Tags + "," + message);
        }

        Dictionary<string, int> vsoProgress = new Dictionary<string, int>();
        private void VSOQuery_Progress(object sender, ProgressChangedEventArgs e)
        {
            vsoProgress[e.UserState as string] = e.ProgressPercentage;

            string text = "";
            foreach (KeyValuePair<string, int> pair in vsoProgress)
            {
                text += pair.Key + ":" + (pair.Value == -1 ? "Done" : pair.Value.ToString()) + "; ";
            }

            c_vso.Text = text;
        }

        private void VSOQuery_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            QueryData data = e.Result as QueryData;
            vsoProgress[data.Dev.Name] = -1;

            string text = "";
            foreach (KeyValuePair<string, int> pair in vsoProgress)
            {
                text += pair.Key + ":" + (pair.Value == -1 ? "Done" : pair.Value.ToString()) + "; ";
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
