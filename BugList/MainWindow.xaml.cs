using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Azure.Amqp.Framing;
using Microsoft.IdentityModel.Tokens;
using Microsoft.TeamFoundation.Work.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Newtonsoft.Json.Linq;
using static System.Net.Mime.MediaTypeNames;

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

        public MainWindow()
        {
            InitializeComponent();

            List<Dev> devs = new List<Dev>();
            devs.Add(new Dev("Debashis", "Debashis Mondal"));
            devs.Add(new Dev("Nitya", "Nitya Sandadi"));
            devs.Add(new Dev("Shae", "Shae Hurst"));
            devs.Add(new Dev("Krishna", "Govinda Krishnamurthy Kandipilli"));
            devs.Add(new Dev("CanHua", "CanHua Li"));
            devs.Add(new Dev("Kannan", "Kannan Bhoopathy"));
            //devs.Add(new Dev("Ankush", "Ankush Jain 🐎"));
            //devs.Add(new Dev("Sandeep", "Sandeep Prusty"));
            //devs.Add(new Dev("Salil", "Salil Kapoor"));

            foreach (Dev dev in devs)
            {
                VSOQuery(dev, new DateTime(2023, 6, 1), true);
            }
        }

        private async void VSOQuery(Dev dev, DateTime startDate, bool lookForResolvedOnly)
        {
            UpdateProgress(dev.Name, 0);
            var credentials = new VssBasicCredential(string.Empty, "");

            // create a wiql object and build our query
            var wiql = new Wiql()
            {
                // NOTE: Even if other columns are specified, only the ID & URL are available in the WorkItemReference
                Query = "Select * " +
                        "From WorkItems " +
                        "Where [System.ChangedBy] Ever '" + dev.NameInVSO + "' " +
                        "And [System.ChangedDate] > '" + startDate.ToShortDateString() + "' "
            };

            // create instance of work item tracking http client
            using (var httpClient = new WorkItemTrackingHttpClient(new Uri("https://o365exchange.visualstudio.com"), credentials))
            {
                // execute the query to get the list of work items in the results
                var result = await httpClient.QueryByWiqlAsync(wiql);

                var ids = result.WorkItems.Select(item => item.Id).ToList();
                if (ids.Count == 0)
                {
                    //error
                }

                int batchSize = 200;
                List<WorkItem> workItems = new List<WorkItem>();
                for (var i = 0; i < (float)ids.Count / batchSize; i++)
                {
                    var partIds = ids.Skip(i * batchSize).Take(batchSize);

                    // build a list of the fields we want to see
                    //var fields = new[] { "System.Id", "System.Title", "System.State" };

                    // get work items for the ids found in query
                    var partWorkItems = await httpClient.GetWorkItemsAsync(partIds, null, result.AsOf);
                    workItems.AddRange(partWorkItems);
                }

                using (MemoryStream stream = new MemoryStream())
                using (StreamWriter writer = new StreamWriter(stream))
                {
                    writer.WriteLine("URL,Assigned To,Title,Changed Date,State,Reason,Tags,Change");
                    for (int i = workItems.Count - 1; i >= 0; i--)
                    {
                        UpdateProgress(dev.Name, (int)(((workItems.Count - i) * 100.00) / workItems.Count));

                        WorkItem workItem = workItems[i];
                        var revisions = await httpClient.GetRevisionsAsync((int)workItem.Id);

                        for (int j = revisions.Count - 1; j >= 0; j--)
                        {
                            WorkItem revision = revisions[j];
                            DateTime changedDate = (DateTime)revision.Fields["System.ChangedDate"];
                            if (changedDate > startDate)
                            {
                                string changedBy = (revision.Fields["System.ChangedBy"] as IdentityRef).DisplayName;
                                if (changedBy.Contains(dev.NameInVSO))
                                {
                                    if (lookForResolvedOnly)
                                    {
                                        var thisState = (string)revision.Fields["System.State"];
                                        var thisReason = (string)revision.Fields["System.Reason"];
                                        var nextState = j > 0 ? (string)revisions[j - 1].Fields["System.State"] : "";
                                        var nextReason = j > 0 ? (string)revisions[j - 1].Fields["System.Reason"] : "";

                                        if ((thisState != nextState) && (thisState == "Completed" || thisState == "Closed" || thisState == "Removed" || thisState == "Resolved"))
                                        {

                                            this.WriteLine(writer, workItem, changedDate, "From state '" + nextState + "' to '" + thisState + "'");
                                            break;
                                        }
                                        else if (thisReason != nextReason)
                                        {
                                            this.WriteLine(writer, workItem, changedDate, "From reason '" + nextReason + "' to '" + thisReason + "'");
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        this.WriteLine(writer, workItem, changedDate, "");
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

                    WriteFiles(dev, stream.ToArray());
                    UpdateProgress(dev.Name, -1);
                }
            }
        }

        private IDictionary<string, int> devProgress = new Dictionary<string, int>();
        private void UpdateProgress(string dev, int progress)
        {
            devProgress[dev] = progress;

            string text = "";
            foreach (KeyValuePair<string, int> pair in devProgress)
            {
                text += pair.Key + ":" + (pair.Value == -1 ? "Done" : pair.Value.ToString()) + "; ";
            }

            c_vso.Text = text;
        }

        private void WriteLine(StreamWriter writer, WorkItem workItem, DateTime changedDate, string message)
        {
            string assignedTo = workItem.Fields.ContainsKey("System.AssignedTo") ? (workItem.Fields["System.AssignedTo"] as IdentityRef).DisplayName : "";
            string title = (string)workItem.Fields["System.Title"];
            string state = (string)workItem.Fields["System.State"];
            string reason = (string)workItem.Fields["System.Reason"];
            string tags = "";

            writer.WriteLine("https://o365exchange.visualstudio.com/Viva%20Ally/_workitems/edit/" + workItem.Id.ToString() + "," + assignedTo + "," + title.Replace(',', ';') + "," + changedDate.ToString() + "," + state + "," + reason + "," + tags + "," + message);
        }

        private void WriteFiles(Dev dev, byte[] data)
        {
            using (FileStream file = File.Create(System.IO.Path.Combine(@"D:\OneDrive\Salil Documents\Desktop", dev.Name + ".csv")))
            {
                file.Write(data, 0, data.Length);
            }
        }
    }
}
