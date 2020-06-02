using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Web.UI;
using System.IO;
using System.Net;
using Newtonsoft.Json;
using Microsoft.Win32;
using iText.Html2pdf;
using iText.Kernel.Pdf;
using iText.Kernel.Utils;
using System.Threading.Tasks;

namespace SignupGeniusPDF
{
    public partial class MainWindow : Window
    {
        static string APIKey;
        static readonly string BaseURL = "https://api.signupgenius.com/v2/k";
        private List<SignupCheckBox> signupCheckboxList = new List<SignupCheckBox>();
        public MainWindow()
        {
            InitializeComponent();
            APIKeyInput.Text = Properties.Settings.Default.APIKey;
            APIKey = Properties.Settings.Default.APIKey;
            if (!string.IsNullOrWhiteSpace(APIKey))
            {
                PopulateSignupList();
            }
            dateInput.Text = DateTime.Today.ToString("MM-dd-yyyy");
            separatePDFCheckBox.IsChecked = Properties.Settings.Default.SeparatePDF;


            string[] args = Environment.GetCommandLineArgs();
            bool isAuto = false;
            string autoFilePath = string.Empty;

            Dictionary<string, string> arguments = new Dictionary<string, string>();
            for(int index = 1; index < args.Length; index += 2)
            {
                string arg = args[index].Replace("-", "");
                arguments.Add(arg, args[index + 1]);
            }
            if (arguments.ContainsKey("auto"))
            {
                if (arguments["auto"] == true.ToString())
                {
                    isAuto = true;
                }
            }
            if (arguments.ContainsKey("outputFolder"))
            {
                autoFilePath = arguments["outputFolder"];
                autoFilePath = autoFilePath.Trim();
            }
            arguments.TryGetValue("SeparateFiles", out string SeparateFilesString);
            if(!bool.TryParse(SeparateFilesString, out bool SeparateFiles))
            {
                SeparateFiles = false;
            }

            if (isAuto && !string.IsNullOrWhiteSpace(autoFilePath))
            {
                string dateText = DateTime.Now.ToString("MM-dd-yyyy");
                var checkboxList = new List<SignUp>();
                foreach (var signUp in signupCheckboxList)
                {
                    if (signUp.IsChecked.GetValueOrDefault())
                    {
                        checkboxList.Add(signUp.SignUp);
                    }
                }
                Directory.CreateDirectory(autoFilePath);
                Task.Run(() =>
                {
                    Dispatcher.Invoke((Action)(() => StatusUpdates.Content = "Retrieving SignUp Data"));
                    Dispatcher.Invoke((Action)(() => SaveButton.IsEnabled = false));

                    var allSlots = GetReportData(dateText, checkboxList);
                    if (allSlots.Count > 0)
                    {
                        Dispatcher.Invoke((Action)(() => StatusUpdates.Content = "Filtering reports for the selected date"));
                        var filteredSlots = FilterReports(allSlots, dateText);
                        if (filteredSlots.Count > 0)
                        {
                            CreateHTML(filteredSlots, autoFilePath, SeparateFiles, DateTime.Now, autoFilePath + "Training Rosters " + DateTime.Now.ToString("M-dd-yyyy"));
                        }
                        else
                        {
                            MessageBox.Show("No sign ups for " + DateTime.Now.ToString("M-dd-yyyy"), "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No sign ups for " + DateTime.Now.ToString("M-dd-yyyy"), "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    Dispatcher.Invoke((Action)(() => StatusUpdates.Content = ""));
                    Dispatcher.Invoke((Action)(() => SaveButton.IsEnabled = true));
                });
            }
        }
        public void CreateHTML(Dictionary<SignUp, List<ReportSlot>> SignUpData, string FolderPath, bool SeparateFile, DateTime RequestedDate, string FileName)
        {
            string FilePath = FileName;
            List<string> AllPages = new List<string>();
            foreach(var roster in SignUpData)
            {
                Dispatcher.Invoke((Action)(() => StatusUpdates.Content = string.Format("Creating PDF for {0}", roster.Key.Title)));
                // Bin slots into times
                Dictionary<DateTime, List<ReportSlot>> binnedSlots = new Dictionary<DateTime, List<ReportSlot>>();
                foreach(var slot in roster.Value)
                {
                    if(binnedSlots.ContainsKey(slot.StartTime))
                    {
                        binnedSlots[slot.StartTime].Add(slot);
                    }
                    else
                    {
                        List<ReportSlot> newList = new List<ReportSlot>();
                        newList.Add(slot);
                        binnedSlots.Add(slot.StartTime, newList);
                    }
                }
                // Sort binned slots by time
                binnedSlots = binnedSlots.OrderBy(training => training.Key).ToDictionary(training => training.Key, training => training.Value);

                StringWriter stringWriter = new StringWriter();
                using (HtmlTextWriter writer = new HtmlTextWriter(stringWriter))
                {
                    writer.WriteLine("<!DOCTYPE html>");
                    writer.WriteLine("<html lang=\"en\">");
                    writer.RenderBeginTag(HtmlTextWriterTag.Head);
                    writer.WriteLine();
                    writer.WriteLine("<meta charset=\"utf-8\" />");
                    writer.RenderBeginTag(HtmlTextWriterTag.Title);
                    writer.Write("Roster for WSTPC Training Sign Ups");
                    writer.RenderEndTag();
                    writer.AddAttribute(HtmlTextWriterAttribute.Rel, "stylesheet");
                    writer.AddAttribute(HtmlTextWriterAttribute.Type, "text/css");
                    writer.AddAttribute(HtmlTextWriterAttribute.Href, "https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css");
                    writer.RenderBeginTag(HtmlTextWriterTag.Link);
                    writer.RenderEndTag();
                    writer.RenderBeginTag(HtmlTextWriterTag.Style);
                    writer.WriteLine("tr,h5,h3{font-family:Helvetica;}");
                    writer.RenderEndTag();
                    writer.RenderEndTag();

                    writer.RenderBeginTag(HtmlTextWriterTag.Body);
                    writer.RenderBeginTag(HtmlTextWriterTag.H3);
                    writer.Write(roster.Key.Title + " Training Roster - " + roster.Value[0].StartTime.ToString("M/dd/yyyy"));
                    writer.RenderEndTag();
                    // Begin Tables
                    foreach (var trainingGroup in binnedSlots)
                    {
                        writer.AddAttribute(HtmlTextWriterAttribute.Class, "noPageBreak");
                        writer.AddStyleAttribute("page-break-inside", "avoid");
                        writer.RenderBeginTag(HtmlTextWriterTag.Div);
                        writer.RenderBeginTag(HtmlTextWriterTag.H5);
                        writer.Write(trainingGroup.Key.ToString("h:mm tt") + " - " + trainingGroup.Value[0].SlotName);
                        writer.RenderEndTag();
                        // Begin actual tables
                        writer.AddAttribute(HtmlTextWriterAttribute.Class, "table table-bordered");
                        writer.AddStyleAttribute("font-size", "11pt");
                        writer.RenderBeginTag(HtmlTextWriterTag.Table);
                        writer.RenderBeginTag(HtmlTextWriterTag.Thead);
                        writer.RenderBeginTag(HtmlTextWriterTag.Tr);
                        // Header Rows
                        writer.WriteLine("<th scope=\"col\">First Name</th>");
                        writer.WriteLine("<th scope=\"col\">Last Name</th>");
                        writer.WriteLine("<th scope=\"col\">Email</th>");

                        if (trainingGroup.Value[0].CustomQuestions != null)
                        {
                            for (int i = 0; i < trainingGroup.Value[0].CustomQuestions.Count; i++)
                            {
                                writer.WriteLine("<th scope=\"col\"></th>");
                            }
                        }
                        writer.WriteLine("<th scope=\"col\">Initial</th>");

                        writer.RenderEndTag();
                        writer.RenderEndTag();

                        // List Signup slots
                        writer.RenderBeginTag(HtmlTextWriterTag.Tbody);
                        foreach (var slot in trainingGroup.Value)
                        {
                            writer.RenderBeginTag(HtmlTextWriterTag.Tr);
                            writer.RenderBeginTag(HtmlTextWriterTag.Td);
                            writer.Write(slot.FirstName);
                            writer.RenderEndTag();
                            writer.RenderBeginTag(HtmlTextWriterTag.Td);
                            writer.Write(slot.LastName);
                            writer.RenderEndTag();
                            writer.RenderBeginTag(HtmlTextWriterTag.Td);
                            writer.Write(slot.Email);
                            writer.RenderEndTag();
                            // Custom Questions
                            if (slot.CustomQuestions != null)
                            {
                                for (int i = 0; i < slot.CustomQuestions.Count; i++)
                                {
                                    writer.RenderBeginTag(HtmlTextWriterTag.Td);
                                    writer.Write(slot.CustomQuestions[i].Value);
                                    writer.RenderEndTag();
                                }
                            }
                            writer.RenderBeginTag(HtmlTextWriterTag.Td);
                            writer.RenderEndTag();
                            writer.RenderEndTag();
                        }
                        writer.RenderEndTag(); // </tbody>
                        writer.RenderEndTag(); // </table>
                        writer.RenderEndTag();
                    }
                    writer.RenderEndTag();
                    writer.WriteLine("</html>");
                }

                string tempFilePath = FolderPath + "\\" + roster.Key.Title + " " + RequestedDate.ToString("M-dd-yyyy") + ".pdf";
                PdfDocument PDF = new PdfDocument(new PdfWriter(tempFilePath));
                
                ConverterProperties converterProperties = new ConverterProperties();
                HtmlConverter.ConvertToPdf(stringWriter.ToString(), PDF, converterProperties);
                
                AllPages.Add(tempFilePath);
            }
            Dispatcher.Invoke((Action)(() => StatusUpdates.Content = "Combining rosters and opening PDF"));
            try
            {
                PdfDocument MergedPDF = new PdfDocument(new PdfWriter(FilePath));
                PdfMerger merger = new PdfMerger(MergedPDF);
                foreach(var PDFFilePath in AllPages)
                {
                    PdfDocument PDF = new PdfDocument(new PdfReader(PDFFilePath));
                    merger.Merge(PDF, 1, PDF.GetNumberOfPages());
                    PDF.Close();
                    if (!SeparateFile)
                    {
                        File.Delete(PDFFilePath);
                    }
                    else
                    {
                        System.Diagnostics.Process.Start(PDFFilePath);
                    }
                }
                MergedPDF.Close();
                System.Diagnostics.Process.Start(FilePath);
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message, "Error Saving File", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        public List<SignUp> GetActiveSignups()
        {
            List<SignUp> signUpList = new List<SignUp>();
            try
            {
                var client = new WebClient();
                client.Headers.Add(HttpRequestHeader.Accept, "application/json");
                var response = client.DownloadString(BaseURL + "/signups/created/active/?user_key=" + APIKey);
                FullSignUpReturn returnData = JsonConvert.DeserializeObject<FullSignUpReturn>(response);
                signUpList.AddRange(returnData.Data.ToList());
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "API Retrieval Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            return signUpList;
            }                                                                           
        private void UpdateAPIKey(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.APIKey = APIKeyInput.Text;
            Properties.Settings.Default.Save();
            APIKey = APIKeyInput.Text;

            PopulateSignupList();
        }
        public void PopulateSignupList()
        {
            List<SignUp> signUps = GetActiveSignups();
            signupCheckboxList.Clear();
            foreach (SignUp signUp in signUps)
            {
                SignupCheckBox box = new SignupCheckBox
                {
                    Content = signUp.Title,
                    IsChecked = true,
                    SignUp = signUp
                };
                signupCheckboxList.Add(box);
            }
            SignupSelection.ItemsSource = null;
            SignupSelection.ItemsSource = signupCheckboxList;
            SignupSelection.Visibility = Visibility.Visible;
        }
        private void SavePDF(object sender, RoutedEventArgs e)
        {
            if (!DateTime.TryParse(dateInput.Text, out DateTime requestedDate))
            {
                MessageBox.Show("Date is not valid", "Date Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string initDirectory;
            if (string.IsNullOrWhiteSpace(Properties.Settings.Default.SaveLocation))
            {
                initDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            else
            {
                initDirectory = Properties.Settings.Default.SaveLocation;
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog
            { 
                Title = "Save PDF",
                Filter = "PDF File (*.pdf)|*.pdf",
                InitialDirectory = initDirectory,
                FileName = "Training Rosters " + requestedDate.ToString("M-dd-yyyy")
            };
            if (saveFileDialog.ShowDialog() == true)
            {
                bool separateFiles = separatePDFCheckBox.IsChecked.GetValueOrDefault();
                Properties.Settings.Default.SaveLocation = Path.GetDirectoryName(saveFileDialog.FileName);
                Properties.Settings.Default.SeparatePDF = separateFiles;
                Properties.Settings.Default.Save();

                string dateText = dateInput.Text;
                var checkboxList = new List<SignUp>();
                foreach(var signUp in signupCheckboxList)
                {
                    if (signUp.IsChecked.GetValueOrDefault())
                    {
                        checkboxList.Add(signUp.SignUp);
                    }
                }
                Task.Run(() =>
                {
                    Dispatcher.Invoke((Action)(() => StatusUpdates.Content = "Retrieving SignUp Data"));
                    Dispatcher.Invoke((Action)(() => SaveButton.IsEnabled = false));

                    var allSlots = GetReportData(dateText, checkboxList);
                    if (allSlots.Count > 0)
                    {
                        Dispatcher.Invoke((Action)(() => StatusUpdates.Content = "Filtering reports for the selected date"));
                        var filteredSlots = FilterReports(allSlots, dateText);
                        if (filteredSlots.Count > 0)
                        {
                            CreateHTML(filteredSlots, Path.GetDirectoryName(saveFileDialog.FileName), separateFiles, requestedDate, saveFileDialog.FileName);
                        }
                        else
                        {
                            MessageBox.Show("No sign ups for " + requestedDate.ToString("M-dd-yyyy"), "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No sign ups for " + requestedDate.ToString("M-dd-yyyy"), "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    Dispatcher.Invoke((Action)(() => StatusUpdates.Content = ""));
                    Dispatcher.Invoke((Action)(() => SaveButton.IsEnabled = true));
                });
            };
        }
        public Dictionary<SignUp, ReportSlot[]> GetReportData(string dateString, List<SignUp> checkboxList)
        {
            Dictionary<SignUp, ReportSlot[]> signupSlots = new Dictionary<SignUp, ReportSlot[]>();
            if (!DateTime.TryParse(dateString, out DateTime requestedDate))
            {
                MessageBox.Show("Date is not valid", "Date Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
            foreach (SignUp item in checkboxList)
            {
                string signupID = item.SignupID;
                try
                {
                    var client = new WebClient();
                    client.Headers.Add(HttpRequestHeader.Accept, "application/json");
                    var response = client.DownloadString(BaseURL + "/signups/report/all/" + signupID + "/?user_key=" + APIKey);
                    
                    var parsedData = JsonConvert.DeserializeObject<FullReportReturn>(response);
                    if (parsedData.Success)
                    {
                        signupSlots.Add(item, parsedData.DataSection.Slots);
                    }
                }
                catch(Exception e)
                {
                    MessageBox.Show(e.Message, "Report Data API Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return null;
                }
            }
            return signupSlots;
        }
        public Dictionary<SignUp, List<ReportSlot>> FilterReports(Dictionary<SignUp, ReportSlot[]> allSlots, string dateText)
        {
            var requestedDate = DateTime.Parse(dateText);
            Dictionary<SignUp, List<ReportSlot>> filteredSlots = new Dictionary<SignUp, List<ReportSlot>>();
            foreach(var signUp in allSlots)
            {
                List<ReportSlot> validSlots = new List<ReportSlot>();
                foreach(var slot in signUp.Value)
                {
                    slot.StartTime = DateTime.Parse(slot.StartTimeString);
                    if(slot.StartTime.Date == requestedDate.Date)
                    {
                        validSlots.Add(slot);
                    }
                }
                if(validSlots.Count > 0)
                {
                    filteredSlots.Add(signUp.Key, validSlots);
                }
            }            
            return filteredSlots;
        }
    }
    public class SignupCheckBox : CheckBox
    {
        public SignUp SignUp { get; set; }
    }
}