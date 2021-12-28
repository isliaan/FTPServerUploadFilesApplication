using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Linq;
using System.Configuration;
using System.Web;
using System.Net;
using System.Security.Policy;

namespace UploadFTPServerFiles
{
    [RunInstaller(true)]
    public partial class FTPFilesUpload : ServiceBase
    {
        int ScheduleTime = Convert.ToInt32(ConfigurationSettings.AppSettings["ThreadTime"]);
        public Thread Worker = null; 
        public FTPFilesUpload()
        {
            InitializeComponent();
        }

        public void onDebug()
        {
            OnStart(null);
        }

        protected override void OnStart(string[] args)
        {
            // TODO: Add code here to start your service.
            this.ScheduleService();
        }

        protected override void OnStop()
        {
            // TODO: Add code here to perform any tear-down necessary to stop your service.
            this.WriteToFile("File Upload Service stopped: " + DateTime.Now);
            this.Schedular.Dispose();
        }

        public void WriteToFile(string Message)
        {
            //throw new Exception("The method or operation is not implemented.");
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            //string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog" + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }

        private Timer Schedular;

        public void ScheduleService()
        {
            try
            {
                Schedular = new Timer(new TimerCallback(SchedularCallback));
                string mode = ConfigurationSettings.AppSettings["Mode"].ToUpper();
                this.WriteToFile("File Upload Service Mode: " + mode);

                //Set the Default Time.
                DateTime scheduledTime = DateTime.MinValue;

                if (mode == "DAILY")
                {
                    //Get the Scheduled Time from AppSettings.
                    scheduledTime = DateTime.Parse(System.Configuration.ConfigurationSettings.AppSettings["ScheduledTime"]);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next day.
                        scheduledTime = scheduledTime.AddDays(1);
                    }
                }

                if (mode.ToUpper() == "INTERVAL")
                {
                    //Get the Interval in Minutes from AppSettings.
                    int intervalMinutes = Convert.ToInt32(ConfigurationSettings.AppSettings["IntervalMinutes"]);

                    //Set the Scheduled Time by adding the Interval to Current Time.
                    scheduledTime = DateTime.Now.AddMinutes(intervalMinutes);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next Interval.
                        scheduledTime = scheduledTime.AddMinutes(intervalMinutes);
                    }
                }

                string ResponseDescription = "";

                //FTP Folder name. Leave blank if you want to Download file from root folder.
                string ftpFolder = "/";
                byte[] fileBytes = null;
                //string LatestFileName = NewFileName("E:/Development/FTPServerDownloadFiles");

                DataTable dtUpFiles = new DataTable();
                dtUpFiles.Columns.Add("Name");
                dtUpFiles.Columns.Add("Date");
                DirectoryInfo di = new DirectoryInfo(@"E:\Development\FTPServerDownloadFiles\");
                FileInfo[] files = di.GetFiles("*.txt");

                List<FileInfo> lastUpdatedFile = new List<FileInfo>();
                DateTime lastUpdate = DateTime.MaxValue;
                foreach (FileInfo file in files)
                {
                    //if (file.LastWriteTime > lastUpdate)
                    //{                    
                    //    lastUpdatedFile.Add(file);
                    //    lastUpdate = file.LastWriteTime;
                    //}
                    DataRow row = dtUpFiles.NewRow();

                    row["Name"] = file.Name;
                    //row["Date"] = Convert.ToDateTime(file.LastWriteTime).ToString("yyyy-MM-dd HH:mm:ss");
                    row["Date"] = Convert.ToDateTime(file.LastWriteTime).ToString("yyyy-MM-dd HH:mm");
                    dtUpFiles.Rows.Add(row);
                }

                string expression;
                string sortOrder;
                expression = "Name <>'' ";
                sortOrder = "Date DESC";
                DataRow[] foundRows = dtUpFiles.Select(expression, sortOrder, DataViewRowState.Added);

                string LatestFileName = Convert.ToString(foundRows[0].ItemArray[0]);  //Convert.ToString(lastUpdatedFile[0]);
                string LatestFileDateTime = Convert.ToString(foundRows[0].ItemArray[1]);

                string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog" + ".txt";

                string PreviousFileInfo = "";
                string searchForFileInfo = "F_CTM_CONTRACTS_";
                string[] FileLinesInfo = File.ReadAllLines(filepath);

                for (int i = FileLinesInfo.Length - 1; i >= 0; i--)
                {
                    if (FileLinesInfo[i].Contains(searchForFileInfo))
                    {
                        PreviousFileInfo = FileLinesInfo[i].ToString();
                    }
                }

                File.WriteAllText(filepath, string.Empty);
                //this.WriteToFile("File Upload Service Started: " + DateTime.Now.Date.ToShortDateString());
                this.WriteToFile("File Upload Service Started: " + DateTime.Now);
                this.WriteToFile("File Upload Service Mode: " + mode);

                if (PreviousFileInfo != (LatestFileName + " " + LatestFileDateTime))
                {
                    using (StreamReader fileStream = new StreamReader("E:/Development/FTPServerDownloadFiles/" + LatestFileName))
                    {
                        fileBytes = Encoding.UTF8.GetBytes(fileStream.ReadToEnd());
                        fileStream.Close();
                    }

                    try
                    {
                        //Create FTP Request.
                        FtpWebRequest request = (FtpWebRequest)WebRequest.Create("ftp://66.147.236.125" + ftpFolder + LatestFileName);
                        request.Method = WebRequestMethods.Ftp.UploadFile;

                        //Enter FTP Server credentials.
                        request.Credentials = new NetworkCredential("administrator", "NM5VFEhRt7");
                        request.ContentLength = fileBytes.Length;
                        request.UsePassive = true;
                        request.UseBinary = true;
                        request.ServicePoint.ConnectionLimit = fileBytes.Length;
                        request.EnableSsl = false;

                        //Fetch the Response and read it into a MemoryStream object.
                        using (Stream requestStream = request.GetRequestStream())
                        {
                            requestStream.Write(fileBytes, 0, fileBytes.Length);
                            requestStream.Close();
                        }

                        FtpWebResponse response = (FtpWebResponse)request.GetResponse();
                        ResponseDescription = response.StatusDescription;
                        response.Close();
                    }
                    catch (WebException ex)
                    {
                        throw new Exception((ex.Response as FtpWebResponse).StatusDescription);
                    }
                }

                TimeSpan timeSpan = scheduledTime.Subtract(DateTime.Now);
                string schedule = string.Format("{0} day(s) {1} hour(s) {2} minute(s) {3} seconds(s)", timeSpan.Days, timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);

                this.WriteToFile("File Upload Service scheduled to run after: " + schedule);
                this.WriteToFile(LatestFileName + " "+LatestFileDateTime);
                //this.WriteToFile("File Upload Service Ended: " + DateTime.Now.Date.ToShortDateString());
                this.WriteToFile("File Upload Service Ended: " + DateTime.Now);

                //Get the difference in Minutes between the Scheduled and Current Time.
                int dueTime = Convert.ToInt32(timeSpan.TotalMilliseconds);

                //Change the Timer's Due Time.
                Schedular.Change(dueTime, Timeout.Infinite);
            }
            catch (Exception ex)
            {
                WriteToFile("File Upload Service Error on: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("SimpleService"))
                {
                    serviceController.Stop();
                }
            }
        }        

        private void SchedularCallback(object e)
        {
            this.WriteToFile("File Upload Service Log: {0}");
            this.ScheduleService();
        }
    }
}
