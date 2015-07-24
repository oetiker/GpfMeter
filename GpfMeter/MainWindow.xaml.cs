using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.Drawing;
using System.Globalization;
using System.Management;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using Forms = System.Windows.Forms;
using Imaging = System.Drawing.Imaging;
using IO = System.IO;
using GpfMeter.Commands;

namespace GpfMeter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private IO.MemoryStream screenshot;
        private string RadioMessage = null;
        public Config cfg = Config.Load();

        public MainWindow()
        {
            InitializeComponent();
            Hide();
            Closing += new CancelEventHandler(MainWindow_Closing);
            updateUi();
        }

        // get screenshot bevore showing the display
        public void StartGpfMeter()
        {
            if (this.IsVisible == false)
            {
                screenshot = getScreenshot();
                UserNote.Text = "";
                SendScreenshot.IsChecked = true;
                this.Show();
            }
            this.Activate();
        }

        private void updateUi()
        {
            Btn0.Content = cfg.Button0;
            Btn1.Content = cfg.Button1;
            Btn2.Content = cfg.Button2;
            Btn3.Content = cfg.Button3;
            Rb0.Content = cfg.Radio0;
            Rb0.IsChecked = true;
            RadioMessage = cfg.Radio0;
            Rb1.Content = cfg.Radio1;
            Rb2.Content = cfg.Radio2;
        }

        private IO.MemoryStream attachScreenshot(MailMessage eMail)
        {
            if (screenshot == null)
            {
                return null;
            }
            ContentType ct = new ContentType();
            ct.MediaType = "image/png";
            ct.Name = FormatFilename("Screenshot.png");
            eMail.Attachments.Add(new Attachment(screenshot, ct));
            eMail.IsBodyHtml = false;
            eMail.BodyEncoding = UTF8Encoding.UTF8;
            return screenshot;
        }

        private IO.MemoryStream attachProcessList(MailMessage eMail)
        {
            var counter = new Dictionary<long, Dictionary<string, double>>();

            ObjectQuery sq = new ObjectQuery("SELECT * FROM Win32_Process");
            ManagementObjectCollection pl1 = (new ManagementObjectSearcher(sq)).Get();

            var counters = new string[] { "KernelModeTime", "UserModeTime", "PageFaults" };
            foreach (ManagementObject p in pl1)
            {
                var data = new Dictionary<string, double>();
                foreach (string key in counters)
                {
                    data.Add(key, Convert.ToDouble(p[key]));
                }
                counter.Add(Convert.ToInt64(p["ProcessId"].ToString()), data);
            }

            Thread.Sleep(1000);

            var memoryStream = new IO.MemoryStream();
            var sw = new IO.StreamWriter(memoryStream);

            string format = "{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10}";

            sw.WriteLine(format, "Name", "Owner", "Session", "Kernel", "User", "WorkingSetSize",
                "PageFileUsage", "PageFaults", "ThreadCount", "CreationDate", "StartOffset");

            ManagementObjectCollection pl2 = (new ManagementObjectSearcher(sq)).Get();

            double sum = 0;

            foreach (ManagementObject p in pl2)
            {
                long pid = Convert.ToInt64(p["ProcessId"].ToString());
                var usage = new Dictionary<string, double>();
                try
                {
                    var start = counter[pid];
                    foreach (string key in counters)
                    {
                        usage[key] = Convert.ToDouble(p[key]) - start[key];
                    }
                }
                catch (KeyNotFoundException)
                {
                    foreach (string key in counters)
                    {
                        usage[key] = Convert.ToDouble(p[key]);
                    }
                }
                counter[pid] = usage;
                sum += usage["UserModeTime"] + usage["KernelModeTime"];
            }

            var bootTime = getBootTime();

            foreach (ManagementObject p in pl2)
            {
                string[] OwnerInfo = new string[2];
                long pid = Convert.ToInt64(p["ProcessId"].ToString());
                // skip system idle process
                if (pid == 0)
                {
                    continue;
                }
                var usage = counter[pid];
                double kcpu = usage["KernelModeTime"] / (sum / 100);
                double ucpu = usage["UserModeTime"] / (sum / 100);

                try
                {
                    p.InvokeMethod("GetOwner", (object[])OwnerInfo);
                }
                catch (Exception e)
                {
                    Console.WriteLine("PID {0}: {1}", pid, e);
                }
                var creation = ManagementDateTimeConverter.ToDateTime(p["CreationDate"].ToString());

                sw.WriteLine(format,
                    p["Name"],
                    OwnerInfo[0],
                    p["SessionId"],
                    kcpu.ToString("F1", CultureInfo.InvariantCulture),
                    ucpu.ToString("F1", CultureInfo.InvariantCulture),
                    p["WorkingSetSize"],
                    p["PageFileUsage"],
                    usage["PageFaults"],
                    p["ThreadCount"],
                    creation.ToString("o"),
                    creation.Subtract(bootTime).TotalSeconds.ToString("F3", CultureInfo.InvariantCulture)
                );
            }

            ContentType ct = new ContentType();
            ct.MediaType = "text/csv";
            ct.Name = FormatFilename("ProcessList.csv");
            sw.Flush();
            memoryStream.Seek(0, IO.SeekOrigin.Begin);
            eMail.Attachments.Add(new Attachment(memoryStream, ct));
            return memoryStream;
        }

        private IO.MemoryStream attachWebPing(MailMessage eMail)
        {
            if (cfg.UrlPingTargets.Count == 0) return null;
            var memoryStream = new IO.MemoryStream();
            var sw = new IO.StreamWriter(memoryStream);

            string format = "{0};{1};{2};{3};{4}";

            sw.WriteLine(format, "Method", "Url", "Date/Time", "ElapsedMs", "Status");


            foreach (string url in cfg.UrlPingTargets)
            {
                var request = WebRequest.Create(url);
                request.Method = "GET";
                request.Timeout = 30 * 1000;
                var timer = Stopwatch.StartNew();
                var start = DateTime.Now.ToString("o", CultureInfo.InvariantCulture);
                timer.Start();
                //get the server response
                try
                {
                    var response = (HttpWebResponse)request.GetResponse();
                    timer.Stop();
                    sw.WriteLine(format, request.Method, url, start, timer.ElapsedMilliseconds.ToString(), response.StatusDescription);
                    // System.Diagnostics.Debug.WriteLine(format, request.Method, url, start, timer.ElapsedMilliseconds.ToString(), response.StatusDescription);
                    response.Close();
                }
                catch (System.Exception ex)
                {
                    sw.WriteLine(format, request.Method, url, start, "-", ex.Message);
                }
            }
            ContentType ct = new ContentType();
            ct.MediaType = "text/csv";
            ct.Name = FormatFilename("WebPings.csv");
            sw.Flush();
            memoryStream.Seek(0, IO.SeekOrigin.Begin);
            eMail.Attachments.Add(new Attachment(memoryStream, ct));
            return memoryStream;
        }

        private DateTime getBootTime()
        {
            ObjectQuery osq = new ObjectQuery("Select * from Win32_OperatingSystem");
            ManagementObjectCollection os = (new ManagementObjectSearcher(osq)).Get();

            foreach (ManagementObject p in os)
            {
                return ManagementDateTimeConverter.ToDateTime(p["LastBootUpTime"].ToString());
            }
            throw new Exception("No Bootup Time Found");
        }

        private IO.MemoryStream getScreenshot()
        {
            Bitmap screenshot = new Bitmap(
                Forms.Screen.PrimaryScreen.Bounds.Width,
                Forms.Screen.PrimaryScreen.Bounds.Height,
                Imaging.PixelFormat.Format32bppArgb);

            Graphics gfxScreenshot = Graphics.FromImage(screenshot);

            gfxScreenshot.CopyFromScreen(
                Forms.Screen.PrimaryScreen.Bounds.X,
                Forms.Screen.PrimaryScreen.Bounds.Y, 0, 0,
                Forms.Screen.PrimaryScreen.Bounds.Size,
                CopyPixelOperation.SourceCopy);

            IO.MemoryStream image = new IO.MemoryStream();
            screenshot.Save(image, Imaging.ImageFormat.Png);
            image.Seek(0, IO.SeekOrigin.Begin);
            gfxScreenshot.Dispose();
            return image;
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            this.Hide();
            if (screenshot != null)
            {
                screenshot.Dispose();
                screenshot = null;
            }
            e.Cancel = true;
        }

        private void RadioClicked(object sender, RoutedEventArgs e)
        {
            RadioMessage = ((RadioButton)sender).Content.ToString();
        }

        private void SendButtonClicked(object sender, RoutedEventArgs e)
        {
            this.Hide();
            if (SendScreenshot.IsChecked == false)
            {
                screenshot.Dispose();
                screenshot = null;
            }
            string message = ((Button)sender).Content.ToString();
            var eMail = new MailMessage();
            UserPrincipal userPrincipal = UserPrincipal.Current;
            eMail.From = new MailAddress(userPrincipal.EmailAddress ?? cfg.MailSettings.DefaultSender);
            foreach (string address in cfg.MailSettings.Recipients)
            {
                eMail.To.Add(address);
            }
            eMail.Subject = message;
            eMail.Body = string.Format("Problem: {0}\nFrequency: {1}\nUser Message: {2}", message, RadioMessage, UserNote.Text);
            var t = new Thread(() => sendEMail(eMail));
            t.Start();
        }

        // method to send email via SMTP
        private void sendEMail(MailMessage eMail)
        {
            SmtpClient client = new SmtpClient();

            var streams = new List<IO.MemoryStream>();
            streams.Add(attachWebPing(eMail));
            streams.Add(attachScreenshot(eMail));
            streams.Add(attachProcessList(eMail));
            try
            {
                client.Port = cfg.MailSettings.Port;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = true;
                string user = cfg.MailSettings.User;
                string password = cfg.MailSettings.Password;
                if (user != null && password != null)
                {
                    client.Credentials = new System.Net.NetworkCredential(user, password);
                }
                client.EnableSsl = cfg.MailSettings.EnableSSL;
                client.Host = cfg.MailSettings.SMTPServer;
                client.Send(eMail);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "GpfMeter Message Exception", MessageBoxButton.OK);
            }
            eMail.Dispose();
            foreach (var stream in streams)
            {
                if (stream != null)
                {
                    stream.Dispose();
                }
            }
        }

        private void QuitApplication(object sender, RoutedEventArgs e)
        {
            this.GpfMeterNotifyIcon.Dispose();
            Application.Current.Shutdown();
        }
        private string FormatFilename(string name)
        {
            var now = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss", CultureInfo.InvariantCulture);
            var user = UserPrincipal.Current.SamAccountName.ToString();
            return string.Format("{0}_{1}_{2}", now, user, name);
        }
    }
} 