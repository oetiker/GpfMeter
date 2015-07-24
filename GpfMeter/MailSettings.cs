using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GpfMeter
{
    public class MailSettings
    {
        public IList<string> Recipients { get; set; }
        public string DefaultSender { get; set; }
        public UInt16 Port { get; set; }
        public string SMTPServer { get; set; }
        public bool EnableSSL { get; set; }
        public string User { get; set; }
        public string Password { get; set; }
    }
}
