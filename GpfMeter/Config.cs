using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using System.Windows;
using Newtonsoft.Json;

namespace GpfMeter
{
    public class Config
    {
        public string Button0 { get; set; }
        public string Button1 { get; set; }
        public string Button2 { get; set; }
        public string Button3 { get; set; }
        public string Radio0 { get; set; }
        public string Radio1 { get; set; }
        public string Radio2 { get; set; }
        public IList<string>  UrlPingTargets { get; set; }
        public MailSettings   MailSettings { get;set; }

        public static Config Load() {
            string path = new FileInfo(Assembly.GetExecutingAssembly().GetName().Name + "Config.json").FullName.ToString();
            if (!File.Exists(path)){
                File.WriteAllText(path, @"{
    'Button0': 'Login dauert lange',
    'Button1': 'Sprachqualität schlecht',
    'Button2': 'Website langsam',
    'Button3': 'PC langsam',
    'Radio0': 'Bemerke ich zum ersten Mal',
    'Radio1': 'kommt in letzter Zeit häufiger vor',
    'Radio2': 'nervt mich schon lange',
    'UrlPingTargets': [
        'https://www.google.com',
        'https://www.admin.ch'
    ],
    'MailSettings': {
        'Recipients': [
            'support@example.com'
        ],
        'DefaultSender': 'support@example.com',
        'Port': 587,
        'SMTPServer': 'smtp.exampl.com',
        'EnableSSL': true,
        'User': 'dummy',
        'Password': 'mostsecret'
    }
}");
            }
            try {
                return JsonConvert.DeserializeObject<Config>(File.ReadAllText(path));
            } catch ( JsonException e){
                MessageBox.Show(string.Format("Failed to read config file {0}: {1}", path, e.ToString()),"JSON Error");
                Application.Current.Shutdown();
            }
            return new Config();
        }
    }
}
