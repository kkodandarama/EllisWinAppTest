using System.Configuration;

namespace EllisWinAppTest
{
    public static class TestData
    {
        public static string Path
        {
            get { return ConfigurationManager.AppSettings["path"]; }
        }

        public static string AltPath
        {
            get { return ConfigurationManager.AppSettings["altpath"]; }
        }

        public static string CSRUsername
        {
            get { return ConfigurationManager.AppSettings["csrusername"]; }
        }

        public static string AppPassword
        {
            get { return ConfigurationManager.AppSettings["apppassword"]; }
        }

        public static string AppDomain
        {
            get { return ConfigurationManager.AppSettings["appdomain"]; }
        }

        public static string ARMUsername
        {
            get { return ConfigurationManager.AppSettings["armusername"]; }
        }

        public static string PRMUsername
        {
            get { return ConfigurationManager.AppSettings["prmusername"]; }
        }

        public static string DMUsername
        {
            get { return ConfigurationManager.AppSettings["dmusername"]; }
        }

        public static string ARRUsername
        {
            get { return ConfigurationManager.AppSettings["arrusername"]; }
        }

        public static string AVPUsername
        {
            get { return ConfigurationManager.AppSettings["avpusername"]; }
        }

        public static string NABSUsername
        {
            get { return ConfigurationManager.AppSettings["nabsusername"]; }
        }

        public static string NAPSUsername
        {
            get { return ConfigurationManager.AppSettings["napsusername"]; }
        }

        public static string NAAMUsername
        {
            get { return ConfigurationManager.AppSettings["naamusername"]; }
        }

        public static string TAUsername
        {
            get { return ConfigurationManager.AppSettings["tausername"]; }
        }

        public static string RAUsername
        {
            get { return ConfigurationManager.AppSettings["rausername"]; }
        }

        public static string PWUsername
        {
            get { return ConfigurationManager.AppSettings["pwusername"]; }
        }

        public static string GSUsername
        {
            get { return ConfigurationManager.AppSettings["gsusername"]; }
        }

        public static string GLUsername
        {
            get { return ConfigurationManager.AppSettings["glusername"]; }
        }
    }
}