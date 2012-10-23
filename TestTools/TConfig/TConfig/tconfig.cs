using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;

namespace TConfig
{
    public class TConfig
    {
        private XmlDocument m_doc = null;

        //xpath definitions
        private string m_xInstalledBuildNumber = "Studio13/TestEnvironment/InstallScriptLocation";

        public TConfig(string path)
        {
            m_doc = new XmlDocument();
            m_doc.Load(path);
        }

        public string InstalledBuildNumber()
        {
            string sBuildNumber = null;
            XmlNode root = m_doc.DocumentElement;
            string sInstallScriptLocation = root.SelectSingleNode(m_xInstalledBuildNumber).InnerText;

            //the install scripts create a single dir at <installscriptlocation>\install\buildno\
            //those scripts delete this dir on uninstall
            //if more than 1 dir then we don't know which is current so we fail
            List<string> d = new List<string>(Directory.EnumerateDirectories(sInstallScriptLocation));
            if (1 != d.Count)
                throw new ApplicationException("More than 1 subfolder under InstallScriptLocation so unable to determine build number");
            DirectoryInfo di = new DirectoryInfo(d[0]);
            sBuildNumber = di.Name;

            return sBuildNumber;
        }
    }
}
