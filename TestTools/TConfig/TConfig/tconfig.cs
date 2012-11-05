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
        private string m_xResultsOutputPath = "Studio13/TestEnvironment/ResultsOutputPath";
        private string m_xActivationKey = "Studio13/TestData/ActivationKey";
        private string m_xTestDataDir = "Studio13/TestEnvironment/TestDataDir";

        public TConfig(string path)
        {
            m_doc = new XmlDocument();

            try
            {
                m_doc.Load(path);
            }
            catch (Exception)
            {
                string msg = "Unable to load xml. Check the path: " + path;
                throw new ApplicationException(msg);
            }
        }

        public string InstalledBuildNumber()
        {
            try
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
            catch (Exception e)
            {
                throw new Exception("Unable to determine installed build number: " + e.InnerException);
            }
        }

        public string ResultsOutputPath(string projectName)
        {
            try
            {
                //read results path root dir from xml
                XmlNode root = m_doc.DocumentElement;
                string sResultsOutputPath = root.SelectSingleNode(m_xResultsOutputPath).InnerText;

                //append path with build number and create dir if needed
                sResultsOutputPath += "\\" + InstalledBuildNumber();
                if (!Directory.Exists(sResultsOutputPath))
                    Directory.CreateDirectory(sResultsOutputPath);

                //append path with projectname
                sResultsOutputPath += "\\" + projectName;

                //create dir if it does not already exists
                if (!Directory.Exists(sResultsOutputPath))
                {
                    Directory.CreateDirectory(sResultsOutputPath);
                }
                else //append dir name with _i to ensure uniqueness
                {
                    string temp = sResultsOutputPath;
                    int i = 0;
                    while (Directory.Exists(temp))
                    {
                        i++;
                        temp = sResultsOutputPath;
                        temp += "_" + i;
                    }

                    sResultsOutputPath = temp;
                    Directory.CreateDirectory(sResultsOutputPath);
                }

                return sResultsOutputPath;
            }
            catch (Exception e)
            {
                throw new Exception("Unable to set RestultsOutputPath: " + e.InnerException);
            }
        }

        public string ActivationKey()
        {
            try
            {
                XmlNode root = m_doc.DocumentElement;
                return root.SelectSingleNode(m_xActivationKey).InnerText;
            }
            catch (Exception e)
            {
                throw new Exception("Unable to determine ActivationKey: " + e.InnerException);
            }
        }

        public string TestDataDir()
        {
            try
            {
                XmlNode root = m_doc.DocumentElement;
                return root.SelectSingleNode(m_xTestDataDir).InnerText;
            }
            catch (Exception e)
            {
                throw new Exception("Unable to determine TestDataDir: " + e.InnerException);
            }
        }
    }
}

