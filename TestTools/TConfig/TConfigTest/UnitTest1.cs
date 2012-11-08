using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TConfig;

namespace TConfigTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            string path = "C:\\git\\desktoptest\\TestTools\\TConfig\\TConfig\\tconfig.xml";

            TConfig.TConfig config = new TConfig.TConfig(path);

            bool f = config.SendMail();

            string sBuild = config.InstalledBuildNumber();

            sBuild += ":";
        }
    }
}
