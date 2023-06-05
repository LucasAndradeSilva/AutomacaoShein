using Microsoft.Win32;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Helpers
{
    public class SeleniumHelper
    {
        public static ChromeDriver StartDriver()
        {
            string ChromeDriverPath = Environment.CurrentDirectory;
            string ChromDriverVersion = GetChromDriverVersion(ChromeDriverPath);
            string ChromeWebVersion = GetWebChromeVersion();

            if (ChromDriverVersion != ChromeWebVersion)
            {
                var urlToDownload = GetURLToDownload(ChromeWebVersion);
                KillAllChromeDriverProcesses();
                DownloadNewVersionOfChrome(urlToDownload, ChromeDriverPath);
                string extract = ExtructZip(ChromeDriverPath);
            }

            var driver = new ChromeDriver(Path.Combine(Environment.CurrentDirectory, "chromedriver.exe"));
            return driver;
        }
        public static IWebElement WaitElement(ChromeDriver driver, By by)
        {
            var tentativas = 0;
            IWebElement webElement = default;

            while (tentativas <= 20)
            {
                try
                {
                    webElement = driver.FindElement(by);
                    break;
                }
                catch (Exception ex)
                {
                    tentativas++;
                }
                finally
                {
                    Thread.Sleep(500);
                }
            }

            return webElement;
        }

        public static AppiumWebElement WaitElement(AndroidDriver<AppiumWebElement> driver, By by)
        {
            var tentativas = 0;
            AppiumWebElement webElement = default;

            while (tentativas <= 20)
            {
                try
                {
                    webElement = driver.FindElement(by);
                    break;
                }
                catch (Exception ex)
                {
                    tentativas++;
                }
                finally
                {
                    Thread.Sleep(500);
                }
            }

            return webElement;
        }

        public static AppiumWebElement WaitForeverElement(AndroidDriver<AppiumWebElement> driver, By by)
        {            
            AppiumWebElement webElement = default;

            while (true)
            {
                try
                {
                    webElement = driver.FindElement(by);
                    break;
                }
                catch (Exception ex)
                {
                    
                }
                finally
                {
                    Thread.Sleep(500);
                }
            }

            return webElement;
        }


        public static (bool, IWebElement) TryFindElement(ChromeDriver driver, By by)
        {
            try
            {
                var element = driver.FindElement(by);
                return (true, element);
            }
            catch (NoSuchElementException ex)
            {                
                return (false, null);
            }            
        }

        public static bool IsElementVisible(IWebElement element)
        {
            if (element is null)            
                return false;
            
            return element.Displayed && element.Enabled;
        }
        private static string GetChromDriverVersion(string ChromeDriverePath)
        {
            string driverversion = "";
            if (File.Exists(ChromeDriverePath + "\\chromedriver.exe"))
            {
                IWebDriver driver = new ChromeDriver(Path.Combine(Environment.CurrentDirectory, "chromedriver.exe"));
                ICapabilities capabilities = ((ChromeDriver)driver).Capabilities;
                driverversion = ((capabilities.GetCapability("chrome") as Dictionary<string, object>)["chromedriverVersion"]).ToString().Split(' ').First();
                driver.Dispose();
            }
            else
            {
                Console.WriteLine("ChromeDriver.exe missing !!");
            }


            return driverversion;
        }
        private static string GetChromeWebPath()
        {
            //Path originates from here: https://chromedriver.chromium.org/downloads/version-selection            
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\chrome.exe"))
            {
                if (key != null)
                {
                    Object o = key.GetValue("");
                    if (!String.IsNullOrEmpty(o.ToString()))
                    {
                        return o.ToString();
                    }
                    else
                    {
                        throw new ArgumentException("Unable to get version because chrome registry value was null");
                    }
                }
                else
                {
                    throw new ArgumentException("Unable to get version because chrome registry path was null");
                }
            }
        }
        private static string GetWebChromeVersion()
        {
            string productVersionPath = GetChromeWebPath();
            if (String.IsNullOrEmpty(productVersionPath))
            {
                throw new ArgumentException("Unable to get version because path is empty");
            }

            if (!File.Exists(productVersionPath))
            {
                throw new FileNotFoundException("Unable to get version because path specifies a file that does not exists");
            }

            var versionInfo = FileVersionInfo.GetVersionInfo(productVersionPath);
            if (versionInfo != null && !String.IsNullOrEmpty(versionInfo.FileVersion))
            {
                return versionInfo.FileVersion;
            }
            else
            {
                throw new ArgumentException("Unable to get version from path because the version is either null or empty: " + productVersionPath);
            }
        }
        private static string GetURLToDownload(string version)
        {
            if (String.IsNullOrEmpty(version))
            {
                throw new ArgumentException("Unable to get url because version is empty");
            }

            //URL's originates from here: https://chromedriver.chromium.org/downloads/version-selection
            string html = string.Empty;
            string urlToPathLocation = @"https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" + String.Join(".", version.Split('.').Take(3));

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlToPathLocation);
            request.AutomaticDecompression = DecompressionMethods.GZip;

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream))
            {
                html = reader.ReadToEnd();
            }

            if (String.IsNullOrEmpty(html))
            {
                throw new WebException("Unable to get version path from website");
            }

            return "https://chromedriver.storage.googleapis.com/" + html + "/chromedriver_win32.zip";
        }
        private static void KillAllChromeDriverProcesses()
        {
            var processes = Process.GetProcessesByName("chromedriver");
            foreach (var process in processes)
            {
                try
                {
                    process.Kill();
                }
                catch
                {

                }
            }
        }
        private static void DownloadNewVersionOfChrome(string urlToDownload, string ChromDriverPath)
        {
            if (String.IsNullOrEmpty(urlToDownload))
            {
                throw new ArgumentException("Unable to get url because urlToDownload is empty");
            }
            using (var client = new WebClient())
            {
                if (File.Exists(ChromDriverPath + "\\chromedriver.zip"))
                {
                    File.Delete(ChromDriverPath + "\\chromedriver.zip");
                }

                client.DownloadFile(urlToDownload, "chromedriver.zip");

                if (File.Exists(ChromDriverPath + "\\chromedriver.zip") && File.Exists(ChromDriverPath + "\\chromedriver.exe"))
                {
                    File.Delete(ChromDriverPath + "\\chromedriver.exe");
                }

            }
        }
        private static string ExtructZip(string ChromeDriverPath)
        {
            string errorMsg = "";
            string zipFileName = "";
            try
            {
                zipFileName = Path.Combine(Environment.CurrentDirectory, "chromedriver.zip");
                string orgFileName = "";                
                using (var zip = ZipFile.OpenRead(zipFileName))
                {
                    foreach (var e in zip.Entries)
                    {
                        e.ExtractToFile(Path.Combine(Environment.CurrentDirectory, "chromedriver.exe"), true);
                        orgFileName = e.Name;
                    }
                }
                File.Delete(zipFileName);
                errorMsg = "Done";
            }
            catch (Exception ex)
            {
                errorMsg = "ExtructZip  file " + zipFileName.Split('\\').Last() + " - " + ex.Message;
            }
            return errorMsg;
        }
    }
}
