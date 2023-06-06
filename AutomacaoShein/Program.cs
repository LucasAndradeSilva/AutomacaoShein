using Castle.Core.Configuration;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Android;
using OpenQA.Selenium.Appium.Enums;
using System.Diagnostics;
using Helpers;
using System.Runtime.InteropServices;

namespace AutomacaoShein
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.Title = "Pontos Shein";

            // Abre emulador
            var processEmulador = new Process();
            processEmulador.StartInfo.FileName = "C:\\Android\\android-sdk\\emulator\\emulator.exe"; // Caminho completo para o executável do emulador
            processEmulador.StartInfo.Arguments = "-avd AutomacaoCelular"; // Argumentos para iniciar o emulador com o AVD desejado
            processEmulador.StartInfo.CreateNoWindow = false;
            processEmulador.Start();

            Thread.Sleep(5000);

            // Appium server
            var processServer = new Process();
            processServer.StartInfo.FileName = "cmd.exe"; // Caminho completo para o executável do emulador
            processServer.StartInfo.Arguments = "/C appium"; // Argumentos para iniciar o emulador com o AVD desejado
            processServer.StartInfo.CreateNoWindow = true;                        
            processServer.Start();

            Thread.Sleep(5000);

            //Appium
            var appiumOptions = new AppiumOptions();
            appiumOptions.AddAdditionalCapability(MobileCapabilityType.PlatformName, "Android");
            appiumOptions.AddAdditionalCapability(MobileCapabilityType.NewCommandTimeout, 3600);
            appiumOptions.AddAdditionalCapability(MobileCapabilityType.DeviceName, "AutomacaoCelular");
            appiumOptions.AddAdditionalCapability("appPackage", "com.zzkko");
            appiumOptions.AddAdditionalCapability("appActivity", "com.shein.welcome.WelcomeActivity");

            var androidDriver = new AndroidDriver<AppiumWebElement>(new Uri("http://127.0.0.1:4723/wd/hub"), appiumOptions);

            Thread.Sleep(8000);

            // Clique botão "Go Shopping"
            androidDriver.FindElement(By.Id("com.zzkko:id/n2")).Click();

            Thread.Sleep(2000);

            // Fecha modal cupom
            try
            {
                androidDriver.FindElement(By.Id("com.zzkko:id/ay7")).Click();
                Thread.Sleep(300);
            }
            catch (Exception ex)
            {
            }

            // Clique botão "Menu"
            androidDriver.FindElement(By.Id("com.zzkko:id/bht")).Click();
            Thread.Sleep(300);

            // Clique botão "Sing in"
            androidDriver.FindElement(By.Id("com.zzkko:id/bix")).Click();
            Thread.Sleep(3000);

            // Clique botão "Google"
            androidDriver.FindElement(By.Id("com.zzkko:id/ai0")).Click();
            Thread.Sleep(7000);

            // Escolhe conta
            androidDriver.FindElement(By.XPath("/hierarchy/android.widget.FrameLayout/android.widget.LinearLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.FrameLayout/android.widget.LinearLayout/android.support.v7.widget.RecyclerView/android.widget.LinearLayout[2]")).Click();
            Thread.Sleep(6000);

            // Clica em pontos            
            SeleniumHelper.WaitForeverElement(androidDriver, By.Id("com.zzkko:id/aaz")).Click();
            Thread.Sleep(3000);

            // Clica em "Check in"
            androidDriver.FindElement(By.Id("com.zzkko:id/bbu")).Click();
            Thread.Sleep(10000);
            
            processServer.Kill();
            processServer.WaitForExit();

            processEmulador.Kill();
            processEmulador.WaitForExit();

            GC.Collect();            

            Environment.Exit(0);
        }
    }
}