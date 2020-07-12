using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FrameWork.Utility
{
    public class SeleniumUtility
    {
        private ChromeOptions _options;
        private IWebDriver _driver;
        private WebDriverWait _wait;

        public SeleniumUtility(int timeouts = 20, int pollingInterval = 500, string type = null)
        {
            _options = new ChromeOptions();
            // 设置开发者模式启动，该模式下webdriver属性为正常值
            _options.AddExcludedArgument("enable-automation");
            // 提示及禁用扩展插件
            _options.AddArguments("--test-type", "--ignore-certificate-errors");
            // 禁掉信息栏
            _options.AddAdditionalCapability("useAutomationExtension", false);
            // 关闭黑窗
            ChromeDriverService cds = ChromeDriverService.CreateDefaultService();
            cds.HideCommandPromptWindow = true;
            // 窗体设置
            SetOption(type);

            _driver = new ChromeDriver(cds, _options);

            // 等待设置
            _wait = new WebDriverWait(_driver, timeout: TimeSpan.FromSeconds(timeouts))
            {
                PollingInterval = TimeSpan.FromMilliseconds(pollingInterval),
            };
            _wait.IgnoreExceptionTypes(typeof(NoSuchElementException));
        }

        public void DoCommand(string command)
        {
            if (string.IsNullOrWhiteSpace(command)) return;

            if (command.StartsWith("$"))
            {
                switch (command.GetSuCommand())
                {
                    case "$goto":
                        DoUrl(command.GetSuValue());
                        break;
                    case "$sleep":
                        DoSleep(command.GetSuValue());
                        break;
                    case "$scroll":
                        DoScroll(command.GetSuValues());
                        break;
                    default:
                        break;
                }
            }
            else
            {

            }
        }

        /// <summary>
        /// 定向网页
        /// </summary>
        /// <param name="url"></param>
        public void DoUrl(string url)
            => _driver.Navigate().GoToUrl(url);

        /// <summary>
        /// 睡眠
        /// </summary>
        /// <param name="time"></param>
        public void DoSleep(string time)
            => Thread.Sleep(int.Parse(time));

        /// <summary>
        /// 滚动条从上到下
        /// </summary>
        /// <param name="args"></param>
        public void DoScroll(string[] args)
        {
            int cnt = int.Parse(args[0]);
            for (int i = 1; i <= cnt; i++)
            {
                string jsCode = "window.scrollTo({top: document.body.scrollHeight / " + cnt + " * " + i + ", behavior: \"smooth\"});";
                IJavaScriptExecutor js1 = (IJavaScriptExecutor)_driver;
                js1.ExecuteScript(jsCode);
                DoSleep(args[1]);
            }
        }

        /// <summary>
        /// ASCII
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static void Save(ChromeDriver driver, string path, ScreenshotImageFormat type, bool isFull = false)
        {
            if (isFull)
            {
                var filePath = path;

                Dictionary<string, Object> metrics = new Dictionary<string, Object>();
                metrics["width"] = driver.ExecuteScript("return Math.max(window.innerWidth,document.body.scrollWidth,document.documentElement.scrollWidth)");
                metrics["height"] = driver.ExecuteScript("return Math.max(window.innerHeight,document.body.scrollHeight,document.documentElement.scrollHeight)");
                //返回当前显示设备的物理像素分辨率与 CSS 像素分辨率的比率
                metrics["deviceScaleFactor"] = driver.ExecuteScript("return window.devicePixelRatio");
                metrics["mobile"] = driver.ExecuteScript("return typeof window.orientation !== 'undefined'");
                driver.ExecuteChromeCommand("Emulation.setDeviceMetricsOverride", metrics);

                driver.GetScreenshot().SaveAsFile(filePath, ScreenshotImageFormat.Png);
            }
            else
            {
                Screenshot shot = driver.TakeScreenshot();
                shot.SaveAsFile(path, type);
            }
        }

        /// <summary>
        /// 窗体设置
        /// </summary>
        /// <param name="type"></param>
        public void SetOption(string type)
        {
            // 设置窗口大小
            if (type.StartsWith("$size"))
            {
                string[] arr = type.GetSuValues();
                _driver.Manage().Window.Size = new Size(int.Parse(arr[0]), int.Parse(arr[1]));
            }
            // 设置窗口位置
            else if (type.StartsWith("$point"))
            {
                string[] arr = type.GetSuValues();
                _driver.Manage().Window.Position = new Point(int.Parse(arr[0]), int.Parse(arr[1]));
            }
            else
            {
                switch (type)
                {
                    // 全屏窗口(F11)
                    case "$fullscreen":
                        _driver.Manage().Window.FullScreen();
                        break;
                    // 最大化窗口(不会阻挡工具栏)
                    case "$maxscreen":
                        _driver.Manage().Window.Maximize();
                        break;
                    // 最小化窗口
                    case "$minscreen":
                        _driver.Manage().Window.Minimize();
                        break;
                    // 隐藏窗口
                    case "$hidescreen":
                        _options.AddArguments("--headless");
                        break;
                    default:
                        break;
                }
            }
        }

        public IWebElement FindIdElement(string id)
            => _wait.Until(dr => dr.FindElement(By.Id(id)));

        public IWebElement FindNameElement(string name)
            => _wait.Until(dr => dr.FindElement(By.Name(name)));

        public IWebElement FindNameElement(string name)
            => _wait.Until(dr => dr.FindElement(By.Name(name)));
    }

    public static class SeleniumHelper
    {
        public static string GetSuCommand(this string command)
        {
            int index = command.IndexOf(":");
            if (index <= 0)
                return command;
            else
                return command.Substring(0, index);
        }

        public static string GetSuValue(this string command)
            => command.Substring(command.IndexOf(":") + 1);

        public static string[] GetSuValues(this string command)
            => command.Substring(command.IndexOf(":") + 1).Split(',');
    }


    public class ChromeDriverEx : ChromeDriver
    {
        private const string SendChromeCommandWithResult = "sendChromeCommandWithResponse";
        private const string SendChromeCommandWithResultUrlTemplate = "/session/{sessionId}/chromium/send_command_and_get_result";

        public ChromeDriverEx(string chromeDriverDirectory, ChromeOptions options)
            : base(chromeDriverDirectory, options)
        {
            CommandInfo commandInfoToAdd = new CommandInfo(CommandInfo.PostCommand, SendChromeCommandWithResultUrlTemplate);
            this.CommandExecutor.CommandInfoRepository.TryAddCommand(SendChromeCommandWithResult, commandInfoToAdd);
        }

        public ChromeDriverEx(ChromeDriverService service, ChromeOptions options)
            : base(service, options)
        {
            CommandInfo commandInfoToAdd = new CommandInfo(CommandInfo.PostCommand, SendChromeCommandWithResultUrlTemplate);
            this.CommandExecutor.CommandInfoRepository.TryAddCommand(SendChromeCommandWithResult, commandInfoToAdd);
        }

        public Screenshot GetFullPageScreenshot()
        {

            string metricsScript = @"({
width: Math.max(window.innerWidth,document.body.scrollWidth,document.documentElement.scrollWidth)|0,
height: Math.max(window.innerHeight,document.body.scrollHeight,document.documentElement.scrollHeight)|0,
deviceScaleFactor: window.devicePixelRatio || 1,
mobile: typeof window.orientation !== 'undefined'
})";
            Dictionary<string, object> metrics = this.EvaluateDevToolsScript(metricsScript);
            this.ExecuteChromeCommand("Emulation.setDeviceMetricsOverride", metrics);

            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters["format"] = "png";
            parameters["fromSurface"] = true;
            object screenshotObject = this.ExecuteChromeCommandWithResult("Page.captureScreenshot", parameters);
            Dictionary<string, object> screenshotResult = screenshotObject as Dictionary<string, object>;
            string screenshotData = screenshotResult["data"] as string;

            this.ExecuteChromeCommand("Emulation.clearDeviceMetricsOverride", new Dictionary<string, object>());

            Screenshot screenshot = new Screenshot(screenshotData);
            return screenshot;
        }

        public new object ExecuteChromeCommandWithResult(string commandName, Dictionary<string, object> commandParameters)
        {
            if (commandName == null)
            {
                throw new ArgumentNullException("commandName", "commandName must not be null");
            }

            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters["cmd"] = commandName;
            parameters["params"] = commandParameters;
            Response response = this.Execute(SendChromeCommandWithResult, parameters);
            return response.Value;
        }

        private Dictionary<string, object> EvaluateDevToolsScript(string scriptToEvaluate)
        {

            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters["returnByValue"] = true;
            parameters["expression"] = scriptToEvaluate;
            object evaluateResultObject = this.ExecuteChromeCommandWithResult("Runtime.evaluate", parameters);
            Dictionary<string, object> evaluateResultDictionary = evaluateResultObject as Dictionary<string, object>;
            Dictionary<string, object> evaluateResult = evaluateResultDictionary["result"] as Dictionary<string, object>;


            Dictionary<string, object> evaluateValue = evaluateResult["value"] as Dictionary<string, object>;
            return evaluateValue;
        }
    }
}
