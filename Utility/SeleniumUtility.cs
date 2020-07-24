using FrameWork.Models;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;

namespace FrameWork.Utility
{
    public class SeleniumUtility : IDisposable
    {
        public string SuTime { get; set; } = DateTime.Now.ToString("yyyyMMddHHmmss");
        private ChromeOptions _options;
        private IWebDriver _driver;
        private WebDriverWait _wait;
        private Dictionary<string, string> _args = new Dictionary<string, string>();
        private Dictionary<string, IWebElement> _elements;

        public bool FlgStop { get; set; } = false;

        public SeleniumUtility(int timeouts = 20, int pollingInterval = 500, string type = "")
        {
            _options = new ChromeOptions();

            _options.AddExcludedArgument("enable-automation");
            _options.AddArguments(new string[2] { "--test-type", "--ignore-certificate-errors" });
            _options.AddArgument("lang=zh_CN.UTF-8");

            _options.AddAdditionalCapability("useAutomationExtension", false);

            ChromeDriverService defaultService = ChromeDriverService.CreateDefaultService();
            defaultService.HideCommandPromptWindow = true;

            SetOption(type);
            _driver = new ChromeDriver(defaultService, _options);
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(timeouts))
            {
                PollingInterval = TimeSpan.FromMilliseconds(pollingInterval)
            };

            _wait.IgnoreExceptionTypes(typeof(NoSuchElementException));

            ClearElements();
        }

        public void AddArg(string key, string value)
            => _args.Add(key, value);

        public void SetArg(string key, string value)
            => _args[key] = value;

        public void DoEvents(List<SeleniumEvent> events,ref string continuEvent)
        {
            string gotNo = null;
            do
            {
                foreach (SeleniumEvent item in events)
                {
                    if (FlgStop) throw new Exception("Stop");

                    if (!string.IsNullOrWhiteSpace(continuEvent))
                    {
                        if (item.Back != continuEvent)
                            continue;
                        else
                            continuEvent = null;
                    }

                    if (!string.IsNullOrWhiteSpace(gotNo))
                    {
                        if (item.No == gotNo)
                            gotNo = null;
                        else
                            continue;
                    }

                    SeleniumEvent temp = new SeleniumEvent()
                    {
                        No = item.No,
                        Key = item.Key,
                        Event = item.Event,
                        Back = item.Back,
                    };

                    ReplaceArgs(temp);

                    gotNo = DoCommand(temp);

                    LogUtility.WriteInfo($"【Event执行成功】-【{item.Key}-{item.Event}】");

                    if (!string.IsNullOrWhiteSpace(gotNo)) break;
                }
            } while (!string.IsNullOrEmpty(gotNo));
        }

        public string DoCommand(SeleniumEvent se)
        {
            if (se.Event.StartsWith("$"))
            {
                string suCommand = se.Event.GetSuCommand();

                switch (suCommand)
                {
                    case "$findnotelement":
                        break;
                    case "$class":
                    case "$id":
                    case "$link":
                    case "$linkpart":
                    case "$xpath":
                    case "$name":
                    case "$cssselect":
                    case "$tag":
                        return GetElement(se);
                    case "$goto":
                        DoUrl(se);
                        break;
                    case "$sleep":
                        DoSleep(se);
                        break;
                    case "$outframe":
                        break;
                    case "$click":
                        Click(se);
                        break;
                    case "$hmove":
                        DoHMove(se);
                        break;
                    case "$scroll":
                        DoScroll(se);
                        break;
                    case "$frame":
                        GoToFrame(se);
                        break;
                    case "$frameout":
                        GoOutFrame();
                        break;
                    case "$back":
                        GoToBack();
                        break;
                    case "$forward":
                        GoToForward();
                        break;
                    case "$refresh":
                        GoToRefresh();
                        break;
                    case "$mov nulle":
                        GoToMove(se);
                        break;
                    case "$down":
                        GoToDown(se);
                        break;
                    default:
                        break;
                }
            }
            else
            {
                DoKeys(se);
                return null;
            }
            return null;
        }

        public void GoToBack()
            => _driver.Navigate().Back();

        public void GoToForward()
            => _driver.Navigate().Forward();

        public void GoToRefresh()
            => _driver.Navigate().Refresh();

        public void GoToDown(SeleniumEvent se)
            => SeleniumHelper.SuDown(Environment.CurrentDirectory + @"\" + SuTime, _elements[se.Key].GetAttribute(se.Event.GetSuValue()));

        private void ReplaceArgs(SeleniumEvent se)
            => _args.Keys.ToList().ForEach(k => se.Event = se.Event.Replace(k, _args[k]));

        public string GetElement(SeleniumEvent se)
        {
            string[] arr = se.Event.GetSuValues();

            WebDriverWait wait;
            if (arr.Length >= 2)
            {
                wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(int.Parse(arr[1])))
                {
                    PollingInterval = TimeSpan.FromMilliseconds(500.0)
                };
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException));
            }
            else
                wait = _wait;

            string suCommand = se.Event.GetSuCommand();

            IWebElement iwebElement = null;
            try
            {
                switch (suCommand)
                {
                    case "$id":
                        iwebElement = FindIdElement(arr[0], wait);
                        break;
                    case "$class":
                        iwebElement = FindClassElement(arr[0], wait);
                        break;
                    case "$link":
                        iwebElement = FindLinkElement(arr[0], wait);
                        break;
                    case "$linkpart":
                        iwebElement = FindLinkPartElement(arr[0], wait);
                        break;
                    case "$xpath":
                        iwebElement = FindXPathElement(arr[0], wait);
                        break;
                    case "$name":
                        iwebElement = FindNameElement(arr[0], wait);
                        break;
                    case "$cssselect":
                        iwebElement = FindCsElement(arr[0], wait);
                        break;
                    case "$tag":
                        iwebElement = FindTagElement(arr[0], wait);
                        break;
                    default:
                        break;
                }
            }
            catch (WebDriverTimeoutException) { }

            if (iwebElement != null)
            {
                if (_elements.ContainsKey(se.Key))
                    _elements[se.Key] = iwebElement;
                else
                    _elements.Add(se.Key, iwebElement);
            }
            else if (wait == _wait)
            {
                FlgStop = true;
                throw new NoSuchElementException($"查找失败[{se.Event}]");
            }

            if (arr.Length == 3 && iwebElement != null)
                return arr[2];
            if (arr.Length == 4 && iwebElement == null)
                return arr[3];
            else
                return null;
        }

        public void DoKeys(SeleniumEvent se)
            => _elements[se.Key].SendKeys(se.Event);

        public void ClearElements()
            => _elements = new Dictionary<string, IWebElement>();

        public void DoUrl(SeleniumEvent se)
            => _driver.Navigate().GoToUrl(se.Event.GetSuValue());

        public void DoSleep(SeleniumEvent se)
            => DoSleep(se.Event.GetSuValue());

        public void DoSleep(string time)
            => Thread.Sleep(int.Parse(time));

        public void DoScroll(SeleniumEvent se)
        {
            string[] args = se.Event.GetSuValues();
            int num = int.Parse(args[0]);
            for (int index = 1; index <= num; ++index)
            {
                ((IJavaScriptExecutor)_driver).ExecuteScript("window.scrollTo({top: document.body.scrollHeight / " + num.ToString() + " * " + index.ToString() + ", behavior: \"smooth\"});", Array.Empty<object>());
                DoSleep(args[1]);
            }
        }

        public void DoHMove(SeleniumEvent se)
        {
            if (!_elements.ContainsKey(se.Key)) return;

            string[] suValues = se.Event.GetSuValues();
            Actions actions = new Actions(_driver);
            actions.ClickAndHold(_elements[se.Key]).MoveByOffset(int.Parse(suValues[0]), int.Parse(suValues[1])).Perform();
            actions.Release(_elements[se.Key]);
        }

        public void Click(SeleniumEvent se)
        {
            if (!_elements.ContainsKey(se.Key)) return;

            IWebElement element = _elements[se.Key];

            if (se.Event.IndexOf(":") <= 0)
            {
                element.Click();
            }
            else
            {
                switch (se.Event.GetSuValue())
                {
                    case "1":
                        element.Click();
                        break;
                    case "2":
                        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", new object[1] { element });
                        break;
                    case "3":
                        new Actions(_driver).MoveToElement(element).Click().Perform();
                        break;
                }
            }
        }

        public void GoToMove(SeleniumEvent se)
            => new Actions(_driver).MoveToElement(_elements[se.Key]).Perform();

        public void GoToFrame(SeleniumEvent se)
        {
            if (!string.IsNullOrWhiteSpace(se.Key))
                _driver.SwitchTo().Frame(_elements[se.Key]);
            else
                _driver.SwitchTo().Frame(int.Parse(se.Event.GetSuValue()));
        }

        public void GoOutFrame()
            => _driver.SwitchTo().DefaultContent();

        public void SavePic(string path, ScreenshotImageFormat type, bool isFull = false)
        {
            if (isFull)
            {
                string str = path;
                ((ChromeDriver)_driver).ExecuteChromeCommand("Emulation.setDeviceMetricsOverride", new Dictionary<string, object>()
                {
                    ["width"] = ((RemoteWebDriver)_driver).ExecuteScript("return Math.max(window.innerWidth,document.body.scrollWidth,document.documentElement.scrollWidth)", Array.Empty<object>()),
                    ["height"] = ((RemoteWebDriver)_driver).ExecuteScript("return Math.max(window.innerHeight,document.body.scrollHeight,document.documentElement.scrollHeight)", Array.Empty<object>()),
                    ["deviceScaleFactor"] = ((RemoteWebDriver)_driver).ExecuteScript("return window.devicePixelRatio", Array.Empty<object>()),
                    ["mobile"] = ((RemoteWebDriver)_driver).ExecuteScript("return typeof window.orientation !== 'undefined'", Array.Empty<object>())
                });
                ((RemoteWebDriver)_driver).GetScreenshot().SaveAsFile(str, ScreenshotImageFormat.Png);
            }
            else
                WebDriverExtensions.TakeScreenshot(_driver).SaveAsFile(path, type);
        }

        public void SetOption(string type = "")
        {
            if (type.StartsWith("$size"))
            {
                string[] suValues = type.GetSuValues();
                _driver.Manage().Window.Size = new Size(int.Parse(suValues[0]), int.Parse(suValues[1]));
            }
            else if (type.StartsWith("$point"))
            {
                string[] suValues = type.GetSuValues();
                _driver.Manage().Window.Position = new Point(int.Parse(suValues[0]), int.Parse(suValues[1]));
            }
            else
            {
                switch (type)
                {
                    case "$fullscreen":
                        _driver.Manage().Window.FullScreen();
                        break;
                    case "$maxscreen":
                        _driver.Manage().Window.Maximize();
                        break;
                    case "$minscreen":
                        _driver.Manage().Window.Minimize();
                        break;
                    case "$hidescreen":
                        _options.AddArguments(new string[1] { "--headless" });
                        break;
                }
            }
        }

        public IWebElement FindIdElement(string id, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.Id(id)));

        public IWebElement FindNameElement(string name, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.Name(name)));

        public IWebElement FindClassElement(string className, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.ClassName(className)));

        public IWebElement FindLinkElement(string link, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.LinkText(link)));

        public IWebElement FindCsElement(string select, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.CssSelector(select)));

        public IWebElement FindLinkPartElement(string link, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.PartialLinkText(link)));

        public IWebElement FindXPathElement(string path, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.XPath(path)));

        public IWebElement FindTagElement(string path, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.TagName(path)));

        private IWebElement DoFindEvent(WebDriverWait wait, Func<IWebDriver, IWebElement> func)
        => wait.Until(dr =>
            {
                if (FlgStop) throw new Exception("Stop");
                return func(dr);
            });

        public void Dispose()
            => _driver.Dispose();
    }

    public static class SeleniumHelper
    {

        public static string GetSuCommand(this string command)
        {
            int length = command.IndexOf(":");
            return length <= 0 ? command : command.Substring(0, length);
        }

        public static string GetSuValue(this string command)
        {
            return command.Substring(command.IndexOf(":") + 1);
        }

        public static string[] GetSuValues(this string command)
        {
            return command.Substring(command.IndexOf(":") + 1).Split('$');
        }

        public static void SuDown(string path, string url)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            // 设置参数
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            //发送请求并获取相应回应数据
            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            //直到request.GetResponse()程序才开始向目标网页发送Post请求
            Stream responseStream = response.GetResponseStream();

            //创建本地文件写入流
            Stream stream = new FileStream(path + $@"\{DateTime.Now:yyyyMMddHHmmss}.xlsx", FileMode.Create);

            byte[] bArr = new byte[1024];
            int size = responseStream.Read(bArr, 0, (int)bArr.Length);
            while (size > 0)
            {
                stream.Write(bArr, 0, size);
                size = responseStream.Read(bArr, 0, (int)bArr.Length);
            }
            stream.Close();
            responseStream.Close();
        }
    }
}
