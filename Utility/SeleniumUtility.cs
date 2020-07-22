using FrameWork.Models;
using FrameWork.ViewModels.Base;
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
        private Dictionary<string, List<IWebElement>> _listElements;
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

        public void DoCommand(SeleniumEvent _se)
        {
            SeleniumEvent se = new SeleniumEvent()
            {
                No = _se.No,
                Event = _se.Event,
                Back = _se.Back,
            };

            ReplaceArgs(se);

            if (se.Event.StartsWith("$"))
            {
                string suCommand = se.Event.GetSuCommand();
                if (suCommand == null)
                    return;

                switch (suCommand)
                {
                    case "$findnotelement":
                        return;
                    case "$class":
                    case "$id":
                    case "$link":
                    case "$linkpart":
                    case "$xpath":
                    case "$name":
                    case "$cssselect":
                    case "$tag":
                        GetElement(se);
                        return;
                    case "$xpaths":
                        GetElements(se);
                        return;
                    case "$goto":
                        DoUrl(se.Event.GetSuValue());
                        return;
                    case "$sleep":
                        DoSleep(se.Event.GetSuValue());
                        return;
                    case "$outframe":
                        return;
                    case "$click":
                        Click(se);
                        return;
                    case "$hmove":
                        DoHMove(se);
                        return;
                    case "$findtext":
                        DoFindText(se);
                        break;
                    case "$findlink":
                        break;
                    case "$scroll":
                        DoScroll(se.Event.GetSuValues());
                        return;
                    case "$frame":
                        GoToFrame(_elements[se.No]);
                        return;
                    case "$frameout":
                        GoOutFrame();
                        return;
                    case "$back":
                        GoToBack();
                        return;
                    case "$forward":
                        GoToForward();
                        return;
                    case "$refresh":
                        GoToBack();
                        return;
                    case "$move":
                        GoToMove(se);
                        return;
                    case "$down":
                        GoToDown(se);
                        return;
                    default:
                        return;
                }
                return;
            }
            else
            {
                DoKeys(se);
            }
        }

        public void GoToBack()
            => _driver.Navigate().Back();

        public void GoToForward()
            => _driver.Navigate().Forward();

        public void GoToRefresh()
            => _driver.Navigate().Refresh();

        public void GoToDown(SeleniumEvent se)
            => SeleniumHelper.SuDown(Environment.CurrentDirectory + @"\" + SuTime, _elements[se.No].GetAttribute(se.Event.GetSuValue()));

        private void ReplaceArgs(SeleniumEvent se)
        {
            foreach (string key in _args.Keys)
            {
                se.Event = se.Event.Replace(key, _args[key]);
            }
        }

        public void GetElements(SeleniumEvent se)
        {
            string[] arr = se.Event.GetSuValues();

            WebDriverWait wait;
            if (arr.Length == 2)
            {
                wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(int.Parse(arr[1])))
                {
                    PollingInterval = TimeSpan.FromMilliseconds(500.0)
                };
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException));
            }
            else
            {
                wait = _wait;
            }

            string suCommand = se.Event.GetSuCommand();

            List<IWebElement> iwebElement = null;
            switch (suCommand)
            {
                case "$class":
                    //iwebElement = FindClassElement(arr[0], wait);
                    break;
                case "$link":
                    //iwebElement = FindLinkElement(arr[0], wait);
                    break;
                case "$linkpart":
                    //iwebElement = FindLinkPartElement(arr[0], wait);
                    break;
                case "$xpaths":
                    iwebElement = FindXPathElements(arr[0], wait);
                    break;
                case "$cssselect":
                    //iwebElement = FindCsElement(arr[0], wait);
                    break;
                case "$tag":
                    break;
                default:
                    return;
            }

            if (iwebElement != null && iwebElement.Count > 0)
                _listElements.Add(se.No, iwebElement);
            else if (wait == _wait)
            {
                throw new NoSuchElementException($"元素不存在[{se.Event}]");
            }
        }

        public void GetElement(SeleniumEvent se)
        {
            string[] arr = se.Event.GetSuValues();

            WebDriverWait wait;
            if (arr.Length == 2)
            {
                wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(int.Parse(arr[1])))
                {
                    PollingInterval = TimeSpan.FromMilliseconds(500.0)
                };
                wait.IgnoreExceptionTypes(typeof(NoSuchElementException));
            }
            else
            {
                wait = _wait;
            }

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
                        return;
                }
            }
            catch (WebDriverTimeoutException)
            {
            }

            if (iwebElement != null)
                _elements.Add(se.No, iwebElement);
            else if (wait == _wait)
            {
                throw new NoSuchElementException($"元素不存在[{se.Event}]");
            }
        }

        public void DoKeys(SeleniumEvent se)
        {
            _elements[se.No].SendKeys(se.Event);
        }

        public void ClearElements()
        {
            _elements = new Dictionary<string, IWebElement>();
            _listElements = new Dictionary<string, List<IWebElement>>();
        }

        public void DoUrl(string url)
        {
            _driver.Navigate().GoToUrl(url);
        }

        public void DoSleep(string time)
        {
            Thread.Sleep(int.Parse(time));
        }

        public void DoScroll(string[] args)
        {
            int num = int.Parse(args[0]);
            for (int index = 1; index <= num; ++index)
            {
                ((IJavaScriptExecutor)_driver).ExecuteScript("window.scrollTo({top: document.body.scrollHeight / " + num.ToString() + " * " + index.ToString() + ", behavior: \"smooth\"});", Array.Empty<object>());
                DoSleep(args[1]);
            }
        }

        public void DoFindText(SeleniumEvent se)
        {
            string[] args = se.Event.GetSuValues();

            List<IWebElement> elements = _wait.Until(d => _elements[args[0]].FindElements(By.TagName(args[1]))).ToList();
            foreach (IWebElement temp in elements)
                if (temp.Text.Contains(args[2]))
                {
                    _elements.Add(se.No, temp);
                }
        }

        public void DoHMove(SeleniumEvent se)
        {
            if (!_elements.ContainsKey(se.No)) return;

            string[] suValues = se.Event.GetSuValues();
            Actions actions = new Actions(_driver);
            actions.ClickAndHold(_elements[se.No]).MoveByOffset(int.Parse(suValues[0]), int.Parse(suValues[1])).Perform();
            actions.Release(_elements[se.No]);
        }

        public void Click(SeleniumEvent se)
        {
            if (!_elements.ContainsKey(se.No)) return;

            IWebElement element = _elements[se.No];

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
            => new Actions(_driver).MoveToElement(_elements[se.No]).Perform();

        public void GoToFrame(IWebElement element)
        {
            _driver.SwitchTo().Frame(element);
        }

        public void GoOutFrame()
        {
            _driver.SwitchTo().DefaultContent();
        }

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

        public IWebElement FindIdElement(string[] arrs, WebDriverWait wait)
            => DoFindEvent(wait, dr => _elements[arrs[1]].FindElement(By.Id(arrs[2])));

        public IWebElement FindNameElement(string name, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.Name(name)));

        public IWebElement FindNameElement(string[] arrs, WebDriverWait wait)
            => DoFindEvent(wait, dr => _elements[arrs[1]].FindElement(By.Name(arrs[2])));

        public IWebElement FindClassElement(string className, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.ClassName(className)));

        public IWebElement FindClassElement(string[] arrs, WebDriverWait wait)
            => DoFindEvent(wait, dr => _elements[arrs[1]].FindElement(By.ClassName(arrs[2])));

        public IWebElement FindLinkElement(string link, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.LinkText(link)));

        public IWebElement FindLinkElement(string[] arrs, WebDriverWait wait)
            => DoFindEvent(wait, dr => _elements[arrs[1]].FindElement(By.LinkText(arrs[2])));

        public IWebElement FindCsElement(string select, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.CssSelector(select)));

        public IWebElement FindCsElement(string[] arrs, WebDriverWait wait)
            => DoFindEvent(wait, dr => _elements[arrs[1]].FindElement(By.CssSelector(arrs[2])));

        public IWebElement FindLinkPartElement(string link, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.PartialLinkText(link)));

        public IWebElement FindLinkPartElement(string[] arrs, WebDriverWait wait)
            => DoFindEvent(wait, dr => _elements[arrs[1]].FindElement(By.PartialLinkText(arrs[2])));

        public IWebElement FindXPathElement(string path, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.XPath(path)));

        public IWebElement FindXPathElement(string[] arrs, WebDriverWait wait)
            => DoFindEvent(wait, dr => _elements[arrs[1]].FindElement(By.XPath(arrs[2])));

        public IWebElement FindTagElement(string path, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.TagName(path)));

        public IWebElement FindTagElement(string[] arrs, WebDriverWait wait)
            => DoFindEvent(wait, dr => _elements[arrs[1]].FindElement(By.TagName(arrs[2])));

        private IWebElement DoFindEvent(WebDriverWait wait, Func<IWebDriver, IWebElement> func)
        => wait.Until(dr =>
            {
                if (FlgStop)
                    throw new Exception("Stop");
                return func(dr);
            });

        public List<IWebElement> FindXPathElements(string path, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElements(By.XPath(path)));

        private List<IWebElement> DoFindEvent(WebDriverWait wait, Func<IWebDriver, IEnumerable<IWebElement>> func)
            => wait.Until(dr =>
            {
                if (FlgStop)
                    throw new Exception("Stop");
                return func(dr);
            }).ToList();

        public void Dispose()
        {
            _driver.Dispose();
        }
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
            return command.Substring(command.IndexOf(":") + 1).Split(';');
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
