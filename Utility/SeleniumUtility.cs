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
        private ChromeOptions _options;
        private IWebDriver _driver;
        private WebDriverWait _wait;
        // Event 控件
        public Dictionary<string, IWebElement> _elements { get; set; }
        // 框 控件
        private List<SeleniumEvent> _logs;
        // 当前Order
        public SeleniumOrder _order;
        // 图片ID
        private int _id;

        public bool FlgStop { get; set; } = false;

        public SeleniumUtility(SeleniumScriptModel model)
        {
            // 设置
            _options = new ChromeOptions();
            _options.AddExcludedArgument("enable-automation");
            _options.AddArguments(new string[2] { "--test-type", "--ignore-certificate-errors" });
            //_options.AddArgument("lang=zh_CN.UTF-8");
            //_options.AddArgument(@"--user-data-dir=C:\Users\dhc\AppData\Local\Google\Chrome\User Data\Default");
            //_options.AddArgument(@"--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36");
            //_options.AddArgument("--flag-switches-begin");
            //_options.AddArgument("--flag-switches-end");
            _options.AddArgument("--enable-audio-service-sandbox");
            _options.AddAdditionalCapability("useAutomationExtension", false);

            // 黑框非表示
            ChromeDriverService defaultService = ChromeDriverService.CreateDefaultService();
            defaultService.HideCommandPromptWindow = true;

            // Set
            _driver = new ChromeDriver(defaultService, _options);
            // Wait
            _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(model.OutTime))
            {
                PollingInterval = TimeSpan.FromMilliseconds(model.ReTry)
            };
            _wait.IgnoreExceptionTypes(typeof(NoSuchElementException));

            // Out Path
            if (!string.IsNullOrWhiteSpace(model.OutPath) && !Directory.Exists(model.OutPath))
                Directory.CreateDirectory(model.OutPath);
        }

        //public void DoEvents(string eve, List<SeleniumEvent> events,ref string continuEvent)
        public void DoEvents(SeleniumScriptModel model)
        {
            foreach (SeleniumOrder item in model.Orders)
            {
                LogUtility.WriteInfo($"【Order开始】-【{item.File}-{item.Case}-{item.Sid}-{item.Back}】");

                // 初期
                _order = item;
                _elements = new Dictionary<string, IWebElement>();
                _logs = new List<SeleniumEvent>();
                _id = 1;
                string goNo = null;

                for (int i = 0; i < model.Events[_order.Sid][_order.Case].Count();)
                {
                    if (FlgStop) throw new Exception("Stop");

                    SeleniumEvent temp = model.Events[_order.Sid][_order.Case][i];

                    // Skip
                    if (string.IsNullOrWhiteSpace(temp.Cmd) && string.IsNullOrWhiteSpace(temp.Value))
                        continue;

                    SeleniumEvent se = new SeleniumEvent()
                    {
                        No = temp.No,
                        Id = temp.Id,
                        Cmd = temp.Cmd,
                        Value = temp.Value,
                        Range = temp.Range,
                    };

                    // Replace Args
                    if (se.Value.Contains("$"))
                        model.Args.Keys.ToList().ForEach(k => se.Value = se.Value.Replace(k, model.Args[k]));

                    // Do Event
                    goNo = DoCommand(model, se);

                    // Add Range
                    if (!string.IsNullOrWhiteSpace(se.Range))
                        _logs.Add(se);

                    // Log
                    LogUtility.WriteInfo($"【Event执行成功】-【{se.Id}-{se.Value}】");

                    // Next
                    if (string.IsNullOrWhiteSpace(goNo))
                        i++;
                    else
                        for (int j = 0; j < model.Events[_order.Sid][_order.Case].Count(); j++)
                            if (goNo == model.Events[_order.Sid][_order.Case][j].No)
                            {
                                i = j;
                                goNo = null;
                                break;
                            }
                }

                LogUtility.WriteInfo($"【Order终了】-【{item.File}-{item.Case}-{item.Sid}-{item.Back}】");
            }
        }

        public string DoCommand(SeleniumScriptModel model, SeleniumEvent se)
        {
            if (se.Cmd.StartsWith("$"))
            {
                switch (se.Cmd)
                {
                    case "$pic":
                        GetPic(model.OutPath);
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
                    case "$win":
                        DoWin(se);
                        break;
                    case "$sleep":
                        DoSleep(se);
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
                    case "$move":
                        GoToMove(se);
                        break;
                    case "$down":
                        GoToDown(model, se);
                        break;
                    case "$maxscreen":
                        WinMax();
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

        private void WinMax()
            => _driver.Manage().Window.Maximize();

        public void GetPic(string basePath)
        {
            // Path Check
            string path = basePath + @"\pics\" + _order.Case;
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            // Range Add
            List<string> lines = new List<string>();
            foreach (SeleniumEvent item in _logs)
            {
                if (_elements.ContainsKey(item.Id))
                    lines.Add(
                        _order.Case + ";" +
                        _id + ";" +
                        _elements[item.Id].Location.X + ";" +
                        _elements[item.Id].Location.Y + ";" +
                        _elements[item.Id].Size.Width + ";" +
                        _elements[item.Id].Size.Height + ";" +
                        item.Range
                    );
            }

            // Add pic
            SavePic(path + @"\" + _id++ + ".jpg");

            // Add log
            if (lines.Count > 0)
                File.AppendAllLines(basePath + @"\pics\log.txt", lines);

            // Clear
            _logs = new List<SeleniumEvent>();
        }

        public void GoToBack()
            => _driver.Navigate().Back();

        public void GoToForward()
            => _driver.Navigate().Forward();

        public void GoToRefresh()
            => _driver.Navigate().Refresh();

        public void GoToDown(SeleniumScriptModel model, SeleniumEvent se)
            => SeleniumHelper.SuDown(model.OutPath + @"\down", _elements[se.Id].GetAttribute(se.Value));

        public string GetElement(SeleniumEvent se)
        {
            string[] arr = se.Value.GetSuValues();

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

            IWebElement iwebElement = null;
            try
            {
                switch (se.Cmd)
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

            // Add Element
            if (iwebElement != null)
            {
                if (_elements.ContainsKey(se.Id))
                    _elements[se.Id] = iwebElement;
                else
                    _elements.Add(se.Id, iwebElement);
            }
            else if (wait == _wait)
            {
                throw new NoSuchElementException($"查找失败[{se.Cmd}]-[{se.Value}]");
            }

            // value;outtime;find go to;not find go to
            if (arr.Length >= 3 && iwebElement != null)
                return arr[2];
            if (arr.Length >= 4 && iwebElement == null)
                return arr[3];
            else
                return null;
        }

        public IWebElement GetElement(string type, string value)
        {
            switch (type)
            {
                case "$id":
                    return FindIdElement(value);
                case "$class":
                    return FindClassElement(value);
                case "$link":
                    return FindLinkElement(value);
                case "$linkpart":
                    return FindLinkPartElement(value);
                case "$xpath":
                    return FindXPathElement(value);
                case "$name":
                    return FindNameElement(value);
                case "$cssselect":
                    return FindCsElement(value);
                case "$tag":
                    return FindTagElement(value);
                    break;
                default:
                    return null;
            }
        }

        public void DoKeys(SeleniumEvent se)
            => _elements[se.Id].SendKeys(se.Value);

        public void DoUrl(SeleniumEvent se)
            => _driver.Navigate().GoToUrl(se.Value);

        public void DoWin(SeleniumEvent se, int cnt = 1)
        {
            foreach (string winHandle in _driver.WindowHandles)
            {
                _driver.SwitchTo().Window(winHandle);

                if (_driver.Title.Contains(se.Value))
                    return;
            }

            if (cnt > 10)
                throw new NoSuchElementException($"查找失败[{se.Cmd}]-[{se.Value}]");
            else
                DoWin(se, cnt++);
        }

        public void DoSleep(SeleniumEvent se)
            => DoSleep(se.Value);

        public void DoSleep(string time)
            => Thread.Sleep(int.Parse(time));

        public void DoScroll(SeleniumEvent se)
        {
            string[] args = se.Value.GetSuValues();
            int num = int.Parse(args[0]);
            for (int index = 1; index <= num; ++index)
            {
                ((IJavaScriptExecutor)_driver).ExecuteScript("window.scrollTo({top: document.body.scrollHeight / " + num.ToString() + " * " + index.ToString() + ", behavior: \"smooth\"});", Array.Empty<object>());
                DoSleep(args[1]);
            }
        }

        public void DoHMove(SeleniumEvent se)
        {
            if (!_elements.ContainsKey(se.Id)) return;

            string[] suValues = se.Value.GetSuValues();
            Actions actions = new Actions(_driver);
            actions.ClickAndHold(_elements[se.Id]).MoveByOffset(int.Parse(suValues[0]), int.Parse(suValues[1])).Perform();
            actions.Release(_elements[se.Id]);
        }

        public void Click(SeleniumEvent se)
        {
            if (!_elements.ContainsKey(se.Id)) return;

            IWebElement element = _elements[se.Id];

            switch (se.Value)
            {
                case "2":
                    ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].click();", new object[1] { element });
                    break;
                case "3":
                    new Actions(_driver).MoveToElement(element).Click().Perform();
                    break;
                default:
                    element.Click();
                    break;
            }
        }

        public void GoToMove(SeleniumEvent se)
            => new Actions(_driver).MoveToElement(_elements[se.Id]).Perform();

        public void GoToFrame(SeleniumEvent se)
        {
            if (!string.IsNullOrWhiteSpace(se.Id))
                _driver.SwitchTo().Frame(_elements[se.Id]);
            else
                _driver.SwitchTo().Frame(int.Parse(se.Value));
        }

        public void GoOutFrame()
            => _driver.SwitchTo().DefaultContent();

        public void SavePic(string path, ScreenshotImageFormat type = ScreenshotImageFormat.Jpeg, bool isFull = false)
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

        public IWebElement FindIdElement(string id)
            => _driver.FindElement(By.Id(id));

        public IWebElement FindNameElement(string name, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.Name(name)));

        public IWebElement FindNameElement(string name)
            => _driver.FindElement(By.Name(name));

        public IWebElement FindClassElement(string className, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.ClassName(className)));

        public IWebElement FindClassElement(string className)
             => _driver.FindElement(By.ClassName(className));

        public IWebElement FindLinkElement(string link, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.LinkText(link)));

        public IWebElement FindLinkElement(string link)
             => _driver.FindElement(By.LinkText(link));

        public IWebElement FindCsElement(string select, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.CssSelector(select)));

        public IWebElement FindCsElement(string select)
            => _driver.FindElement(By.CssSelector(select));

        public IWebElement FindLinkPartElement(string link, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.PartialLinkText(link)));

        public IWebElement FindLinkPartElement(string link)
            => _driver.FindElement(By.PartialLinkText(link));

        public IWebElement FindXPathElement(string path, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.XPath(path)));

        public IWebElement FindXPathElement(string path)
            => _driver.FindElement(By.XPath(path));

        public IWebElement FindTagElement(string path, WebDriverWait wait)
            => DoFindEvent(wait, dr => dr.FindElement(By.TagName(path)));

        public IWebElement FindTagElement(string path)
            => _driver.FindElement(By.TagName(path));

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
        public static string[] GetSuValues(this string command)
        {
            return command.Split(';');
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
