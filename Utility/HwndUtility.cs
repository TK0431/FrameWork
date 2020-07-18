using FrameWork.Consts;
using FrameWork.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;

namespace FrameWork.Utility
{
    public static class User32Utility
    {
        #region 定数

        public const int GW_CHILD = 0x5;
        public const int GW_OWNER = 0x4;
        public const int GW_HWNDNEXT = 0x2;
        public const int WM_GETTEXTLENGTH = 0x0E;
        public const int WM_GETTEXT = 0x0D;
        public const int MOUSEEVENTF_MOVE = 0x0001;
        public const int MOUSEEVENTF_LEFTDOWN = 0x0002;
        public const int MOUSEEVENTF_LEFTUP = 0x0004;
        public const int WM_SETTEXT = 0x0C;
        public const int TCM_GETCURSEL = 0x1300 + 11;
        public const int TCM_SETCURSEL = 0x1300 + 12;
        public const int TCM_SETCURFOCUS = 0x1300 + 48;
        public const int CB_SETCURSEL = 0x014E;

        #endregion

        #region 函数

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindow(IntPtr hwnd, int wCmd);

        [DllImport("user32.dll", EntryPoint = "GetDesktopWindow", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr GetDesktopWindow();

        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        public static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "SendMessage")]
        public static extern int SendMessage(IntPtr hwnd, int wMsg, int wParam, int lParam);

        [DllImport("user32.dll", EntryPoint = "SendMessageA")]
        public static extern int SendMessage(IntPtr hwnd, int wMsg, int wParam, StringBuilder lParam);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern int GetWindowRect(IntPtr hwnd, out RectModel lpRect);

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, IntPtr dwExtraInfo);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool GetCursorPos(out Point pt);

        //安装钩子
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern int SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hInstance, int threadId);
        //卸载钩子
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern bool UnhookWindowsHookEx(int idHook);
        //调用下一个钩子
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern int CallNextHookEx(int idHook, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern int SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        public static extern int mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);


        [DllImport("user32.dll")]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int nMaxCount);

        #endregion
    }

    /// <summary>
    /// 句柄类
    /// </summary>
    public static class HwndUtility
    {
        /// <summary>
        /// 获取控件值长度
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static int GetHwndValueLen(IntPtr hwnd)
            => User32Utility.SendMessage(hwnd, User32Utility.WM_GETTEXTLENGTH, 0, 0);

        /// <summary>
        /// 获取控件值长度
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static int GetHwndWinLen(IntPtr hwnd)
            => User32Utility.GetWindowTextLength(hwnd);

        /// <summary>
        /// 获取控件值
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static string GetHwndValue(IntPtr hwnd, HwndItemType tp)
        {
            if (tp == HwndItemType.Exe)
            {
                StringBuilder hwndValue = new StringBuilder(GetHwndWinLen(hwnd) + 1);
                User32Utility.GetWindowText(hwnd, hwndValue, hwndValue.Capacity);
                return hwndValue.ToString();
            }
            else
            {
                StringBuilder hwndValue = new StringBuilder(GetHwndValueLen(hwnd) + 1);

                User32Utility.SendMessage(hwnd, User32Utility.WM_GETTEXT, hwndValue.Capacity, hwndValue);
                return hwndValue.ToString();
            }
        }

        /// <summary>
        /// 获取控件类
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static string GetHwndClass(IntPtr hwnd)
        {
            StringBuilder hwndClass = new StringBuilder(255);
            User32Utility.GetClassName(hwnd, hwndClass, hwndClass.Capacity);
            return hwndClass.ToString();
        }

        /// <summary>
        /// 获取第一个子控件
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static IntPtr GetChildHwnd(IntPtr hwnd)
            => User32Utility.GetWindow(hwnd, User32Utility.GW_CHILD);

        /// <summary>
        /// 获取同级别下一个控件
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static IntPtr GetNextHwnd(IntPtr hwnd)
            => User32Utility.GetWindow(hwnd, User32Utility.GW_HWNDNEXT);

        /// <summary>
        /// 获取同级别下一个控件
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static IntPtr GetOwnerHwnd(IntPtr hwnd)
            => User32Utility.GetWindow(hwnd, User32Utility.GW_OWNER);

        /// <summary>
        /// 获取桌面句柄
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static List<IntPtr> GetDeskHwnds2()
        {
            List<IntPtr> result = new List<IntPtr>();

            IntPtr hwnd = User32Utility.GetDesktopWindow();
            hwnd = GetChildHwnd(hwnd);

            while (hwnd != IntPtr.Zero)
            {
                result.Add(hwnd);
                hwnd = GetNextHwnd(hwnd);
            }

            return result;
        }

        /// <summary>
        /// 获取桌面句柄
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static List<IntPtr> GetDeskHwnds()
        {
            List<IntPtr> result = new List<IntPtr>();

            IntPtr hwnd = User32Utility.FindWindow(null, null);

            while (hwnd != IntPtr.Zero)
            {
                result.Add(hwnd);
                hwnd = GetNextHwnd(hwnd);
            }

            return result;
        }

        /// <summary>
        /// 获取桌面句柄
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static List<IntPtr> GetChildrenHwnds(IntPtr hwnd)
        {
            List<IntPtr> result = new List<IntPtr>();

            hwnd = GetChildHwnd(hwnd);
            while (hwnd != IntPtr.Zero)
            {
                result.Add(hwnd);
                hwnd = GetNextHwnd(hwnd);
            }

            return result;
        }

        /// <summary>
        /// 获取桌面Models
        /// </summary>
        /// <returns></returns>
        public static List<HwndModel> GetDeskHwndModels()
        {
            List<HwndModel> result = new List<HwndModel>();
            foreach (IntPtr hwnd in GetDeskHwnds())
                result.Add(GetHwndModel(hwnd));
            return result;
        }

        /// <summary>
        /// 获取桌面Models
        /// </summary>
        /// <returns></returns>
        public static List<HwndModel> GetChildrenModels(IntPtr hwnd)
        {
            List<HwndModel> result = new List<HwndModel>();
            GetChildrenHwnds(hwnd).ForEach(x => result.Add(GetHwndModel(x)));
            return result;
        }

        /// <summary>
        /// 获取顶级父句柄
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static IntPtr GetTopParentHwnd(IntPtr hwnd)
        {
            List<IntPtr> topList = GetDeskHwnds();

            //IntPtr topHwnd = User32Utility.GetDesktopWindow();
            //if (topHwnd == hwnd) return hwnd;

            //IntPtr tempHwnd = User32Utility.GetParent(hwnd);
            //while (topHwnd != tempHwnd && tempHwnd != IntPtr.Zero)
            //{
            //    hwnd = tempHwnd;
            //    tempHwnd = User32Utility.GetParent(hwnd);
            //}

            while (!topList.Contains(hwnd) && hwnd != IntPtr.Zero)
            {
                hwnd = User32Utility.GetParent(hwnd);
            }

            return hwnd;
        }

        /// <summary>
        /// Get All Hwnd
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static List<IntPtr> GetAllHwnds(IntPtr hwnd)
        {
            List<IntPtr> result = new List<IntPtr>();

            if (hwnd != IntPtr.Zero)
            {
                result.Add(hwnd);

                hwnd = GetChildHwnd(hwnd);
                while (hwnd != IntPtr.Zero)
                {
                    result.AddRange(GetAllHwnds(hwnd));
                    hwnd = GetNextHwnd(hwnd);
                }
            }

            return result;
        }

        /// <summary>
        /// Get All Hwnd
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static List<HwndModel> GetAllModels(IntPtr hwnd)
        {
            List<HwndModel> result = new List<HwndModel>();
            GetAllHwnds(hwnd).ForEach(x => result.Add(GetHwndModel(x)));
            return result;
        }

        /// <summary>
        /// 获取类型
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static HwndItemType GetHwndType(IntPtr hwnd)
        {
            // 判断是否存在父节点
            //IntPtr pHwnd = GetOwnerHwnd(hwnd);
            //IntPtr topHwnd = User32Utility.GetParent(hwnd);
            if (GetDeskHwnds().Contains(hwnd)) return HwndItemType.Exe;

            // 判断是否存在子节点
            IntPtr cHwnd = GetChildHwnd(hwnd);
            if (cHwnd == IntPtr.Zero) return HwndItemType.Control;

            return HwndItemType.Group;
        }

        /// <summary>
        /// 获取句柄Model
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        public static HwndModel GetHwndModel(IntPtr hwnd, IntPtr parentHwnd = new IntPtr())
        {
            HwndModel result = new HwndModel() { HwndId = hwnd };

            // 获取类型
            result.Type = GetHwndType(hwnd);

            Console.WriteLine(hwnd);
            // 获取控件值
            result.Value = GetHwndValue(hwnd, result.Type);
            Console.WriteLine(result.Value);
            // 获取控件类
            result.Class = GetHwndClass(hwnd);

            // Desk坐标
            RectModel rect = new RectModel();
            User32Utility.GetWindowRect(hwnd, out rect);
            result.DeskX = (int)rect.Left;
            result.DeskY = (int)rect.Top;

            // Exe坐标
            RectModel parRect = new RectModel();
            if (parentHwnd == IntPtr.Zero)
                parentHwnd = GetTopParentHwnd(hwnd);
            User32Utility.GetWindowRect(parentHwnd, out parRect);
            result.ExeX = (int)(rect.Left - parRect.Left);
            result.ExeY = (int)(rect.Top - parRect.Top);

            // 宽高
            result.Width = (int)(rect.Right - rect.Left);
            result.Height = (int)(rect.Bottom - rect.Top);

            return result;
        }
    }
}
