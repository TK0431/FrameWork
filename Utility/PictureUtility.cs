using System;
using System.Drawing;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Linq;

namespace FrameWork.Utility
{
    public static class PictureUtility
    {
        #region PictureCompare

        /// <summary>
        /// Picture Compare Start
        /// </summary>
        /// <param name="pic1"></param>
        /// <param name="pic2"></param>

        public static List<Rectangle> Compare(string picPath1, string picPath2, Margins margin)
        {
            // Picture Get
            Bitmap bmp1 = (Bitmap)Bitmap.FromFile(picPath1);
            Bitmap bmp2 = (Bitmap)Bitmap.FromFile(picPath2);

            int width = (bmp1.Width < bmp2.Width ? bmp1.Width : bmp2.Width) - margin.Left - margin.Right;
            int height = (bmp1.Height < bmp2.Height ? bmp1.Height : bmp2.Height) - margin.Top - margin.Bottom;
            Rectangle range = new Rectangle(margin.Left, margin.Top, width, height);
            // 分析
            List<Point> rects = PicRgbCompare(bmp1, bmp2, range);

            // 差分あるの場合
            if (rects == null)
            {
                return new List<Rectangle>() { range };
            }
            else if (rects.Count != 0)
            {
                // Result : Red Square
                List<Rectangle> results = GetRangeString(rects);

                return results;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get Red
        /// </summary>
        /// <param name="rects"></param>
        /// <returns></returns>
        private static List<Rectangle> GetRangeString(List<Point> rects, int size = 10)
        {
            List<Rectangle> result = new List<Rectangle>();

            foreach (Point point in rects)
            {
                List<Rectangle> addRect = new List<Rectangle>();
                bool inFlg = false;
                foreach (Rectangle rect in result)
                {
                    if (rect.Top - 10 <= point.Y && point.Y <= rect.Bottom + 10 && rect.Left - 10 <= point.X && point.X <= rect.Right + 10)
                    {
                        inFlg = true;
                        addRect.Add(rect);
                    }
                }
                if (inFlg)
                {
                    Rectangle rage;
                    if (addRect.Count > 1)
                    {
                        foreach (Rectangle rect in addRect)
                        {
                            result.Remove(rect);
                        }
                        int left = addRect.Min(x => x.Left);
                        int top = addRect.Min(x => x.Top);
                        int right = addRect.Max(x => x.Right);
                        int buttom = addRect.Max(x => x.Bottom);

                        rage = new Rectangle(new Point(left, top), new Size(right - left, buttom - top));
                    }
                    else
                    {
                        rage = addRect[0];
                        result.Remove(addRect[0]);
                    }
                    rage.X = rage.Left <= point.X ? rage.X : point.X;
                    rage.Y = rage.Top <= point.Y ? rage.Y : point.Y;
                    rage.Width = rage.Right >= point.X ? rage.Width : point.X - rage.X;
                    rage.Height = rage.Bottom >= point.Y ? rage.Height : point.Y - rage.Y;
                    result.Add(rage);
                }
                else
                {
                    result.Add(new Rectangle(point, new Size(0, 0)));
                }
            }
            return result;
        }


        /// <summary>
        /// 图像颜色
        /// </summary>
        [StructLayout(LayoutKind.Explicit)]
        private struct ICColor
        {
            [FieldOffset(0)]
            public byte B;
            [FieldOffset(1)]
            public byte G;
            [FieldOffset(2)]
            public byte R;
        }

        /// <summary>
        /// 比较两个图像
        /// </summary>
        /// <param name="bmp1"></param>
        /// <param name="bmp2"></param>
        /// <param name="margin"></param>
        /// <returns></returns>
        private static List<Point> PicRgbCompare(Bitmap bmp1, Bitmap bmp2, Rectangle range, int OutRgb = 500000)
        {
            // 戻り値
            List<Point> rects = new List<Point>();

            // 比較範囲RGB取得
            BitmapData bd1 = bmp1.LockBits(range, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
            BitmapData bd2 = bmp2.LockBits(range, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);

            try
            {
                unsafe
                {
                    int w = 0, h = 0;
                    while (h < bd1.Height && h < bd2.Height)
                    {
                        byte* p1 = (byte*)bd1.Scan0 + h * bd1.Stride;
                        byte* p2 = (byte*)bd2.Scan0 + h * bd2.Stride;
                        w = 0;
                        while (w < bd1.Width && w < bd2.Width)
                        {
                            if (w >= bd1.Width || w >= bd2.Width) break;
                            if (h >= bd1.Height || h >= bd2.Height) break;

                            ICColor* pc1 = (ICColor*)(p1 + w * 3);
                            ICColor* pc2 = (ICColor*)(p2 + w * 3);
                            if (pc1->R != pc2->R || pc1->G != pc2->G || pc1->B != pc2->B)
                            {
                                //当前块有某个象素点颜色值不相同.也就是有差异.
                                int bw = Math.Min(1, bd1.Width - w);
                                int bh = Math.Min(1, bd1.Height - h);
                                rects.Add(new Point(w + range.Left, h + range.Top));
                                if (rects.Count > OutRgb)
                                {
                                    return null;
                                }
                                goto E;
                            }

                        E:
                            w++;
                        }
                        h++;
                    }
                }
            }
            finally
            {
                bmp1.UnlockBits(bd1);
                bmp2.UnlockBits(bd2);
            }

            return rects;
        }

        #endregion

        #region KeyBord Get Pictrue

        //private event KeyEventHandler KeyDownEvent;
        //private event KeyPressEventHandler KeyPressEvent;
        //private event KeyEventHandler KeyUpEvent;
        //private delegate int HookProc(int nCode, Int32 wParam, IntPtr lParam);
        //static int hKeyboardHook = 0; //声明键盘钩子处理的初始值
        ////值在Microsoft SDK的Winuser.h里查询
        //public const int WH_KEYBOARD_LL = 13;   //线程键盘钩子监听鼠标消息设为2，全局键盘监听鼠标消息设为13
        //HookProc KeyboardHookProcedure; //声明KeyboardHookProcedure作为HookProc类型
        ////键盘结构
        //[StructLayout(LayoutKind.Sequential)]
        //public class KeyboardHookStruct
        //{
        //    public int vkCode;  //定一个虚拟键码。该代码必须有一个价值的范围1至254
        //    public int scanCode; // 指定的硬件扫描码的关键
        //    public int flags;  // 键标志
        //    public int time; // 指定的时间戳记的这个讯息
        //    public int dwExtraInfo; // 指定额外信息相关的信息
        //}

        ////使用此功能，安装了一个钩子
        //[DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        //public static extern int SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hInstance, int threadId);
        ////调用此函数卸载钩子
        //[DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        //public static extern bool UnhookWindowsHookEx(int idHook);
        ////使用此功能，通过信息钩子继续下一个钩子
        //[DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        //public static extern int CallNextHookEx(int idHook, int nCode, Int32 wParam, IntPtr lParam);

        //// 取得当前线程编号（线程钩子需要用到）
        //[DllImport("kernel32.dll")]
        //static extern int GetCurrentThreadId();
        ////使用WINDOWS API函数代替获取当前实例的函数,防止钩子失效
        //[DllImport("kernel32.dll")]
        //public static extern IntPtr GetModuleHandle(string name);
        //public void Start()
        //{
        //    // 安装键盘钩子
        //    if (hKeyboardHook == 0)
        //    {
        //        KeyboardHookProcedure = new HookProc(KeyboardHookProc);
        //        hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, KeyboardHookProcedure, GetModuleHandle(System.Diagnostics.Process.GetCurrentProcess().MainModule.ModuleName), 0);
        //        //hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, KeyboardHookProcedure, Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().GetModules()[0]), 0);
        //        //************************************
        //        //键盘线程钩子
        //        //SetWindowsHookEx( 2,KeyboardHookProcedure, IntPtr.Zero, GetCurrentThreadId());//指定要监听的线程idGetCurrentThreadId(),
        //        //键盘全局钩子,需要引用空间(using System.Reflection;)
        //        //SetWindowsHookEx( 13,MouseHookProcedure,Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().GetModules()[0]),0);
        //        //
        //        if (hKeyboardHook == 0)
        //        {
        //            Stop();
        //            throw new Exception("安装键盘钩子失败");
        //        }
        //    }
        //}

        //public void Stop()
        //{
        //    bool retKeyboard = true;
        //    if (hKeyboardHook != 0)
        //    {
        //        retKeyboard = UnhookWindowsHookEx(hKeyboardHook);
        //        hKeyboardHook = 0;
        //    }
        //    if (!(retKeyboard)) throw new Exception("卸载钩子失败！");
        //}

        ////ToAscii职能的转换指定的虚拟键码和键盘状态的相应字符或字符
        //[DllImport("user32")]
        //public static extern int ToAscii(int uVirtKey, //[in] 指定虚拟关键代码进行翻译。
        //                                 int uScanCode, // [in] 指定的硬件扫描码的关键须翻译成英文。高阶位的这个值设定的关键，如果是（不压）
        //                                 byte[] lpbKeyState, // [in] 指针，以256字节数组，包含当前键盘的状态。每个元素（字节）的数组包含状态的一个关键。如果高阶位的字节是一套，关键是下跌（按下）。在低比特，如果设置表明，关键是对切换。在此功能，只有肘位的CAPS LOCK键是相关的。在切换状态的NUM个锁和滚动锁定键被忽略。
        //                                 byte[] lpwTransKey,
        //                                 int fuState);
        ////获取按键的状态
        //[DllImport("user32")]
        //public static extern int GetKeyboardState(byte[] pbKeyState);
        //[DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        //private static extern short GetKeyState(int vKey);
        //private const int WM_KEYDOWN = 0x100;//KEYDOWN
        //private const int WM_KEYUP = 0x101;//KEYUP
        //private const int WM_SYSKEYDOWN = 0x104;//SYSKEYDOWN
        //private const int WM_SYSKEYUP = 0x105;//SYSKEYUP
        //private int KeyboardHookProc(int nCode, Int32 wParam, IntPtr lParam)
        //{
        //    // 侦听键盘事件
        //    if ((nCode >= 0) && (KeyDownEvent != null || KeyUpEvent != null || KeyPressEvent != null))
        //    {
        //        KeyboardHookStruct MyKeyboardHookStruct = (KeyboardHookStruct)Marshal.PtrToStructure(lParam, typeof(KeyboardHookStruct));
        //        // raise KeyDown
        //        if (KeyDownEvent != null && (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN))
        //        {
        //            Keys keyData = (Keys)MyKeyboardHookStruct.vkCode;
        //            KeyEventArgs e = new KeyEventArgs(keyData);
        //            KeyDownEvent(this, e);
        //        }
        //        //键盘按下
        //        if (KeyPressEvent != null && wParam == WM_KEYDOWN)
        //        {
        //            byte[] keyState = new byte[256];
        //            GetKeyboardState(keyState);
        //            byte[] inBuffer = new byte[2];
        //            if (ToAscii(MyKeyboardHookStruct.vkCode, MyKeyboardHookStruct.scanCode, keyState, inBuffer, MyKeyboardHookStruct.flags) == 1)
        //            {
        //                KeyPressEventArgs e = new KeyPressEventArgs((char)inBuffer[0]);
        //                KeyPressEvent(this, e);
        //            }
        //        }
        //        // 键盘抬起
        //        if (KeyUpEvent != null && (wParam == WM_KEYUP || wParam == WM_SYSKEYUP))
        //        {
        //            Keys keyData = (Keys)MyKeyboardHookStruct.vkCode;
        //            KeyEventArgs e = new KeyEventArgs(keyData);
        //            KeyUpEvent(this, e);
        //        }
        //    }
        //    //如果返回1，则结束消息，这个消息到此为止，不再传递。
        //    //如果返回0或调用CallNextHookEx函数则消息出了这个钩子继续往下传递，也就是传给消息真正的接受者
        //    return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
        //}
        //~KeyboardHook()
        //{
        //    Stop();
        //}
        #endregion
    }
}
