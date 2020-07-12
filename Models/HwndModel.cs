using FrameWork.Consts;
using System;

namespace FrameWork.Models
{
    /// <summary>
    /// 句柄Model
    /// </summary>
    public class HwndModel
    {
        /// <summary>
        /// Hwnd Id
        /// </summary>
        public IntPtr HwndId { set; get; }

        /// <summary>
        /// 类型
        /// </summary>
        public HwndItemType Type { set; get; } = HwndItemType.Control;

        /// <summary>
        /// 内容
        /// </summary>
        public string Value { set; get; }

        /// <summary>
        /// 类
        /// </summary>
        public string Class { set; get; }

        /// <summary>
        /// Desk 坐标：Left
        /// </summary>
        public int DeskX { get; set; }

        /// <summary>
        /// Desk 座標：Top
        /// </summary>
        public int DeskY { get; set; }

        /// <summary>
        /// App 座標：Left
        /// </summary>
        public int ExeX { get; set; }

        /// <summary>
        /// App 座標：Top
        /// </summary>
        public int ExeY { get; set; }

        /// <summary>
        /// Width
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Height
        /// </summary>
        public int Height { get; set; }
    }

    /// <summary>
    /// RectModel
    /// </summary>
    public struct RectModel
    {
        /// <summary>
        /// Left
        /// </summary>
        public int Left;

        /// <summary>
        /// Top
        /// </summary>
        public int Top;

        /// <summary>
        /// Right
        /// </summary>
        public int Right;

        /// <summary>
        /// Bottom
        /// </summary>
        public int Bottom;
    }
}
