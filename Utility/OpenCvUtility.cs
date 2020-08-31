using FrameWork.Models;
using OpenCvSharp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameWork.Utility
{
    public static class OpenCvUtility
    {
        public static void AddRects(string path, List<PicRectModel> rects)
        {
            Mat src = ReadImg(path);

            foreach (PicRectModel item in rects)
            {
                Rect rect = new Rect(item.X, item.Y, item.Width, item.Height);
                Scalar color = Scalar.FromRgb(
                    Convert.ToInt32(item.Color.Substring(1, 2), 16),
                    Convert.ToInt32(item.Color.Substring(3, 2), 16),
                    Convert.ToInt32(item.Color.Substring(5, 2), 16));
                AddRect(src, rect, color, item.Thickness);
            }

            SaveImg(path, src);
        }

        public static Mat ReadImg(string path)
            => new Mat(path, ImreadModes.Color);


        public static void Show(Mat src, string title = "Show")
            => Cv2.ImShow(title, src);

        public static void Show(string path, string title = "Show")
            => Cv2.ImShow(title, ReadImg(path));

        public static void SaveImg(string file, Mat src)
            => Cv2.ImWrite(file, src);

        public static void AddRect(Mat src, Rect rect, Scalar color, int thickness = 2)
            => Cv2.Rectangle(src, rect, color, thickness);
    }
}
