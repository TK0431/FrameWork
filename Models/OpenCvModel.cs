using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameWork.Models
{
    public class PicRectModel
    {
        public string Sheet { get; set; }

        public int PicNo { get; set; }

        public int X { get; set; }

        public int Y { get; set; }

        public int Width { get; set; }

        public int Height { get; set; }

        public string Color { get; set; }

        public int Thickness { get; set; }

        public static List<PicRectModel> GetList(string file)
        {
            List<PicRectModel> result = new List<PicRectModel>();

            foreach (string line in File.ReadAllLines(file))
            {
                string[] arr = line.Split(';');
                result.Add(new PicRectModel()
                {
                    Sheet = arr[0],
                    PicNo = int.Parse(arr[1]),
                    X = int.Parse(arr[2]),
                    Y = int.Parse(arr[3]),
                    Width = int.Parse(arr[4]),
                    Height = int.Parse(arr[5]),
                    Color = arr[6],
                    Thickness = int.Parse(arr[7]),
                });
            }

            return result;
        }
    }
}
