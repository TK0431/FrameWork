using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace FrameWork.Utility
{
    public class FileUtility
    {
        public static bool isValidFileContent(string filePath1, string filePath2)
        {
            string[] lines1 = File.ReadAllLines(filePath1);
            string[] lines2 = File.ReadAllLines(filePath2);

            if (lines1.Length != lines2.Length) return false;

            for (int i = 0; i < lines1.Length; i++)
                if (lines1[i] != lines2[i]) return false;

            return true;
        }
    }
}
