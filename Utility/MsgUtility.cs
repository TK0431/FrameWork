using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FrameWork.Utility
{
    public static class MsgUtility
    {
        private const string TITLE = "Message";

        public static void Error(string msg, Exception e = null)
        {
            if (e == null)
            {
                MessageBox.Show(msg, TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (MessageBox.Show(msg + "\n（詳細のエラーの表示？）", TITLE, MessageBoxButtons.OKCancel, MessageBoxIcon.Error) == DialogResult.OK)
                {
                    MessageBox.Show($"({e.TargetSite.Name}){e.Message}", TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public static void Info(string msg)
        {
            MessageBox.Show(msg, TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
