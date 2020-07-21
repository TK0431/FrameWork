using FrameWork.Models;
using PropertyChanged;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Controls;

namespace FrameWork.ViewModels.Base
{
    public class SeleniumBaseViewModel : BaseViewModel
    {
        /// <summary>
        /// Exel脚本Model
        /// </summary>
        public SeleniumScriptModel ExcelModel { get; set; }

        /// <summary>
        /// Exel脚本Model
        /// </summary>
        public bool FlgStop { get; set; } = false;
    }
}
