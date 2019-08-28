using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Xml.Linq;

namespace WPFAIReportCheck
{
    /// <summary>
    /// OptionWindow.xaml 的交互逻辑
    /// </summary>
    public partial class OptionWindow : Window
    {
        public OptionWindow()
        {
            InitializeComponent();
            var config = XDocument.Load(@"Option.config");
            var t = Convert.ToInt32(config.Element("configuration").Element("FindUnitError").Value);
            if(t==1)
            {
                FindUnitErrorCheckBox.IsChecked= true;
            }

        }
    }
}
