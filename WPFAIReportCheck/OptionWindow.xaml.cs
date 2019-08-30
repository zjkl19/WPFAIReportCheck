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
            //var t = Convert.ToInt32(config.Element("configuration").Element("FindUnitError").Value);
            //if(t==1)
            //{
            //    FindUnitErrorCheckBox.IsChecked= true;
            //}

            foreach (var e in config.Elements("configuration").Descendants())
            {
                try
                {
                    CheckBox cb = new CheckBox();
                    cb.Name= e.Attribute("Name").Value.ToString();
                    cb.Content = e.Attribute("Content").Value.ToString();
                    if(cb.Content.ToString().Length<8)
                    {
                        cb.Width = 130;
                    }
                    else if(cb.Content.ToString().Length > 8 && cb.Content.ToString().Length <13)
                    {
                        cb.Width = 180;
                    }
                    else    //字符串>=13的情况待定
                    {
                        cb.Width = 200;
                    }
                    cb.Height = 30;
                    cb.VerticalAlignment= VerticalAlignment.Center;
                    cb.VerticalContentAlignment = VerticalAlignment.Center;
                    //cb.Tag = i;
                    //Thickness thickness = new Thickness(5);
                    //cb.VerticalContentAlignment = VerticalAlignment.Center;
                    //cb.Margin = thickness;
                    //cb.Click += cb_Click;
                    //cb.DataContext = new Test() { ID = i, Name = cb.Content.ToString() };
                    StackPanelContent.Children.Add(cb);
                    //dicSelPersonnal.Add(cb.Content.ToString(), false);
                }
                catch (Exception ex)
                {
                    //TODO:数据格式不正确时的异常处理
                    throw ex;
                }
            }

            //for (int i = 1; i <= 15; i++)
            //{

            //    CheckBox cb = new CheckBox();
            //    cb.Content = "名称" + i;
            //    cb.Width = 80;
            //    cb.Tag = i;
            //    Thickness thickness = new Thickness(5);
            //    cb.VerticalContentAlignment = VerticalAlignment.Center;
            //    cb.Margin = thickness;
            //    cb.Click += cb_Click;
            //    cb.DataContext = new Test() { ID = i, Name = cb.Content.ToString() };
            //    StackPanelContent.Children.Add(cb);
            //    dicSelPersonnal.Add(cb.Content.ToString(), false);

            //}

        }
        class Test
        {
            public int ID { get; set; }
            public string Name { get; set; }
        }
    }
}
