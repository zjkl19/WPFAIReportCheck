using NLog;
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
        public List<CheckBox> OptionCheckBoxList;
        //private ILogger _log;
        public OptionWindow()
        {
            InitializeComponent();
            var config = XDocument.Load(@"Option.config");
            try
            {
                OptionCheckBoxList = new List<CheckBox>();
                foreach (var e in config.Elements("configuration").Descendants())
                {
                    CheckBox cb = new CheckBox
                    {
                        Name = e.Attribute("Name").Value.ToString(),
                        Content = e.Attribute("Content").Value.ToString(),
                        Height = 30,
                        VerticalAlignment = VerticalAlignment.Center,
                        VerticalContentAlignment = VerticalAlignment.Center
                    };

                    if (cb.Content.ToString().Length < 10)
                    {
                        cb.Width = 130;
                    }
                    else if (cb.Content.ToString().Length > 10 && cb.Content.ToString().Length < 15)
                    {
                        cb.Width = 180;
                    }
                    else    //字符串>=13的情况待定
                    {
                        cb.Width = 200;
                    }

                    if(Convert.ToInt32(e.Value)==1)
                    {
                        cb.IsChecked = true;
                    }
                    //cb.Tag = i;
                    //Thickness thickness = new Thickness(5);
                    //cb.VerticalContentAlignment = VerticalAlignment.Center;
                    //cb.Margin = thickness;
                    //cb.Click += cb_Click;
                    //cb.DataContext = new Test() { ID = i, Name = cb.Content.ToString() };

                    StackPanelContent.Children.Add(cb);

                    //dicSelPersonnal.Add(cb.Content.ToString(), false);
                    OptionCheckBoxList.Add(cb);
                }

            }
            catch (Exception ex)
            {
                //TODO:数据格式不正确时的异常处理
                throw ex;
            }
        }

        //TODO：单元测试
        private void Save_Button_Click(object sender, RoutedEventArgs e)
        {
            SaveOption();

        }

        public void SaveOption()
        {
            try
            {
                foreach (var cb in OptionCheckBoxList)
                {
                    var config = XDocument.Load(@"Option.config");
                    var MethodUpdate = (from p in config.Element("configuration").Elements() where (p.Attribute("Name").Value) == cb.Name select p).FirstOrDefault();
                    if (MethodUpdate != null)
                    {
                        MethodUpdate.Value = (cb.IsChecked ?? false) ? "1" : "0";
                        config.Save(@"Option.config");
                    }

                }
                MessageBox.Show("保存设置成功！");
            }
            catch (Exception ex)
            {
#if DEBUG
                throw ex;
#else
                //log.Error(ex, $"Save_Button_Click保存选项xml配置文件出错，错误信息：{ ex.Message.ToString()}");
#endif
            }
        }
    }
}
