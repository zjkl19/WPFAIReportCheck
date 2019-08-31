using Aspose.Words;
using Microsoft.WindowsAPICodePack.Dialogs;
//using log4net;
//using log4net.Repository;
using Ninject;
using System.Windows;


namespace WPFAIReportCheck
{

    public partial class MainWindow : Window
    {
        //免责声明
        private void DisclaimerButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("校核结果仅供参考");
        }
    }
}
