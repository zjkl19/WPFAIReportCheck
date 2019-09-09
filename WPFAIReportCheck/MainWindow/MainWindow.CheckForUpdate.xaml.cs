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
        private void CheckForUpdateButton_Click(object sender, RoutedEventArgs e)
        {
            Repository.CheckForUpdate.CheckByRestClient();
        }
    }
}
