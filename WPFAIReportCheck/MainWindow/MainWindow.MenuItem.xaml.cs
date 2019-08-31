﻿using System;
using System.Windows;


namespace WPFAIReportCheck
{

    public partial class MainWindow : Window
    {
        private void MenuItem_About_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("系统框架设计及编程：路桥监测研究所林迪南、陈思远等");
        }

        private void MenuItem_Exit_Click(object sender, RoutedEventArgs e)
        {
            Environment.Exit(0);
        }

        private void MenuItem_ViewSourceCode_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/zjkl19/WPFAIReportCheck/");
        }

        private void MenuItem_Option_Click(object sender, RoutedEventArgs e)
        {
            var w = new OptionWindow();
            w.Top = 0.4 * (App.ScreenHeight - w.Height);
            w.Left = 0.5 * (App.ScreenWidth - w.Width);
            w.Show();
        }
    }
}
