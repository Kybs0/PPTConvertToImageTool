﻿using System;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PPTConvertToImageTool
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

        }

        private void ExportButton_OnClick(object sender, RoutedEventArgs e)
        {
            var pptFile = PptFilePathTextBox.Text;
            var result = new PptToImagesConverterByMicrosoft().ConvertToImages(pptFile, ImagesFileFolderTextBox.Text);
        }

        private void ExportButton2_OnClick(object sender, RoutedEventArgs e)
        {
            var pptFile = PptFilePathTextBox.Text;
            var result = new PptToImageConverterByAspose().ConvertToImages(pptFile, ImagesFileFolderTextBox.Text);
        }
    }
}
