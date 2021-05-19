using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Office_Installer_Tool
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        readonly string DefaultButtonText = null;
        public MainWindow()
        {
            InitializeComponent();
            DefaultButtonText = (string)Button_Install.Content;
            try
            {
                string filename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "officeKey.txt");
                if (File.Exists(filename))
                {
                    using (StreamReader streamReader = new StreamReader(filename, Encoding.UTF8))
                    {
                        for (int i = 0; i < 3; i++)
                        {
                            string text = streamReader.ReadLine();
                            var sp = text.Split(new char[] { '=' });
                            if (sp.Length > 1)
                            {
                                switch (sp[0].ToLower().Trim())
                                {
                                    case "office":
                                        TextBox_OfficeKey.Text = sp[1].ToUpper().Trim();
                                        break;
                                    case "visio":
                                        TextBox_VisioKey.Text = sp[1].ToUpper().Trim();
                                        break;
                                    case "project":
                                        TextBox_ProjectKey.Text = sp[1].ToUpper().Trim();
                                        break;
                                }
                            }
                        }
                    }
                }
            }
            catch { }
        }
        HashSet<string> list = new HashSet<string>() { "Access", "Groove", "Lync", "OneDrive", "OneNote", "Outlook", "Publisher" };
        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            string select = ((CheckBox)sender).Content as string;
            if (list.Contains(select))
                list.Remove(select);
        }
        private void CheckBox_UnChecked(object sender, RoutedEventArgs e)
        {
            string select = ((CheckBox)sender).Content as string;
            if (!list.Contains(select))
                list.Add(select);
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            string configureFilePath = Path.GetTempFileName();
            using (StreamWriter sw = new StreamWriter(configureFilePath, false, Encoding.UTF8))
            {
                sw.WriteLine("<Configuration>");

                sw.WriteLine($"<Add OfficeClientEdition=\"{ComboBox_OfficeClientEdition.Text.Replace("x", "")}\" SourcePath=\"{(RadioButton_DownloadPath.IsChecked == true ? AppDomain.CurrentDomain.BaseDirectory : Path.GetTempPath())}\" Channel=\"PerpetualVL2019\">");

                #region Office
                if (ComboBox_OfficeVersion.SelectedIndex == 0)
                    AddExcludeApp("ProPlus2019Volume", TextBox_OfficeKey.Text);
                else
                    AddExcludeApp("Standard2019Volume", TextBox_OfficeKey.Text);

                #endregion Office
                #region Visio
                if (Checkbox_VisioPro.IsChecked == true)
                {
                    if (ComboBox_OfficeVersion.SelectedIndex == 0)
                        AddExcludeApp("VisioPro2019Volume", TextBox_VisioKey.Text);
                    else
                        AddExcludeApp("VisioStd2019Volume", TextBox_VisioKey.Text);                
                }
                #endregion Visio

                #region Project
                if (Checkbox_ProjectPro.IsChecked == true)
                {
                    if (ComboBox_OfficeVersion.SelectedIndex == 0)
                        AddExcludeApp("ProjectPro2019Volume", TextBox_ProjectKey.Text);
                    else
                        AddExcludeApp("ProjectStd2019Volume", TextBox_ProjectKey.Text);                  
                }
                #endregion Project

                sw.WriteLine("<Product ID=\"ProofingTools\">");
                sw.WriteLine("<Language ID=\"en-us\" />");
                sw.WriteLine("<Language ID=\"zh-tw\" />");
                sw.WriteLine("</Product>");

                sw.WriteLine("</Add>");
                sw.WriteLine("<Updates Enabled=\"TRUE\" />");
                if (CheckBox_RemoveMSI.IsChecked == true)
                    sw.WriteLine("<RemoveMSI />");
                sw.WriteLine("</Configuration>");
                sw.Flush();
                void AddExcludeApp(string productID, string key)
                {
                    sw.WriteLine($"<Product ID=\"{productID}\" PIDKEY=\"{key}\">");
                    sw.WriteLine("<Language ID=\"MatchOS\" />");
                    foreach (var s in list)
                    {
                        sw.WriteLine($"<ExcludeApp ID=\"{s}\" />");
                    }
                    sw.WriteLine("</Product>");
                }
            }
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();            
            startInfo.FileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Tools\\setup.exe");
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            //設定啟動動作,確保以管理員身份執行 
            startInfo.Verb = "runas";
            

            startInfo.Arguments = $"/download \"{configureFilePath}\"";
            Button_Install.IsEnabled = false;
            Cursor = Cursors.Wait;
            Button_Install.Content = "Downloading...";
            using var process = System.Diagnostics.Process.Start(startInfo);
            await process.WaitForExitAsync();

            Button_Install.Content = "Installing...";
            startInfo.Arguments = $"/configure \"{configureFilePath}\"";
            System.Diagnostics.Process.Start(startInfo);
            await process.WaitForExitAsync();
            Cursor = null;
            Button_Install.Content = DefaultButtonText;
            Button_Install.IsEnabled = true;

           
        }
    }
}
