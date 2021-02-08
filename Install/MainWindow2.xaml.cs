using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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

namespace Install
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow2 : Window
    {
        private ProgramsEntity InstallEntity = Common.InstallEntity;

        /// <summary>
        /// 是否安装完成
        /// </summary>
        private bool IsFinished = false;
        /// <summary>
        /// 安装路径
        /// </summary>
        private string InstallPath = "";
        /// <summary>
        /// 快捷方式名称
        /// </summary>
        private string shortName = "";


        public MainWindow2()
        {
            InitializeComponent();

            this.Loaded += MainWindow_Loaded;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                this.btnClose.Click += BtnClose_Click;
                this.btnMin.Click += BtnMin_Click;
                this.grid_title.MouseDown += Grid_title_MouseDown;

                if (!string.IsNullOrEmpty(InstallEntity.DefaultInstallPath) && InstallEntity.DefaultInstallPath.Contains("{0}"))
                {
                    string Local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                    InstallEntity.DefaultInstallPath = string.Format(InstallEntity.DefaultInstallPath, Local);
                }

                this.txtInstallationPath.Text = InstallEntity.DefaultInstallPath;

            }
            catch (Exception ex)
            {
            }
        }

        private void BtnMin_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void Grid_title_MouseDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            if (!IsFinished)
            {
                MessageBoxResult result = MessageBox.Show("未完成安装，是否确认“退出“", "提示", MessageBoxButton.OKCancel);
                if (result != MessageBoxResult.OK)
                {
                    return;
                }
            }
            Application.Current.Shutdown();
        }


        /// <summary>
        /// 单击浏览安装路径按钮事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <summary>
        /// 单击浏览安装路径按钮事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
                fbd.Description = "请选择安装路径";
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string selectPath = fbd.SelectedPath;
                    if (selectPath.Length == 3)
                    {
                        selectPath += InstallEntity.HomeDirectory;
                    }
                    string fileName = selectPath.Split('\\')[selectPath.Split('\\').Count() - 1];
                    if (!fileName.Equals(InstallEntity.InstallFolderName))
                    {
                        selectPath = selectPath + @"\" + InstallEntity.InstallFolderName;
                    }
                    this.txtInstallationPath.Text = selectPath;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 单击安装按钮事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSetup_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                try
                {
                    //遍历用户录入路径字符串的每一个字符
                    foreach (char userPathChar in this.txtInstallationPath.Text.ToArray<char>())
                    {
                        //判断用户录入的路径字符串中是否包含有特殊非法字符
                        foreach (var pathChars in System.IO.Path.GetInvalidPathChars())
                        {
                            if (userPathChar.Equals(pathChars))
                            {
                                MessageBox.Show("安装路径中包含有非法字符，请重新输入或选择。", "提示", MessageBoxButton.OK);
                                return;
                            }
                        }
                    }
                    //判断用户录入的路径径字符串是否包含根
                    //if (!System.IO.Path.IsPathRooted(this.txtInstallationPath.Text))
                    //{
                    //    MessageBox.Show("安装路径中没有包含根路径，请重新输入或选择。", "提示", MessageBoxButton.OK);
                    //    return;
                    //}
                    //当用户录入的路径不存在
                    if (!Directory.Exists(this.txtInstallationPath.Text))
                    {
                        //通过使用用户录入的路径创建目录,判断用户录入的路径是否正确
                        Directory.CreateDirectory(this.txtInstallationPath.Text);

                    }
                }
                catch (Exception ex1)
                {
                    MessageBox.Show(ex1.Message, "提示", MessageBoxButton.OK);
                    return;
                }

                //执行软件安装
                this.Setup();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 执行软件安装
        /// </summary>
        private void Setup()
        {
            try
            {
                IsFinished = false;
                //获取用户选择路径中的最底层文件夹名称
                string fileName = this.txtInstallationPath.Text.Split('\\')[this.txtInstallationPath.Text.Split('\\').Count() - 1];

                //当用户选择的安装路径中最底层的文件夹名称不是“XthkDecryptionTool”时，自动在创建一个“XthkDecryptionTool”文件夹，防止在删除的时候误删别的文件
                if (!fileName.Equals(InstallEntity.InstallFolderName))
                {
                    this.txtInstallationPath.Text = this.txtInstallationPath.Text + @"\" + InstallEntity.InstallFolderName;
                }
                //安装路径
                InstallPath = this.txtInstallationPath.Text;

                //显示安装进度界面
                //this.tcMain.SelectedIndex = 1;
                this.grid_one.Visibility = Visibility.Collapsed;
                this.grid_two.Visibility = Visibility.Visible;
                this.grid_three.Visibility = Visibility.Collapsed;

                //检测是否已经打开
                Process[] procCoursewareDecryptionTool = Process.GetProcessesByName(InstallEntity.AppProcessName);
                if (procCoursewareDecryptionTool.Any())
                {
                    if (MessageBox.Show("“" + InstallEntity.DisplayName + "”正在运行中，是否强制覆盖程序？", "提示", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                    {
                        Common.IsAppKill(InstallEntity.AppProcessName);
                    }
                    else
                    {
                        Application.Current.Shutdown();

                    }
                }

                //创建用户指定的安装目录文件夹
                Directory.CreateDirectory(InstallPath);
                ZIPHelper.ActionProgress -= ActionProgressResult;
                ZIPHelper.ActionProgress += ActionProgressResult;

                this.pbSchedule.Value = 0;
                this.txtSchedule.Text = "0%";

                //将软件解压到用户指定目录
                ZIPHelper.Extract(Install.SetupFiles.Setup, InstallPath, 1024 * 1204);
                //将嵌入的资源释放到用户选择的安装目录下面（卸载程序）
                string uninstallPath = this.txtInstallationPath.Text + @"\" + InstallEntity.UninstallName;
                FileStream fsUninstall = System.IO.File.Open(uninstallPath, FileMode.Create);
                fsUninstall.Write(Install.SetupFiles.Uninstall, 0, Install.SetupFiles.Uninstall.Length);
                fsUninstall.Close();

                //将嵌入的资源释放到用户选择的安装目录下面（快捷图标）
                string InstallIcoPath = this.txtInstallationPath.Text + InstallEntity.IconDirectoryPath;
                FileStream fsInstallIcoPath = System.IO.File.Open(InstallIcoPath, FileMode.Create);
                var InstallIco = Install.SetupFiles.IcoInstall;
                byte[] byInstall = Common.ImageToByteArray(InstallIco);
                fsInstallIcoPath.Write(byInstall, 0, byInstall.Length);
                fsInstallIcoPath.Close();

                //将嵌入的资源释放到用户选择的安装目录下面（快捷卸载图标）
                string UninstallIcoPath = this.txtInstallationPath.Text + InstallEntity.UninstallIconDirectoryPath;
                FileStream fsUninStallIco = System.IO.File.Open(UninstallIcoPath, FileMode.Create);
                var UnInstallIco = Install.SetupFiles.IcoUninstall;
                byte[] byUnInstall = Common.ImageToByteArray(UnInstallIco);
                fsUninStallIco.Write(byUnInstall, 0, byUnInstall.Length);
                fsUninStallIco.Close();

                //释放卸载程序完成，更新进度条
                this.pbSchedule.Value = this.pbSchedule.Value + 1;
                this.txtSchedule.Text = Math.Round((this.pbSchedule.Value / this.pbSchedule.Maximum * 100), 0).ToString() + "%";


                //添加开始菜单快捷方式
                RegistryKey HKEY_CURRENT_USER = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders");
                string programsPath = HKEY_CURRENT_USER.GetValue("Programs").ToString();//获取开始菜单程序文件夹路径
                Directory.CreateDirectory(programsPath + InstallEntity.MenuFolder);//在程序文件夹中创建快捷方式的文件夹

                //更新进度条
                this.pbSchedule.Value = this.pbSchedule.Value + 1;
                this.txtSchedule.Text = Math.Round((this.pbSchedule.Value / this.pbSchedule.Maximum * 100), 0).ToString() + "%";

                //快捷方式名称";
                string IconPath = InstallPath + InstallEntity.IconDirectoryPath;
                string UninstallIconPath = InstallPath + InstallEntity.UninstallIconDirectoryPath;
                string InstallExePath = InstallPath + @"\" + InstallEntity.AppExeName;
                string ExeUnInstallPath = InstallPath + @"\" + InstallEntity.UninstallName;

                //开始菜单打开快捷方式
                shortName = programsPath + InstallEntity.MenuFolder + InstallEntity.ShortcutName;
                Common.CreateShortcut(shortName, InstallExePath, IconPath);//创建快捷方式


                //更新进度条
                this.pbSchedule.Value = this.pbSchedule.Value + 1;
                this.txtSchedule.Text = Math.Round((this.pbSchedule.Value / this.pbSchedule.Maximum * 100), 0).ToString() + "%";
                //开始菜单卸载快捷方式
                Common.CreateShortcut(programsPath + InstallEntity.MenuFolder + InstallEntity.UninstallShortcutName, ExeUnInstallPath, UninstallIconPath);//创建卸载快捷方式

                //更新进度条
                this.pbSchedule.Value = this.pbSchedule.Value + 1;
                this.txtSchedule.Text = Math.Round((this.pbSchedule.Value / this.pbSchedule.Maximum * 100), 0).ToString() + "%";

                //添加桌面快捷方式
                string desktopPath = HKEY_CURRENT_USER.GetValue("Desktop").ToString();//获取桌面文件夹路径
                shortName = desktopPath + @"\" + InstallEntity.ShortcutName;
                Common.CreateShortcut(shortName, InstallExePath, IconPath);//创建快捷方式

                //常见控制面板“程序与功能”
                //可以往root里面写，root需要管理员权限，如果使用了管理员权限，主程序也会以管理员打开，如需常规打开，需要在打开进程的时候做降权处理
                RegistryKey CUKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32);
                var currentVersion = CUKey.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall");
                Dictionary<string, string> dic = new Dictionary<string, string>();
                dic.Add("DisplayIcon", InstallExePath);//显示的图标的exe
                dic.Add("DisplayName", InstallEntity.DisplayName);//名称
                dic.Add("Publisher", InstallEntity.Publisher);//发布者
                dic.Add("UninstallString", ExeUnInstallPath);//卸载的exe路径
                dic.Add("DisplayVersion", InstallEntity.VersionNumber);
                RegistryKey CurrentKey = CUKey.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + InstallEntity.DisplayName, true);
                if (CurrentKey == null)
                {
                    //说明这个路径不存在，需要创建
                    CUKey.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + InstallEntity.DisplayName);
                    CurrentKey = CUKey.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + InstallEntity.DisplayName, true);
                }
                foreach (var item in dic)
                {
                    CurrentKey.SetValue(item.Key, item.Value);
                }
                CurrentKey.Close();


                //更新进度条
                this.pbSchedule.Value = this.pbSchedule.Value + 1;
                this.txtSchedule.Text = Math.Round((this.pbSchedule.Value / this.pbSchedule.Maximum * 100), 0).ToString() + "%";

                //安装完毕，显示结束界面
                this.grid_one.Visibility = Visibility.Collapsed;
                this.grid_two.Visibility = Visibility.Collapsed;
                this.grid_three.Visibility = Visibility.Visible;

                IsFinished = true;
            }
            catch (Exception)
            {
                //安装完毕，显示结束界面
                this.grid_one.Visibility = Visibility.Visible;
                this.grid_two.Visibility = Visibility.Collapsed;
                this.grid_three.Visibility = Visibility.Collapsed;
                throw;
            }
        }

        public void ActionProgressResult(double t1, double t2, string t3)
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                this.pbSchedule.Maximum = t1;
                this.pbSchedule.Value = t2;
                this.txtSchedule.Text = t3;
            }));
        }


        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            string strFilePath = InstallPath + @"\" + InstallEntity.AppExeName;
            System.Diagnostics.Process.Start(shortName);
            this.Close();
        }

        #region

        #endregion




    }
}
