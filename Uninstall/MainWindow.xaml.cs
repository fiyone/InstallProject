using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using System.Security.Principal;
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

namespace Uninstall
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private ProgramsEntity InstallEntity = Common.InstallEntity;
        /// <summary>
        /// 卸载残余文件的脚本路径
        /// </summary>
        private string batPath = null;

        public MainWindow()
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
            Application.Current.Shutdown();
        }

        private void UninstallExe()
        {
            try
            {
                //获取路径
                string cnctKJPTPath = AppDomain.CurrentDomain.BaseDirectory;
                //需要删除的个数
                double deleteNumber = 0;
                //获取需要删除的个数
                this.GetDeleteNumber(cnctKJPTPath, ref deleteNumber);
                //不能删除正在运行中的本程序
                deleteNumber = deleteNumber - 1;
                //设置进度条的值
                this.pbSchedule.Value = 0;
                this.pbSchedule.Minimum = 0;
                this.pbSchedule.Maximum = deleteNumber;
                //开始执行卸载
                this.UninstallFile(cnctKJPTPath, this.pbSchedule, this.txtSchedule, deleteNumber, cnctKJPTPath);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 通过递归获取需要删除的个数
        /// </summary>
        /// <param name="path">需要删除的文件夹路径</param>
        /// <param name="deleteNumber">需要删除的个数</param>
        private void GetDeleteNumber(string path, ref double deleteNumber)
        {
            try
            {
                //过去当前路径下的所有文件夹
                DirectoryInfo di = new DirectoryInfo(path);
                DirectoryInfo[] di1 = null;
                bool isErr = false;
                try
                {
                    di1 = di.GetDirectories();
                }
                catch (Exception e)
                {
                    if (e.Message != null && e.Message.Contains("拒绝"))
                    {
                        SetUserDirectoryAccessControl(path, true);
                        di1 = di.GetDirectories();
                        isErr = true;
                    }
                }
                if (di1 != null)
                {
                    foreach (DirectoryInfo item in di1)
                    {
                        //递归文件夹
                        GetDeleteNumber(item.FullName, ref deleteNumber);
                    }

                }
                //过去当前路径下的所有文件
                FileInfo[] fi = di.GetFiles();
                foreach (FileInfo item in fi)
                {
                    //记录文件个数
                    deleteNumber++;
                }
                if (isErr)
                {
                    SetUserDirectoryAccessControl(path, false);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// 通过递归执行程序卸载
        /// </summary>
        /// <param name="path">卸载路径</param>
        /// <param name="pb">卸载进度条</param>
        /// <param name="tb">卸载百分比</param>
        /// <param name="deleteNumber">需要删除的个数</param>
        private void UninstallFile(string path, ProgressBar pb, TextBlock tb, double deleteNumber, string RootPath = "")
        {
            try
            {
                string[] FileSystem = null;
                try
                {
                    FileSystem = Directory.GetFileSystemEntries(path);
                }
                catch (Exception e)
                {
                    if (e.Message != null && e.Message.Contains("拒绝"))
                    {
                        SetUserDirectoryAccessControl(path, true);
                        FileSystem = Directory.GetFileSystemEntries(path);
                    }
                }

                if (FileSystem != null)
                {
                    foreach (string fileFullName in FileSystem)
                    {
                        if (File.Exists(fileFullName))
                        {
                            FileInfo fi = new FileInfo(fileFullName);
                            //将只读文件改为可以删除
                            if (fi.Attributes.ToString().IndexOf("ReadOnly") != -1)
                            {
                                fi.Attributes = FileAttributes.Normal;
                            }
                            if (!fi.Name.Equals(InstallEntity.UninstallName))//不能删除正在运行中的卸载程序
                            {
                                try
                                {
                                    File.Delete(fileFullName);//删除文件
                                }
                                catch (Exception ex1)
                                {
                                    if (ex1.Message != null && ex1.Message.Contains("拒绝"))
                                    {
                                        SetUserDirectoryAccessControl(fileFullName, true);
                                        try
                                        {
                                            File.Delete(fileFullName);
                                        }
                                        catch (Exception ex2)
                                        {
                                            if (ex2.Message != null && ex2.Message.Contains("拒绝"))
                                            {
                                                //continue;
                                            }
                                        }

                                    }
                                }
                                pb.Value++;
                                this.txtSchedule.Text = Math.Round((pb.Value / deleteNumber) * 100, 0).ToString() + "%";
                            }
                        }
                        else
                        {
                            UninstallFile(fileFullName, pb, tb, deleteNumber, RootPath);//递归删除子文件夹   
                        }
                    }
                    try
                    {
                        if (path != RootPath)
                        {
                            Directory.Delete(path, true);
                        }
                    }
                    catch (Exception e)
                    {
                        if (e.Message != null && e.Message.Contains("拒绝"))
                        {
                            SetUserDirectoryAccessControl(path, true);
                            Directory.Delete(path, true);
                        }
                    }

                }

                //卸载程序正在运行中，无法删除
                //Directory.Delete(path); 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// 卸载快捷方式
        /// </summary>
        private void UninstallShortcut()
        {
            try
            {
                //获取系统默认目录
                RegistryKey HKEY_CURRENT_USER = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders");
                //获取开始菜单程序文件夹路径
                string programsPath = HKEY_CURRENT_USER.GetValue("Programs").ToString();

                //删除开始菜单快捷方式
                if (Directory.Exists(programsPath + InstallEntity.MenuFolder))
                {
                    //Directory.Delete(programsPath + @"\弘远泰斯", true);
                    if (File.Exists(programsPath + InstallEntity.MenuFolder + @"\" + InstallEntity.ShortcutName))
                    {
                        File.Delete(programsPath + InstallEntity.MenuFolder + @"\" + InstallEntity.ShortcutName);
                    }
                    if (File.Exists(programsPath + InstallEntity.MenuFolder + @"\" + InstallEntity.UninstallShortcutName))
                    {
                        File.Delete(programsPath + InstallEntity.MenuFolder + @"\" + InstallEntity.UninstallShortcutName);
                    }
                }

                //获取桌面文件夹路径
                string desktopPath = HKEY_CURRENT_USER.GetValue("Desktop").ToString();

                //删除桌面快捷方式
                if (File.Exists(desktopPath + @"\" + InstallEntity.ShortcutName))
                {
                    File.Delete(desktopPath + @"\" + InstallEntity.ShortcutName);
                }

                //常见控制面板“程序与功能”
                RegistryKey CUKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32);
                RegistryKey CurrentKey = CUKey.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + InstallEntity.DisplayName, true);
                if (CurrentKey != null)
                {
                    //删除
                    CUKey.DeleteSubKeyTree(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\" + InstallEntity.DisplayName);
                    CurrentKey.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// 生成卸载残留文件的脚本
        /// </summary>
        /// <returns>生成脚本的路径</returns>
        private void GenerateUninstallBAT()
        {
            try
            {
                //创建卸载BAT文件
                StringBuilder sb = new StringBuilder();//创建BAT内容
                sb.Append("choice /t:y,2 /n >nul" + Environment.NewLine);//开始执行删除前等待2秒用于关闭Uninstall.exe程序
                sb.Append("rd/s/q " + AppDomain.CurrentDomain.BaseDirectory);//删除文件夹命令
                string FileName = Guid.NewGuid().ToString("N") + ".bat";
                batPath = System.IO.Path.GetTempPath() + FileName;//创建BAT保存目录
                FileStream fsBAT = System.IO.File.Open(batPath, FileMode.Create);//创建BAT文件
                StreamWriter swBAT = new StreamWriter(fsBAT, Encoding.Default);//将BAT内容写入BAT文件中
                swBAT.Write(sb.ToString());
                swBAT.Close();
                fsBAT.Close();

                try
                {
                    Process proc = new Process();
                    proc.StartInfo.WorkingDirectory = System.IO.Path.GetTempPath();
                    proc.StartInfo.FileName = FileName;
                    proc.StartInfo.Arguments = string.Format("10");//this is argument
                    proc.StartInfo.CreateNoWindow = true;
                    proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    proc.Start();
                    proc.WaitForExit();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
                }
                return;
                //删除注册表数据
                //RegistryKey CUKey = Registry.CurrentUser;
                //RegistryKey CNCTKJPTKey = CUKey.OpenSubKey(@"Software\cncthlj\KJPT", true);
                //if (CNCTKJPTKey == null)
                //{
                //说明这个路径不存在，不需要删除

                //}
                //删除注册表CurrentUser/Software/CNCT/KJPT中的UseID项
                //else
                //{
                //删除软件注册表目录
                //CNCTKJPTKey.DeleteSubKeyTree("ECSV1", false);

                //////将是否记录用户账号设为false
                ////CNCTKJPTKey.DeleteValue("IsNeedRememberUserID");
                //////删除UserID项
                ////CNCTKJPTKey.DeleteValue("UserID", false);
                //}
                //CNCTKJPTKey.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 单击“卸载”图片事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void iUninstall_MouseLeftButtonDown(object sender, RoutedEventArgs e)
        {
            try
            {
                if (MessageBox.Show("是否要删除“" + InstallEntity.DisplayName + "”程序？", "提示", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                {
                    //检查软件是否在运行中
                    Process[] pCnctKJPT = Process.GetProcessesByName(InstallEntity.AppProcessName);
                    if (pCnctKJPT.Length > 0)
                    {
                        if (MessageBox.Show("“" + InstallEntity.DisplayName + "”正在运行中，是否强制关闭程序？", "提示", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                        {
                            IsAppKill(InstallEntity.AppProcessName);
                        }
                        else
                        {
                            return;
                        }
                    }
                    //设置界面控件的可见性
                    this.grid_one.Visibility = Visibility.Collapsed;
                    this.grid_Uninstalling.Visibility = Visibility.Visible;
                    this.grid_Three.Visibility = Visibility.Collapsed;

                    //杀死守护进程 保护进程是否正在使用
                    //IsAppKill("Xthk.CoursewareDecryptionTool.ConsoleApp");

                    //卸载程序
                    this.UninstallExe();
                    //删除快捷方式
                    this.UninstallShortcut();
                    //创建卸载残余文件的脚本
                    //this.GenerateUninstallBAT();

                    //设置界面控件的可见性
                    this.grid_one.Visibility = Visibility.Collapsed;
                    this.grid_Uninstalling.Visibility = Visibility.Collapsed;
                    this.grid_Three.Visibility = Visibility.Visible;
                }
            }
            catch (Exception ex)
            {
                //设置界面控件的可见性
                this.grid_one.Visibility = Visibility.Visible;
                this.grid_Uninstalling.Visibility = Visibility.Collapsed;
                this.grid_Three.Visibility = Visibility.Collapsed;
                MessageBox.Show(ex.ToString(), "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool IsAppKill(String procName, int TryCount = 10)
        {
            try
            {
                int tryCnt = 0;

                System.Diagnostics.Process[] myProcesses = System.Diagnostics.Process.GetProcessesByName(procName);

                while (myProcesses.Length > 0)
                {
                    try
                    {
                        //Application.DoEvents();
                        DispatcherHelper.DoEvents();
                        myProcesses[0].Kill();
                        myProcesses[0].WaitForExit();
                        myProcesses[0].Close();
                        myProcesses = System.Diagnostics.Process.GetProcessesByName(procName);
                    }
                    catch (Exception ex2)
                    {

                    }
                    tryCnt += 1;
                    if (tryCnt >= TryCount)//尝试10次后，终止。
                    {
                        return false;
                    }
                }

            }
            catch (Exception ex)
            {

            }
            return true;
        }


        /// <summary>
        /// 单击“完成”图片事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void iFinish_MouseLeftButtonDown(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Close();
                Task.Factory.StartNew(() =>
                {
                    //创建卸载残余文件的脚本
                    //this.GenerateUninstallBAT();
                    string ExecutablePath = AppDomain.CurrentDomain.BaseDirectory + InstallEntity.UninstallName;
                    DeleteItselfByCMD(ExecutablePath);
                });

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DeleteItselfByCMD(string ExecutablePath)
        {
            ProcessStartInfo psi = new ProcessStartInfo("cmd.exe", "/C ping 1.1.1.1 -n 1 -w 1000 > Nul & Del " + ExecutablePath);
            psi.WindowStyle = ProcessWindowStyle.Hidden;
            psi.CreateNoWindow = true;
            Process.Start(psi);
            //Application.Current.Shutdown();
        }

        /// <summary>
        /// 设置文件夹权限
        /// </summary>
        /// <param name="isRemove"></param>
        public static void SetUserDirectoryAccessControl(string path, bool isRemove)
        {

            DirectorySecurity ds = Directory.GetAccessControl(path);
            FileSystemAccessRule fsa = new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.WorldSid, null), FileSystemRights.FullControl, AccessControlType.Deny);
            if (isRemove)
                ds.RemoveAccessRule(fsa);
            else
                ds.AddAccessRule(fsa);
            Directory.SetAccessControl(path, ds);
        }


    }




}
