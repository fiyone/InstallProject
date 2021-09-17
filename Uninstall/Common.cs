using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Xml;

namespace Uninstall
{
    public class Common
    {
        public static ProgramsEntity InstallEntity;
        public static InstallConfigurationEntity GetSysInstallConfig(string FileName)
        {
            try
            {
                InstallConfigurationEntity Result = new InstallConfigurationEntity();

                string StrXml = "";
                //如果不存在 则从嵌入资源内读取 BlockSet.xml 
                Assembly asm = Assembly.GetExecutingAssembly();//读取嵌入式资源
                using (Stream sm = asm.GetManifestResourceStream(FileName))
                {
                    StreamReader reader = new StreamReader(sm);
                    StrXml = reader.ReadToEnd();
                    reader.Close();
                    reader.Dispose();
                }
                XmlDocument xd = new XmlDocument();
                xd.LoadXml(StrXml);
                foreach (XmlNode item in xd.ChildNodes)
                {
                    foreach (XmlNode item1 in item.ChildNodes)
                    {
                        switch (item1.Name)
                        {
                            case "Programs":
                                List<ProgramsEntity> ListPrograms = new List<ProgramsEntity>();
                                foreach (XmlNode item2 in item1.ChildNodes)
                                {
                                    if (item2.Name == "#comment")
                                    {
                                        continue;
                                    }
                                    ProgramsEntity programs = new ProgramsEntity();
                                    foreach (XmlNode item3 in item2.ChildNodes)
                                    {
                                        if (item3.Name == "#comment")
                                        {
                                            continue;
                                        }
                                        programs = (ProgramsEntity)SetDataValue(programs, item3.Name, item3.InnerText);
                                    }
                                    ListPrograms.Add(programs);
                                }
                                Result.ListPrograms = ListPrograms;
                                break;
                            case "UsingProgramName":
                                Result.UsingProgramName = item1.InnerText;
                                break;
                            default:
                                break;
                        }
                    }
                }

                return Result;
            }
            catch (Exception)
            {

                throw;
            }

        }
        /// <summary>
        /// 根据反射赋值
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="Name"></param>
        /// <param name="Value"></param>
        /// <returns></returns>
        public static object SetDataValue(object obj, string Name, string Value)
        {
            if (obj != null)
            {
                Type type = obj.GetType();
                foreach (var item in type.GetProperties())
                {
                    if (item.Name == Name)
                    {
                        item.SetValue(obj, string.IsNullOrEmpty(Value) ? null : Value, null);

                    }
                }
                return obj;
            }
            return null;
        }
    }

    public class InstallConfigurationEntity
    {
        public string UsingProgramName { get; set; }

        public List<ProgramsEntity> ListPrograms { get; set; }
    }

    public class ProgramsEntity
    {
        /// <summary>
        /// 配置名称 Xthk.CoursewareDecryptionTool
        /// </summary>
        public string ProgramName { get; set; }
        /// <summary>
        /// 默认安装路径 D:\Xthk\XthkDecryptionTool
        /// </summary>
        public string DefaultInstallPath { get; set; }
        /// <summary>
        /// 程序名称 
        /// </summary>
        public string DisplayName { get; set; }
        /// <summary>
        /// 发布者 
        /// </summary>
        public string Publisher { get; set; }
        /// <summary>
        /// 应用程序EXE名称 Xthk.CoursewareDecryptionTool.exe
        /// </summary>
        public string AppExeName { get; set; }
        /// <summary>
        /// 应用程序进程名称（覆盖安装和卸载用） Xthk.CoursewareDecryptionTool
        /// </summary>
        public string AppProcessName { get; set; }
        /// <summary>
        /// 卸载程序EXE名称 Uninstall.exe
        /// </summary>
        public string UninstallName { get; set; }
        /// <summary>
        /// 主目录（当路径为系统盘（C:/）时加一个子目录） Xthk
        /// </summary>
        public string HomeDirectory { get; set; }
        /// <summary> 
        /// 开始菜单快捷方式文件夹 \XthkDecryptionTool\
        /// </summary>
        public string MenuFolder { get; set; }
        /// <summary>
        /// 快捷方式名称
        /// </summary>
        public string ShortcutName { get; set; }
        /// <summary>
        /// 快捷方式卸载名称 
        /// </summary>
        public string UninstallShortcutName { get; set; }
        /// <summary>
        /// 安装目录中快捷图标名称 \favicon-xv.ico
        /// </summary>
        public string IconDirectoryPath { get; set; }
        /// <summary>
        /// 安装目录中卸载快捷图标名称 \Uninstall.ico
        /// </summary>
        public string UninstallIconDirectoryPath { get; set; }
        /// <summary>
        /// 安装的文件夹 XthkDecryptionTool
        /// </summary>
        public string InstallFolderName { get; set; }
        /// <summary>
        /// 版本号 0.0.0.1
        /// </summary>
        public string VersionNumber { get; set; }
        /// <summary>
        /// 启动路径
        /// </summary>
        public string StartupUri { get; set; }

    }


    public static class DispatcherHelper
    {
        [SecurityPermissionAttribute(SecurityAction.Demand, Flags = SecurityPermissionFlag.UnmanagedCode)]
        public static void DoEvents()
        {
            DispatcherFrame frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, new DispatcherOperationCallback(ExitFrames), frame);
            try { Dispatcher.PushFrame(frame); }
            catch (InvalidOperationException) { }
        }
        private static object ExitFrames(object frame)
        {
            ((DispatcherFrame)frame).Continue = false;
            return null;
        }
    }

}
