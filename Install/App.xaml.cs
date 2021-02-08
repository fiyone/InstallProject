using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;

namespace Install
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            //添加程序集解析事件
            AppDomain.CurrentDomain.AssemblyResolve += OnResolveAssembly;

            InstallConfigurationEntity InstallConfigEntity = Common.GetSysInstallConfig("Install.InstallApp.config");
            var FindList = InstallConfigEntity.ListPrograms.Where(o => o.ProgramName == InstallConfigEntity.UsingProgramName).ToList();
            if (FindList.Count == 1)
            {
                Common.InstallEntity = FindList[0];
            }
            else
            {
                MessageBox.Show("系统配置异常！");
                Application.Current.Shutdown();
            }
            Application currApp = Application.Current;
            currApp.StartupUri = new Uri(Common.InstallEntity.StartupUri, UriKind.RelativeOrAbsolute);

        }

        private static Assembly OnResolveAssembly(object sender, ResolveEventArgs args)
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            var executingAssemblyName = executingAssembly.GetName();
            var resName = executingAssemblyName.Name + ".resources";

            AssemblyName assemblyName = new AssemblyName(args.Name); string path = "";
            if (resName == assemblyName.Name)
            {
                path = executingAssemblyName.Name + ".g.resources"; ;
            }
            else
            {
                path = assemblyName.Name + ".dll";
                if (assemblyName.CultureInfo != null && assemblyName.CultureInfo.Equals(CultureInfo.InvariantCulture) == false)
                {
                    path = String.Format(@"{0}\{1}", assemblyName.CultureInfo, path);
                }
            }

            using (Stream stream = executingAssembly.GetManifestResourceStream(path))
            {
                if (stream == null)
                    return null;

                byte[] assemblyRawBytes = new byte[stream.Length];
                stream.Read(assemblyRawBytes, 0, assemblyRawBytes.Length);
                return Assembly.Load(assemblyRawBytes);
            }
        }


    }
}
