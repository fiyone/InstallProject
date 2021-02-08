using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace Uninstall
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            InstallConfigurationEntity InstallConfigEntity = Common.GetSysInstallConfig("Uninstall.InstallApp.config");
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
    }
}
