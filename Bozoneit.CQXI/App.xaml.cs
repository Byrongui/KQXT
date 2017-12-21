using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Threading;

namespace Bozoneit.CQXI
{

   

    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        static App()
        {
            
            //DispatcherUnhandledException += App_DispatcherUnhandledException;
            
        }

        /// <summary>
        /// 是否有多个线程
        /// </summary>
        private static bool multiProcess = false;

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            try
            {

                if (multiProcess)
                {
                    e.Handled = true;
                    return;
                }
               
                StringBuilder stringBuilder = new StringBuilder();
                stringBuilder.AppendFormat("应用程序出现了一个未知的异常\r\n" + e.Exception.Message);
                stringBuilder.AppendFormat("\r\n" + e.Exception.StackTrace);

                Exception innerException = e.Exception.InnerException;

                while (innerException != null)
                {
                    stringBuilder.AppendFormat("\r\n" + innerException.Message);
                    stringBuilder.AppendFormat("\r\n" + innerException.StackTrace);
                    innerException = innerException.InnerException;
                }
                MessageBox.Show(@"应用程序出现了一个未知的异常
请重启程序再试！

注意事项
 1，请注意保持文件的格式正确。
 2，请注意不要改动原始文件。
 3，请不要用其他工具打开要处理的文件。
                    
                    ");

              

                e.Handled = true;
                this.Shutdown();
            }
            catch
            {
            }
        }
    }
}
