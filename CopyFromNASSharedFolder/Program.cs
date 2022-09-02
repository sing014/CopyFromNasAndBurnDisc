using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Configuration.Install;
using System.Reflection;
using System.Collections;

namespace CopyFromNASSharedFolder
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        static void Main(params string[] saArg)
        {
            if(saArg.Length < 1)
            {
                ServiceBase[] ServicesToRun1;
                ServicesToRun1 = new ServiceBase[]
                {
                    new Service1()
                };
                ServiceBase.Run(ServicesToRun1);
            }
            else
            {
                SelfInstall(saArg);
            }
            
        }

        static void SelfInstall(params string[] args)
        {
            // 自我安裝程式碼

            using (AssemblyInstaller installer1 = new AssemblyInstaller())
            {
                IDictionary dict1 = new Hashtable();
                installer1.UseNewContext = true;
                installer1.Path = Assembly.GetCallingAssembly().Location;
                switch(args[0].ToLower())
                {
                    case "/i":
                    case "-i":
                        installer1.Install(dict1);
                        installer1.Commit(dict1);
                        break;
                    case "/u":
                    case "-u":
                        installer1.Uninstall(dict1);
                        break;
                    default:
                        installer1.Install(dict1);
                        installer1.Commit(dict1);
                        return;
                }
            }

        }

    }
}
