using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.ServiceProcess;
using System.Threading;

namespace IISMRestartService
{
    public partial class Service1 : ServiceBase
    {
        bool running;
        string appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        Thread IISMRestThread;

        //Dictonary of days of the week
        Dictionary<string, DayOfWeek> days = new Dictionary<string, DayOfWeek>
                {
                    {"mon", DayOfWeek.Monday},
                    {"tue", DayOfWeek.Tuesday},
                    {"wed", DayOfWeek.Wednesday},
                    {"thu", DayOfWeek.Thursday},
                    {"fri", DayOfWeek.Friday},
                    {"sat", DayOfWeek.Saturday},
                    {"sun", DayOfWeek.Sunday}
                };

        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Thread.CurrentThread.Name = "MainThread";
            IISMRestThread = new Thread(new ThreadStart(IISMRestart));
            IISMRestThread.Name = "IISMRestThread";
            IISMRestThread.Start();
        }

        private void IISMRestart()
        {
            try
            {
                running = true;

                if (!Directory.Exists(appPath + "\\Logs\\"))
                {
                    Directory.CreateDirectory(appPath + "\\Logs\\");
                }

                //get the date and time from the Config file
                string weekday = ConfigurationManager.AppSettings["DayOfReset"];
                string time = ConfigurationManager.AppSettings["TimeOfReset"];

                //Get the day of the week from the Dictonary using the Config file
                DayOfWeek dayOfWeek = days[weekday.ToLower()];

                // Parse the time using TimeSpan
                TimeSpan timeOfDay = TimeSpan.Parse(time);

                // Get the current date
                DateTime currentDateTime = DateTime.Now;

                // Calculate the target date and time
                DateTime targetDateTime = currentDateTime.Date.AddDays((int)dayOfWeek - (int)currentDateTime.DayOfWeek).Add(timeOfDay);

                // Restart IIS Once at the start and then every Day at the time specified in the Config file
                if (targetDateTime > currentDateTime) targetDateTime = targetDateTime.AddDays(-7);

                while (running)
                {
                    ProcessStartInfo iisResetInfo = new ProcessStartInfo("iisreset.exe");
                    Process iisReset = new Process();
                    currentDateTime = DateTime.Now;

                    if (currentDateTime > targetDateTime)
                    {

                        try
                        {
                            int n = 3;

                            do
                            {
                                iisResetInfo.RedirectStandardOutput = true;
                                iisResetInfo.UseShellExecute = false;
                                iisResetInfo.CreateNoWindow = true;
                                iisReset.StartInfo = iisResetInfo;
                                iisReset.Start();
                                iisReset.WaitForExit();
                            } while (n-- > 0 && iisReset.ExitCode != 0);
                        }
                        catch (Exception ex)
                        {
                            Logger.WriteErrorLog(ex.Message);

                        }
                        finally
                        {
                            if (iisReset.ExitCode != 0)
                            {
                                Logger.WriteErrorLog("IIS Restart Failed at " + DateTime.Now);
                            }
                        }
                        targetDateTime = targetDateTime.AddDays(7);

                        Logger.WriteDebugLog("IIS Restarted at " + DateTime.Now);
                        Logger.WriteDebugLog("IIS will restarted at " + targetDateTime);

                        Thread.Sleep(300 * 1000);
                    }
                    Thread.Sleep(30 * 1000);
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }
        }

        protected override void OnStop()
        {
            running = false;
            try
            {
                if (IISMRestThread != null && IISMRestThread.IsAlive)
                {
                    IISMRestThread.Abort();
                }
            }
            catch (Exception ex)
            {
                Logger.WriteErrorLog(ex.Message);
            }

        }

        public void OnDebug()
        {
            OnStart(null);
        }
    }
}
