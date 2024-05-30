using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace sugarloaf_programming {
    class Program {
        private static Process proc = new Process();
        private static bool debug = false;
        private static System.Threading.Timer close_program;
        static void Main(string[] args) {
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += new UnhandledExceptionEventHandler(MyHandler);


            string head = "1";
            while (true)
            {
                try
                { head = File.ReadAllText("../../config/head.txt"); break; } catch { Thread.Sleep(50); }
            }
            FileDelete("../../config/head.txt");
            File.WriteAllText("call_exe_tric.txt", "");
            File.WriteAllText("sugarloaf_programming_running_" + head + ".txt", "");
            string hex_file = File.ReadAllText("../../config/sugarloaf_program_hex.txt");
            string e2lite = File.ReadAllText("../../config/sugarloaf_program_e2lite.txt");
            string name_project = "sugarloaf.rpj";
            try
            { name_project = File.ReadAllText("../../config/sugarloaf_program_name_project.txt"); } catch { }
            int timeout = Convert.ToInt32(File.ReadAllText("../../config/test_head_" + head + "_timeout.txt"));
            debug = Convert.ToBoolean(File.ReadAllText("../../config/test_head_" + head + "_debug.txt"));
            close_program = new System.Threading.Timer(TimerCallback, null, 0, timeout + 20000);
            try
            {
                File.WriteAllText("sugarloaf_program_" + head + ".bat",
                              "\"C:\\Program Files (x86)\\Renesas Electronics\\Programming Tools\\Renesas Flash Programmer " +
                              "V3.05\\RFPV3.exe\" /silent \"D:\\svn\\2020_SENSITECTH_SugarLoaf_Automation\\2.Design Documents\\" +
                              "7.Test Application Program\\Bat file for programming\\sugarloaf\\" + name_project + "\" " +
                              "/file \"D:\\svn\\2020_SENSITECTH_SugarLoaf_Automation\\1.Customer Documents\\7.Firmware\\" +
                              hex_file + "\" /tool " + e2lite + " /log \"sugarloaf_program_" + head + ".log");
            } catch { return; }
            Console.WriteLine("head = " + head);
            Console.WriteLine("hex_file = " + hex_file);
            Console.WriteLine("e2lite = " + e2lite);
            if (debug == true)
                Console.ReadKey();

            Stopwatch timeout_ = new Stopwatch();
            proc.StartInfo.WorkingDirectory = "";
            proc.StartInfo.CreateNoWindow = true;
            proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            proc.StartInfo.FileName = "sugarloaf_program_" + head + ".bat";
            bool flag_e2 = false;
            string data = "";
            for (int i = 1; i <= 2; i++)
            {
                FileDelete("sugarloaf_program_" + head + ".log");
                while (true)
                {
                    try
                    { proc.Start(); break; } catch { Thread.Sleep(50); }
                }
                data = "";
                Console.Write("wait log file");
                timeout_.Restart();
                while (timeout_.ElapsedMilliseconds < timeout)
                {
                    try
                    {
                        data = File.ReadAllText("sugarloaf_program_" + head + ".log");
                    } catch (Exception)
                    {
                        Console.Write(".");
                        Thread.Sleep(250);
                        continue;
                    }
                    timeout_.Stop();
                    break;
                }
                if (timeout_.IsRunning)
                { File.WriteAllText("test_head_" + head + "_result.txt", "timeout\r\nFAIL"); return; }
                bool flag_e2lite = true;
                if (!data.Contains("Verifying data"))
                { flag_e2lite = false; Console.WriteLine(""); Console.WriteLine("not Verifying data"); }
                else
                { Console.WriteLine(""); Console.WriteLine("Verifying data"); }
                if (!data.Contains("Operation completed."))
                { flag_e2lite = false; Console.WriteLine("not Operation completed."); }
                else
                    Console.WriteLine("Operation completed.");
                Console.WriteLine(hex_file);
                if (debug == true)
                    Console.ReadKey();
                if (flag_e2lite == true)
                { flag_e2 = true; break; }
                wait_discom(head);
            }
            FileDelete("sugarloaf_program_" + head + ".bat");
            data = data.Replace("'", "").Replace(",", "").Replace("\n", "").Replace("\r", "");
            if (flag_e2 == true)
                File.WriteAllText("test_head_" + head + "_result.txt", hex_file + "\r\nPASS");
            else
                File.WriteAllText("test_head_" + head + "_result.txt", data + "\r\nFAIL");
            FileDelete("sugarloaf_programming_running_" + head + ".txt");
            FileDelete("sugarloaf_programming_discom_" + head + ".txt");
        }

        private static bool flag_close = false;
        private static void TimerCallback(Object o) {
            if (!flag_close)
            { flag_close = true; return; }
            if (debug)
                return;
            if (flag_close)
                Environment.Exit(0);
        }
        private static void wait_discom(string head) {
            File.WriteAllText("sugarloaf_programming_discom_" + head + ".txt", "");
            List<string> head_all = new List<string>();
            for (int kj = 1; kj <= 4; kj++)
            {
                head_all.Add(kj.ToString());
            }
            head_all.Remove(head);
            int minminmin = 10;
            while (true)
            {
                bool flag_running = false;
                bool flag_discom = false;
                bool flag_discom_run = true;
                foreach (string xc in head_all)
                {
                    try
                    {
                        File.ReadAllText("sugarloaf_programming_running_" + xc + ".txt");
                        flag_running = true;
                    } catch { }
                    try
                    {
                        File.ReadAllText("sugarloaf_programming_discom_" + xc + ".txt");
                        if (minminmin > Convert.ToInt32(xc))
                            minminmin = Convert.ToInt32(xc);
                        flag_discom = true;
                    } catch { }
                    if (flag_running && !flag_discom)
                        flag_discom_run = false;
                }
                if (flag_discom_run)
                    break;
                Thread.Sleep(50);
            }
            if (Convert.ToInt32(head) < minminmin)
            {
                get_name_e2();
                discom("disable");
                discom("enable");
            }
            else
                Thread.Sleep(2500);
        }

        private static void get_name_e2() {
            ManagementObjectSearcher objOSDetails2 =
               new ManagementObjectSearcher(@"SELECT * FROM Win32_PnPEntity where DeviceID Like ""USB%""");
            ManagementObjectCollection osDetailsCollection2 = objOSDetails2.Get();
            foreach (ManagementObject usblist in osDetailsCollection2)
            {
                string arrport = usblist.GetPropertyValue("NAME").ToString();
                if (arrport.Contains("Renesas"))
                {
                    name_e2 = arrport;
                }
            }
        }
        private static string name_e2 = "Renesas E2 Lite";
        private static void discom(string cmd) {//enable disable//
            Process devManViewProc = new Process();
            devManViewProc.StartInfo.FileName = "DevManView.exe";
            devManViewProc.StartInfo.Arguments = "/" + cmd + " \"" + name_e2 + "\"";
            devManViewProc.Start();
            devManViewProc.WaitForExit();
        }


        /// <summary>
        /// Log program catch to csv file
        /// </summary>
        /// <param name="text"></param>
        private static void LogProgramCatch(string text) {
            string path = "D:\\LogError\\SugarloafProgramCatch";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            DateTime now = DateTime.Now;
            StreamWriter swOut = new StreamWriter(path + "\\" + now.Year + "_" + now.Month + ".csv", true);
            string time = now.Day.ToString("00") + ":" + now.Hour.ToString("00") + ":" + now.Minute.ToString("00") + ":" + now.Second.ToString("00");
            swOut.WriteLine(time + "," + text);
            swOut.Close();
        }

        /// <summary>
        /// Event Exception Catch Program
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void MyHandler(object sender, UnhandledExceptionEventArgs args) {
            Exception e = (Exception)args.ExceptionObject;
            LogProgramCatch(e.StackTrace);
        }

        /// <summary>
        /// For while loop delete file
        /// </summary>
        /// <param name="path"></param>
        private static void FileDelete(string path) {
            while (true)
            {
                try
                {
                    File.Delete(path);
                    break;
                } catch
                {
                    Thread.Sleep(50);
                }
            }
        }
    }
}
