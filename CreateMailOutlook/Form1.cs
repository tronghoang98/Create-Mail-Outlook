using CreateMailOutlook.Common;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.Net.NetworkInformation;
using System.Reflection.PortableExecutable;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace CreateMailOutlook
{
    public partial class Form1 : Form
    {
        private string PATH_ROOT = "image";
        private Thread thread;

        private string PATH_MAIL = "C:\\Program Files\\WindowsApps\\Microsoft.OutlookForWindows_1.2025.219.400_x64__8wekyb3d8bbwe\\olk.exe";
        public Form1()
        {
            InitializeComponent();
             GetInfo();

        }
        public void WriteLog(string content)
        {
            this.Invoke(new Action(() =>
            {
                txtLog.AppendText(content + "\n");
            }));
        }
        public void GetInfo()
        {
            try
            {
                using (RegistryKey registry = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", true))
                {
                    if (registry != null)
                    {
                        txtProductId.Text = registry.GetValue("ProductId").ToString();

                        txtInstallDate.Text = registry.GetValue("InstallDate").ToString();

                        txtProductName.Text = registry.GetValue("ProductName").ToString();
                    }

                }
                using (RegistryKey registry = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\SQMClient", true))
                {
                    if (registry != null)
                    {
                        txtMachineId.Text = registry.GetValue("MachineId").ToString();

                    }

                }
                NetworkInterface[] networkInterfaces = NetworkInterface.GetAllNetworkInterfaces();
                txtMAC.Text = networkInterfaces.Where(w => w.OperationalStatus == OperationalStatus.Up && !string.IsNullOrEmpty(w.GetPhysicalAddress().ToString())).First().GetPhysicalAddress().ToString();
            }
            catch (Exception)
            {

            }
        }
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        const uint SWP_NOSIZE = 0x0001;
        const uint SWP_NOZORDER = 0x0004;
        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtPathOut.Text))
            {
                PATH_MAIL = txtPathOut.Text;
            }
            thread = new Thread(() =>
           {
               //start app maill
               // Process.Start("outlookmail:");
               while (true)
               {
                   try
                   {
                       using (RegistryKey registry = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", true))
                       {
                           if (registry != null)
                           {
                               var proValue = $"{Common.Utils.RandomString(5)}-{Common.Utils.RandomString(5)}-{Common.Utils.RandomString(5)}-{Common.Utils.RandomString(5)}";
                               registry.SetValue("ProductId", proValue, RegistryValueKind.String);
                               WriteLog("Set ProductId : " + proValue);
                               var insValue = (new DateTimeOffset(2023, 2, 2, 2, 2, 2, new TimeSpan())).ToUnixTimeSeconds();
                               registry.SetValue("InstallDate", insValue, RegistryValueKind.DWord);
                               WriteLog("Set InstallDate: " + insValue.ToString());
                               var pronameValue = "Windows 10";
                               registry.SetValue("ProductName", pronameValue, RegistryValueKind.String);
                               WriteLog("Set ProductName: " + pronameValue.ToString());
                           }

                       }


                       using (RegistryKey registry = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\SQMClient", true))
                       {
                           if (registry != null)
                           {
                               var machine = $"{{{Guid.NewGuid().ToString().ToUpper()}}}";
                               registry.SetValue("MachineId", machine, RegistryValueKind.String);
                               WriteLog("Set MachineId: " + machine.ToString());
                           }

                       }


                       // proxy c1
                       if (!string.IsNullOrEmpty(txtProxy.Text))
                       {
                           using (RegistryKey registry = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings", true))
                           {
                               if (registry != null)
                               {
                                   registry.SetValue("ProxyEnable", 1);
                                   registry.SetValue("ProxyServer", txtProxy.Text);
                                   //Console.WriteLine("Proxy đã được thiết lập với xác thực tự động từ Credential Manager.");
                                   // RunCMD("cmdkey /add:proxy.example.com /user:admin /pass:123456");
                                   System.Diagnostics.Process.Start("cmd.exe", "/C \"ipconfig /flushdns\"");
                                   WriteLog("Set Proxy: " + txtProxy.Text);
                               }
                           }
                       }

                       #region proxy c2
                       //proxy c2

                       //string key = @"Software\Microsoft\Windows\CurrentVersion\Internet Settings";
                       //using (RegistryKey registry = Registry.CurrentUser.OpenSubKey(key, true))
                       //{
                       //    string pacUrl = "file:///"+Path.Combine(Application.ExecutablePath,"proxy/proxy.pac");
                       //    if (registry != null)
                       //    {
                       //        registry.SetValue("AutoConfigURL", pacUrl);
                       //        registry.SetValue("ProxyEnable", 1);
                       //        Console.WriteLine("Proxy PAC đã được thiết lập!");
                       //    }
                       //}
                       #endregion

                       //set mac
                       string newMacAddress = $"{Common.Utils.RandomString(2)}-{Common.Utils.RandomString(2)}-{Common.Utils.RandomString(2)}-{Common.Utils.RandomString(2)}-{Common.Utils.RandomString(2)}-{Common.Utils.RandomString(2)}";

                       // Tên của card mạng (có thể lấy từ "ipconfig")
                       string interfaceName = "Ethernet";

                       //  string command = $"interface ipv4 set address name=\"{interfaceName}\" static 192.168.1.100 255.255.255.0 192.168.1.1";
                       string macCommand = $"interface set interface name=\"{interfaceName}\" mac={newMacAddress}";

                       Common.Utils.RunCMD(macCommand);

                       WriteLog("Set Mac: " + newMacAddress);

                       GetInfo();
                   }
                   catch (Exception)
                   {

                   }
                   Random rm = new Random();
                   List<string> lstHo = File.ReadAllLines("data/ho.txt").ToList();
                   List<string> lstDem = File.ReadAllLines("data/dem.txt").ToList();
                   List<string> lstTen = File.ReadAllLines("data/dem.txt").ToList();
                   string ho = lstHo[rm.Next(9)];
                   string ten = $"{lstDem[rm.Next(9)]} {lstTen[rm.Next(9)]}";
                   string mail = Regex.Replace(Utils.ToVietNameseChacracter((ho.Trim() + ten.Trim() + Utils.RandomString(4))).ToLower(), "(\\s+)|([^\\w])", "");
                   string pass = Common.Utils.RandomString(8);
                   WriteLog("Mở app..");
                   Process.Start(PATH_MAIL);
                   while (true)
                   {
                       try
                       {

                           var process = Process.GetProcessesByName("olk");
                           if (process[0].MainWindowHandle != IntPtr.Zero)
                           {
                               Common.Utils.BringToFrontAPP(process[0].MainWindowHandle);
                               SetWindowPos(process[0].MainWindowHandle, IntPtr.Zero, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
                               break;
                           }
                           Thread.Sleep(1000);
                       }
                       catch (Exception)
                       {

                       }
                   }

                   WriteLog("Click tạo tk...");
                   Point createPos1 = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "1.jpeg")), true);
                   if (createPos1.X == 0 && createPos1.Y == 0)
                   {
                       Common.Utils.CloseProcess();
                       continue;
                   }
                   Common.Utils.LeftMouseClick(createPos1.X + 2, createPos1.Y + 2);


                   Point createPos2 = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "2.jpeg")), true);
                   if (createPos2.X == 0 && createPos2.Y == 0)
                   {
                       Common.Utils.CloseProcess();
                       continue;

                   }
                   WriteLog("Nhập mail...");
                   this.Invoke(new Action(() =>
                   {
                       Common.Utils.SenKeys(mail);
                   }));
                   Thread.Sleep(1000);

                   WriteLog("Chọn miền ...");
                   Common.Utils.LeftMouseClick(createPos2.X, createPos2.Y);
                   Thread.Sleep(1000);

                   Common.Utils.LeftMouseClick(createPos2.X, createPos2.Y + 25);
                   Thread.Sleep(1000);

                   WriteLog("Tiếp theo...");
                   Point createPosNext = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "next.jpeg")), true);
                   if (createPosNext.X == 0 && createPosNext.Y == 0)
                   {
                       Common.Utils.CloseProcess();
                       continue;
                   }
                   Common.Utils.LeftMouseClick(createPosNext.X, createPosNext.Y);
                   Thread.Sleep(1000);

                   WriteLog("Nhập pass...");
                   Point createPosPass = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "password.jpeg")), true);
                   if (createPosPass.X == 0 && createPosPass.Y == 0)
                   {
                       Common.Utils.CloseProcess();
                       continue;
                   }
                   this.Invoke(new Action(() =>
                   {
                       Common.Utils.SenKeys(pass);
                   }));
                   Thread.Sleep(1000);

                   WriteLog("Tiếp theo...");
                   Point createPosNext1 = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "next.jpeg")), true);
                   if (createPosNext1.X == 0 && createPosNext1.Y == 0)
                   {
                       Common.Utils.CloseProcess();
                       continue;
                   }
                   Common.Utils.LeftMouseClick(createPosNext1.X, createPosNext1.Y);
                   Thread.Sleep(1000);

                   WriteLog("Nhập họ tên...");
                   Point createPosName = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "name.jpeg")), true);
                   if (createPosName.X == 0 && createPosName.Y == 0)
                   {
                       Common.Utils.CloseProcess();
                       continue;
                   }
                   this.Invoke(new Action(() =>
                   {

                       Common.Utils.SenKeys(ho);
                   }));
                   Thread.Sleep(1000);

                   this.Invoke(new Action(() => { Common.Utils.SenKeys("{TAB}"); }));
                   Thread.Sleep(1000);

                   this.Invoke(new Action(() => { Common.Utils.SenKeys(ten); }));
                   Thread.Sleep(1000);


                   WriteLog("Tiếp theo...");
                   Point createPosNext2 = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "next.jpeg")), true);
                   if (createPosNext2.X == 0 && createPosNext2.Y == 0)
                   {
                       Common.Utils.CloseProcess();
                       continue;
                   }
                   Common.Utils.LeftMouseClick(createPosNext2.X, createPosNext2.Y);
                   Thread.Sleep(1000);

                   Point createPosBirthDate = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "birthdate.jpeg")), true);
                   if (createPosBirthDate.X == 0 && createPosBirthDate.Y == 0)
                   {
                       Common.Utils.CloseProcess();
                       continue;
                   }

                   WriteLog("Chọn tháng sinh...");
                   Point createPosMonth = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "month.jpeg")), true);
                   if (createPosMonth.X == 0 && createPosMonth.Y == 0)
                   {
                       Common.Utils.CloseProcess();
                       continue;
                   }
                   Common.Utils.LeftMouseClick(createPosMonth.X, createPosMonth.Y);
                   Thread.Sleep(1000);
                   Common.Utils.LeftMouseClick(createPosMonth.X, createPosMonth.Y);
                   Thread.Sleep(1000);

                   WriteLog("Chọn ngày sinh...");

                   Point createPosDay = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "day.jpeg")), true);
                   if (createPosDay.X == 0 && createPosDay.Y == 0)
                   {
                       Common.Utils.CloseProcess();
                       continue;
                   }
                   Common.Utils.LeftMouseClick(createPosDay.X, createPosDay.Y);
                   Thread.Sleep(1000);
                   Common.Utils.LeftMouseClick(createPosDay.X, createPosDay.Y);
                   Thread.Sleep(1000);
                   this.Invoke(new Action(() =>
                   {

                       Common.Utils.SenKeys("{TAB}");
                   }));
                   Thread.Sleep(1000);
                   WriteLog("Nhập năm sinh...");

                   this.Invoke(new Action(() => { Common.Utils.SenKeys("1994"); }));
                   Thread.Sleep(1000);

                   WriteLog("Tiếp theo...");
                   Point createPosNext3 = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "next.jpeg")), true);
                   if (createPosNext3.X == 0 && createPosNext3.Y == 0)
                   {
                       Common.Utils.CloseProcess();
                       continue;
                   }
                   Common.Utils.LeftMouseClick(createPosNext3.X, createPosNext3.Y);
                   Thread.Sleep(1000);

               }
           });
            thread.Start();
            btnStop.Enabled = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            thread.Resume();
            thread.Abort();
            thread = null;
            btnStart.Enabled = true;
        }
    }
}
