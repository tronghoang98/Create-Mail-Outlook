using Emgu.CV.Structure;
using Emgu.CV;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management;
using System.Runtime.InteropServices;
using System.IO;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.Net;
using Emgu.CV.Reg;
using System.Xml.Linq;

 namespace CreateMailOutlook.Common
{
    public static class Utils
    {
        public static string GetRegistry(string registryPath, string registryKey)
        {
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registryPath))
                {
                    string result = "";
                    if (key != null)
                    {
                        result = key.GetValue(registryKey).ToString();
                    }
                    return result;
                }
            }
            catch (Exception ex)
            {

                return ex.Message;
            }
        }
        public static string SetRegistry(string registryPath, string registryKey, string value)
        {
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(registryKey, true))
                {
                    if (key != null)
                    {
                        key.SetValue(registryKey, value, RegistryValueKind.String);

                    }
                    return "OK";
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public static string RandomString(int length)
        {
            Random random = new Random();
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";

            return new string(Enumerable.Repeat(chars, length)
                .Select(s => s[random.Next(s.Length)]).ToArray()).ToUpper();
        }
        public static void RunCMD(string cmd)
        {
            ProcessStartInfo psi = new ProcessStartInfo()
            {
                FileName = "cmd.exe",
                Arguments = $"/c {cmd}",
                Verb = "runas", // Chạy với quyền Admin
                UseShellExecute = true, // Bắt buộc phải bật khi dùng runas
                CreateNoWindow = false // Để hiển thị CMD
            };

            try
            {
                Process process = Process.Start(psi);
                process.WaitForExit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi: " + ex.Message);
            }
        }
       
        public static void ChangeInfo()
        {

            using (RegistryKey registry = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", true))
            {
                if (registry != null)
                {
                    registry.SetValue("ProductId", $"{RandomString(5)}-{RandomString(5)}-{RandomString(5)}-{RandomString(5)}", RegistryValueKind.String);

                    registry.SetValue("InstallDate", (new DateTimeOffset(2023,2,2,2,2,2,new TimeSpan())).ToUnixTimeSeconds(), RegistryValueKind.DWord);

                    registry.SetValue("ProductName", $"Windows 10", RegistryValueKind.String);
                }
              
            }
            using (RegistryKey registry = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\SQMClient", true))
            {
                if (registry != null)
                {
                    registry.SetValue("MachineId", $"{{{Guid.NewGuid().ToString().ToUpper()}}}", RegistryValueKind.String);

                   
                }

            }



           // proxy c1
            using (RegistryKey registry = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings", true))
            {
                if (registry != null)
                {
                    registry.SetValue("ProxyEnable", 1);
                    registry.SetValue("ProxyServer", "172.24.10.10");
                    //Console.WriteLine("Proxy đã được thiết lập với xác thực tự động từ Credential Manager.");
                    // RunCMD("cmdkey /add:proxy.example.com /user:admin /pass:123456");
                    System.Diagnostics.Process.Start("cmd.exe", "/C \"ipconfig /flushdns\"");

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
            string newMacAddress = $"{RandomString(2)}-{RandomString(2)}-{RandomString(2)}-{RandomString(2)}-{RandomString(2)}-{RandomString(2)}";

            // Tên của card mạng (có thể lấy từ "ipconfig")
            string interfaceName = "Ethernet";

          //  string command = $"interface ipv4 set address name=\"{interfaceName}\" static 192.168.1.100 255.255.255.0 192.168.1.1";
            string macCommand = $"interface set interface name=\"{interfaceName}\" mac={newMacAddress}";

            RunCMD(macCommand);


            using (RegistryKey registry = Registry.LocalMachine.OpenSubKey(@"HARDWARE\DESCRIPTION\System", true))
            {
                if (registry != null)
                {
                    registry.SetValue("SystemBiosVersion", "LENOVO - 1340\r\nR2AET59W(1.34)\r\nLenovo - 1340\r\n");
                }
            }

            
            using (RegistryKey registry = Registry.LocalMachine.OpenSubKey(@"HARDWARE\DESCRIPTION\System\BIOS", true))
            {
                if (registry != null)
                {
                    registry.SetValue("BaseBoardManufacturer", "LENOVO");
                    registry.SetValue("BaseBoardProduct", "21JK006HVA");
                    registry.SetValue("BIOSReleaseDate", "08/27/2024");
                    registry.SetValue("BIOSVendor", "LENOVO");
                    registry.SetValue("BIOSVersion", "R2AET59W(1.34)");
                    registry.SetValue("SystemFamily", "ThinkPad E14 Gen 5");
                    registry.SetValue("SystemManufacturer", "LENOVO");
                    registry.SetValue("SystemProductName", "21JK006HVA");
                    registry.SetValue("SystemSKU", "LENOVO_MT_21JK_BU_Think_FM_ThinkPad E14 Gen 5");
                    registry.SetValue("SystemVersion", "ThinkPad E14 Gen 5");

                }
            }
        }
        
        static void ConnectVPN(string vpnName, string username, string password)
        {
            try
            {
                // Tạo lệnh rasdial để kết nối với VPN
                string command = $"rasdial {vpnName} {username} {password}";

                // Khởi chạy quá trình với lệnh
                ProcessStartInfo psi = new ProcessStartInfo("cmd.exe", $"/C {command}")
                {
                    CreateNoWindow = true,
                    UseShellExecute = false
                };

                Process process = Process.Start(psi);
                process.WaitForExit();

                Console.WriteLine("Đã kết nối VPN thành công.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi khi kết nối VPN: " + ex.Message);
            }
        }
        public static string GetDeviceId()
        {
            string path = "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\SQMClient";
            string key = "MachineId";
            return GetRegistry(path, key);
        }
        public static string GetProductId()
        {
            string path = "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion";
            string key = "ProductId";
            return GetRegistry(path, key);
        }
        public static string GetBIOSSerial()
        {
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BIOS");
                return searcher.Get().Cast<ManagementObject>().FirstOrDefault()["SerialNumber"].ToString();
            }
            catch (Exception e)
            {
                return "";
            }

        }
        public static List<string> GetSerialDisk()
        {
            List<string> lstResult = new List<string>();
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_DiskDrive");
                foreach (ManagementObject wmi in searcher.Get())
                {
                    lstResult.Add(wmi["SerialNumber"].ToString());

                }
                return lstResult;
            }
            catch (Exception)
            {

                return lstResult;
            }
        }


        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;

        public static void LeftMouseClick(int xpos, int ypos)
        {
            SetCursorPos(xpos, ypos);
            mouse_event(MOUSEEVENTF_LEFTDOWN, xpos, ypos, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, xpos, ypos, 0, 0);
        }
        [DllImport("User32.dll")]
        private static extern bool ShowWindow(IntPtr handle, int nCmdShow);
        [DllImport("USER32.DLL")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        public static bool BringToFrontAPP(IntPtr handle)
        {
            if (handle == IntPtr.Zero)
            {
                return false;
            }

            bool flag = ShowWindow(handle, 1);
            bool flag2 = SetForegroundWindow(handle);
            return flag && flag2;
        }
        public static bool SenKeys(string content)
        {
            try
            {
                if (content.StartsWith('+') || content.StartsWith('^') || content.StartsWith('%') || content.StartsWith('{'))
                {
                    SendKeys.Send(content);
                }
                else
                {
                    char[] strArr = content.ToCharArray();
                    foreach (var item in strArr)
                    {
                        SendKeys.Send(item.ToString());
                        Thread.Sleep(200);
                    }
                }
                
                return true;
            }
            catch (Exception)
            {

                return false;
            }
        }

        public static void CloseProcess()
        {
            foreach (var process in Process.GetProcessesByName("olk"))
            {
                process.Kill();
            }
            Thread.Sleep(15000);
        }
        public static System.Drawing.Point GetPostionItem( Bitmap bitTemp,bool isSave=false)
        {
            System.Drawing.Point maxLoc = new System.Drawing.Point();
            double maxVal = 0.0;
            int count = 0;
            while (count<=15)
            {
                Bitmap bmpScreenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, PixelFormat.Format32bppArgb);
                Graphics gfxScreenshot = Graphics.FromImage(bmpScreenshot);
                gfxScreenshot.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);
                // bmpScreenshot.Save("a.jpeg", ImageFormat.Png);

                Image<Bgr, byte> temp = bitTemp.ToImage<Bgr, byte>();

                Image<Bgr, byte> source = bmpScreenshot.ToImage<Bgr, byte>();

                Mat imgOut = new Mat();
                //lấy nét hình ảnh

                //Mat imgOutTemp = new Mat();
                //CvInvoke.Canny(temp, imgOutTemp, 200, 0);
                //Mat imgOutSorc = new Mat();
                //CvInvoke.Canny(source, imgOutSorc, 200, 0);
                //tìm ảnh temp trong source
                CvInvoke.MatchTemplate(source, temp, imgOut, Emgu.CV.CvEnum.TemplateMatchingType.CcoeffNormed);


                //lấy vị trí rec tìm được
                double minVal = 0.0;
               
                System.Drawing.Point minLoc = new System.Drawing.Point();
                CvInvoke.MinMaxLoc(imgOut, ref minVal, ref maxVal, ref minLoc, ref maxLoc);
                if (maxVal >= 0.8)
                {
                    if (isSave)
                    {
                        //vẽ rec tìm được
                        System.Drawing.Rectangle rec = new System.Drawing.Rectangle(maxLoc, temp.Size);
                        var outI = source.Copy();
                        CvInvoke.Rectangle(source, rec, new MCvScalar(0, 0, 255), -1);
                        CvInvoke.AddWeighted(outI, 0.5, source, 1 - 0.5, 0, outI);
                        source = outI;

                        if (!Directory.Exists("result"))
                        {
                            Directory.CreateDirectory("result");
                        }
                        //lưu ảnh
                        source.Save($"result/1.jpg");
                    }
                    break;
                }
                Thread.Sleep(1000);
                count++;
            }
            if (maxVal >= 0.8)
            {
                return maxLoc;
            }
            else
            {
                return new Point();
            }
            
        }



    }

   public class InfoBios()
    {
        List<string> lstVendor = new List<string>() { "Dell", "HP", "Lenovo", "ASUS","Acer" };
        public string BIOSVendor { get; set; }

        private string BIOSVersion { get; set; }

        public string SystemFamily { get; set; }
        public string SystemManufacturer { get; set; }
        public string SystemProductName { get; set; }
        public string SystemSKU { get; set; }
        public string SystemVersion { get; set; }
        public string BaseBoardProduct { get; set; }
        public string BaseBoardManufacturer { get; set; }

        public InfoBios GetInfo()
        {
            Random random = new Random();
            var i = random.Next(1);
            InfoBios infoBios = new InfoBios();
            switch (lstVendor[i])
            {
                case "Dell":
                    infoBios.BIOSVendor = "Dell Inc.";
                    infoBios.BIOSVersion = $"A{random.Next(2)}";

                    break;
                case "HP":
                    infoBios.BIOSVendor = "HP";
                    infoBios.BIOSVersion = $"F{random.Next(2)}";
                    break;
                case "Lenovo":
                    infoBios.BIOSVendor = "HP";
                    infoBios.BIOSVersion = $"N{random.Next(2)}";
                    break;
                case "ASUS":
                    infoBios.BIOSVendor = "ASUS";
                    infoBios.BIOSVersion = $"X.{random.Next(2)}";
                    break;
                case "Acer":
                    infoBios.BIOSVendor = "Acer";
                    infoBios.BIOSVersion = $"Vx.{random.Next(2)}";
                    break;
            }
            return infoBios;
        }
    }
}
