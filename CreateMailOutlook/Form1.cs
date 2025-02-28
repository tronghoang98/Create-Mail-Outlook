using System.Diagnostics;
using System.Drawing.Imaging;

namespace CreateMailOutlook
{
    public partial class Form1 : Form
    {
        private string PATH_ROOT = "D:\\Desktop\\Image";

        private string PATH_MAIL = "C:\\Users\\u38811\\AppData\\Local\\Microsoft\\WindowsApps\\olk.exe";
        public Form1()
        {
            InitializeComponent();

        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            //start app maill
            // Process.Start("outlookmail:");
            while (true)
            {
                //  Process.Start(PATH_MAIL);
                var a = $"{{{Guid.NewGuid()}}}";
                while (true)
                {
                    try
                    {
                        var process = Process.GetProcessesByName("olk");
                        if (process[0].MainWindowHandle != IntPtr.Zero)
                        {
                            Common.Utils.BringToFrontAPP(process[0].MainWindowHandle);
                            break;
                        }
                        Thread.Sleep(1000);
                    }
                    catch (Exception)
                    {

                    }
                }


                Point createPos1 = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "1.jpeg")), true);
                if (createPos1.X == 0&&createPos1.Y == 0){
                    Common.Utils.CloseProcess();
                    continue;
                }
                Common.Utils.LeftMouseClick(createPos1.X+2, createPos1.Y+2);


                Point createPos2 = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "2.jpeg")), true);
                if (createPos2.X == 0 && createPos2.Y == 0)
                {
                    Common.Utils.CloseProcess();
                    continue;

                }

                Common.Utils.SenKeys("vuchuyen1222@hotmail.com");
                Thread.Sleep(1000);

                Point createPosNext = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "next.jpeg")), true);
                if (createPosNext.X == 0 && createPosNext.Y == 0)
                {
                    Common.Utils.CloseProcess();
                    continue;
                }
                Common.Utils.LeftMouseClick(createPosNext.X, createPosNext.Y);
                Thread.Sleep(1000);
                Point createPosPass = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "password.jpeg")), true);
                if (createPosPass.X == 0 && createPosPass.Y == 0)
                {
                    Common.Utils.CloseProcess();
                    continue;
                }
                Common.Utils.SenKeys("Hoang1998");
                Thread.Sleep(1000);

                Point createPosNext1 = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "next.jpeg")), true);
                if (createPosNext1.X == 0 && createPosNext1.Y == 0)
                {
                    Common.Utils.CloseProcess();
                    continue;
                }
                Common.Utils.LeftMouseClick(createPosNext1.X, createPosNext1.Y);
                Thread.Sleep(1000);

                Point createPosName = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "name.jpeg")), true);
                if (createPosName.X == 0 && createPosName.Y == 0)
                {
                    Common.Utils.CloseProcess();
                    continue;
                }
                Common.Utils.SenKeys("Ha");
                Thread.Sleep(1000);

                Common.Utils.SenKeys("{TAB}");
                Thread.Sleep(1000);

                Common.Utils.SenKeys("Trong Hoang");
                Thread.Sleep(1000);

                Point createPosNext2 = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "next.jpeg")), true);
                if (createPosNext2.X == 0 && createPosNext2.Y == 0)
                {
                    Common.Utils.CloseProcess();
                    continue;
                }
                Common.Utils.LeftMouseClick(createPosNext2.X, createPosNext2.Y);
                Thread.Sleep(1000);

                Point createPosBirthDate = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "bỉthdate.jpeg")), true);
                if (createPosBirthDate.X == 0 && createPosBirthDate.Y == 0)
                {
                    Common.Utils.CloseProcess();
                    continue;
                }


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

                Common.Utils.SenKeys("{TAB}");
                Thread.Sleep(1000);
                Common.Utils.SenKeys("1994");
                Thread.Sleep(1000);

                Point createPosNext3 = Common.Utils.GetPostionItem(new Bitmap(Path.Combine(PATH_ROOT, "next.jpeg")), true);
                if (createPosNext3.X == 0 && createPosNext3.Y == 0)
                {
                    Common.Utils.CloseProcess();
                    continue;
                }
                Common.Utils.LeftMouseClick(createPosNext3.X, createPosNext3.Y);
                Thread.Sleep(1000);




            }
        }
    }
}
