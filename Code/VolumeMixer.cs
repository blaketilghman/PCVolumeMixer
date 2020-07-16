using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.Remoting;
using System.Management;
using AudioSwitcher.AudioApi.CoreAudio;
using AudioSwitcher.AudioApi;

namespace VolumeMixer
{

    public partial class Form1 : Form
    {
        //public partial class App : Application, ISingleInstanceApp;

        //public string nircmdStart = Application.StartupPath + @"\nircmdc.exe";

        private KeyHandler ghk;

        CoreAudioDevice defaultPlaybackDevice = new CoreAudioController().DefaultPlaybackDevice;

        public static string selectedApp0 = "placeholder.exe";
        public static string selectedApp1 = "placeholder.exe";
        public static string selectedApp2 = "placeholder.exe";
        public static string selectedApp3 = "placeholder.exe";

        private const int APPCOMMAND_VOLUME_MUTE = 0x80000;
        private const int APPCOMMAND_VOLUME_UP = 0xA0000;
        private const int APPCOMMAND_VOLUME_DOWN = 0x90000;
        private const int WM_APPCOMMAND = 0x319;

        public static string strSelectedApp0;
        public static string strSelectedApp1;
        public static string strSelectedApp2;
        public static string strSelectedApp3;
        public static string strSelectedAppID0;
        public static string strSelectedAppID1;
        public static string strSelectedAppID2;
        public static string strSelectedAppID3;
        // these = MAX_INT so that nothing changes by default
        public static int intSelectedAppID0 = int.MaxValue;
        public static int intSelectedAppID1 = int.MaxValue;
        public static int intSelectedAppID2 = int.MaxValue;
        public static int intSelectedAppID3 = int.MaxValue;
        public static int volInitRead0;
        public static int volInitRead1;
        public static int volInitRead2;
        public static int volInitRead3;

        public bool customizingButtons = false;

        public int cbNumItems;

        //internal static IMMDeviceEnumerator deviceEnumerator;

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, IntPtr extraInfo);
        public const int KEYEVENTF_EXTENTEDKEY = 1;
        public const int KEYEVENTF_KEYUP = 0;
        public const int VK_MEDIA_NEXT_TRACK = 0xB0;// code to jump to next track
        public const int VK_MEDIA_PLAY_PAUSE = 0xB3;// code to play or pause a song
        public const int VK_MEDIA_PREV_TRACK = 0xB1;// code to jump to prev track

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string strClassName, string strWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);


        public static string deviceID;

        //string[] sortedProcs = Process.GetProcesses().Select<Process, string>(F => F.ProcessName + "- " + F.Id).ToArray();
        //string[] sortedProcsId = Process.GetProcesses().Select<Process, string>(F => "" + F.Id).ToArray();

        public const int SW_SHOWNORMAL = 1;
        [DllImport("user32.dll")]

        public static extern IntPtr ShowWindow(IntPtr hwnd, int nCmdShow);


        public static Process RunningInstance()
        {
            Process current = Process.GetCurrentProcess();

            Process[] processes = Process.GetProcessesByName(current.ProcessName);

            //Loop through the running processes in with the same name 
            foreach (Process process in processes)
            {
                //Ignore the current process 
                if (process.Id != current.Id)
                {
                    //Make sure that the process is running from the exe file. 
                    if (Assembly.GetExecutingAssembly().Location.Replace("/", "\\") == current.MainModule.FileName)
                    {
                        ShowWindow(process.MainWindowHandle, SW_SHOWNORMAL);
                        //Return the other process instance. 
                        return process;
                    }
                }
            }
            //No other instance was found, return null. 
            return null;
        }

        private string AutodetectArduinoPort()
        {
            ManagementScope connectionScope = new ManagementScope();
            SelectQuery serialQuery = new SelectQuery("SELECT * FROM Win32_SerialPort");
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(connectionScope, serialQuery);

            try
            {
                foreach (ManagementObject item in searcher.Get())
                {
                    string desc = item["Description"].ToString();
                    string deviceId = item["DeviceID"].ToString();

                    if (desc.Contains("Arduino"))
                    {
                        deviceID = deviceId;
                        Console.WriteLine(deviceId);
                        return deviceId;
                    }
                }
            }
            catch (ArgumentNullException e)
            {
                MessageBox.Show("Please make sure that the Volume Mixer is connected.");
            }

            return null;
            
        }
        public static SerialPort activeSerial;
        public Form1()
        {
            InitializeComponent();

            //ghk = new KeyHandler(Keys.Scroll, this);

            //ghk.Register();

            //comboBox1.Items.Clear();
            //Process[] MyProcess = Process.GetProcesses();
            //Array.Sort(MyProcess);
            
            // dump all audio devices
            foreach (AudioDevice device in AudioUtilities.GetAllDevices())
            {
                Console.WriteLine(device.FriendlyName);
            }

            // dump all audio sessions
            foreach (AudioSession session in AudioUtilities.GetAllSessions())
            {
                if (session.Process != null)
                {
                    // only the one associated with a defined process
                    Console.WriteLine(session.Process.ProcessName);
                }
            }
        }
        private void HandleHotkey()
        {
            // Do stuff...
            Show();
        }
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == Constants.WM_HOTKEY_MSG_ID)
                HandleHotkey();
            base.WndProc(ref m);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            /*if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                //Hide();
            }
            */
            if (e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult result = MessageBox.Show("Do you really want to exit?", "Volume Mixer Closing", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    Application.Exit();
                    
                }
                else
                {
                    e.Cancel = true;
                }
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void OnKeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            if (e.Alt && e.KeyCode == Keys.M)
            {
                WindowState = FormWindowState.Normal;
            }
        }

        public static Process[] MyProcess = Process.GetProcesses();

        private void Form1_Load(object sender, EventArgs e)
        {
            StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size(267, 452);
            gbButton1.Location = new Point(12, 10);
            gbButton2.Location = new Point(12, 100);
            gbButton3.Location = new Point(12, 200);
            gbButton4.Location = new Point(12, 300);
            AutodetectArduinoPort();
            activeSerial = new SerialPort(deviceID, 115200, Parity.None, 8, StopBits.One);
            RunningInstance();
            connectToArduino();

            foreach (AudioDevice device in AudioUtilities.GetAllDevices())
            {
                Console.WriteLine(device.FriendlyName);
            }

            // dump all audio sessions
            foreach (AudioSession session in AudioUtilities.GetAllSessions())
            {
                if (session.Process != null)
                {
                    // only the one associated with a defined process
                    Console.WriteLine(session.Process.ProcessName + " " + session.ProcessId + " - " + session.Process.MainWindowTitle);
                }
            }
            notifyIcon1.Visible = false;

            cbSelectApp0.Items.Clear();
            foreach (AudioSession session in AudioUtilities.GetAllSessions())
            {
                if (session.Process != null) //&& session.Process.MainWindowTitle != "")
                {
                    if (session.Process.ProcessName.Contains("Discord"))
                    {
                        if (session.Process.MainWindowTitle == "")
                        {
                            cbSelectApp0.Items.Add($"{session.Process.ProcessName}.exe (Voice):{session.Process.Id}");
                            cbSelectApp1.Items.Add($"{session.Process.ProcessName}.exe (Voice):{session.Process.Id}");
                            cbSelectApp2.Items.Add($"{session.Process.ProcessName}.exe (Voice):{session.Process.Id}");
                            cbSelectApp3.Items.Add($"{session.Process.ProcessName}.exe (Voice):{session.Process.Id}");
                        }
                        else if (!session.Process.MainWindowTitle.Contains(":"))
                        {
                            cbSelectApp0.Items.Add($"{session.Process.ProcessName}.exe (TTS):{session.Process.Id}");
                            cbSelectApp1.Items.Add($"{session.Process.ProcessName}.exe (TTS):{session.Process.Id}");
                            cbSelectApp2.Items.Add($"{session.Process.ProcessName}.exe (TTS):{session.Process.Id}");
                            cbSelectApp3.Items.Add($"{session.Process.ProcessName}.exe (TTS):{session.Process.Id}");
                        }
                    }
                    else
                    {
                        cbSelectApp0.Items.Add($"{session.Process.ProcessName}.exe:{session.Process.Id}");
                        cbSelectApp1.Items.Add($"{session.Process.ProcessName}.exe:{session.Process.Id}");
                        cbSelectApp2.Items.Add($"{session.Process.ProcessName}.exe:{session.Process.Id}");
                        cbSelectApp3.Items.Add($"{session.Process.ProcessName}.exe:{session.Process.Id}");
                    }
                }
            }
            cbNumItems = cbSelectApp0.Items.Count;
            //Console.WriteLine(cbNumItems + "!!!!!!!!!!!!!!!!!!!!!!");

        }

        private SerialPort connectToArduino()
        {
            //SerialPort activeSerial = new SerialPort("com3", 9600, Parity.None, 8, StopBits.One); //Change COM4 to your arduino port
            activeSerial.ReadBufferSize = 128;

            //**Holy Grail for the Pro Micro**
            activeSerial.DtrEnable = true;
            activeSerial.RtsEnable = true;
            //********************************

            activeSerial.Open();
            activeSerial.DataReceived += new SerialDataReceivedEventHandler(activeSerial_DataReceived);
            

            return activeSerial;
        }
        void activeSerial_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {

            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 500;

            //SerialPort serial = new SerialPort("com3", 9600, Parity.None, 8, StopBits.One);
            var dataFromArduino = activeSerial.ReadLine();
            
            Console.WriteLine(dataFromArduino);

            

            if (dataFromArduino.StartsWith("A"))
            {
                string strVol = dataFromArduino.Substring(dataFromArduino.IndexOf(":") + 1);
                float intVol = Convert.ToInt64(strVol);
                if (dataFromArduino.StartsWith("A0"))
                {
                    if (chkSysVol.Checked == true)
                    {
                        intSelectedAppID0 = 0;
                        defaultPlaybackDevice.SetVolumeAsync(intVol);

                    }
                    else
                    {
                        SetApplicationVolume(intSelectedAppID0, intVol);
                    }
                }
                if (dataFromArduino.StartsWith("A1"))
                {
                    SetApplicationVolume(intSelectedAppID1, intVol);
                }
                if (dataFromArduino.StartsWith("A2"))
                {
                    SetApplicationVolume(intSelectedAppID2, intVol);
                }
                if (dataFromArduino.StartsWith("A3"))
                {
                    SetApplicationVolume(intSelectedAppID3, intVol);
                }

            }

            //-------------------------------------------------------
            //TODO: Handle error when index is wrong

            //0
            if (dataFromArduino.StartsWith("LP0"))
            {
                this.Invoke((MethodInvoker)delegate ()
                {
                    if (cbSelectApp0.SelectedItem.ToString() == "")
                    {
                        cbSelectApp0.SelectedIndex = 0;
                    }
                    if (Convert.ToInt64(cbSelectApp0.SelectedIndex) - 1 >= 0)
                    {
                        cbSelectApp0.SelectedIndex -= 1;
                    }
                });
            }
            if (dataFromArduino.StartsWith("NP0"))
            {
                this.Invoke((MethodInvoker)delegate ()
                {
                    if (cbSelectApp0.Text.ToString() == "")
                    {
                        cbSelectApp0.SelectedIndex = 0;

                    }
                    if (Convert.ToInt64(cbSelectApp0.SelectedIndex) + 1 < cbNumItems)
                    {
                        cbSelectApp0.SelectedIndex ++;
                    }
                });
            }

            //1
            if (dataFromArduino.StartsWith("LP1"))
            {
                this.Invoke((MethodInvoker)delegate ()
                {
                    if (cbSelectApp1.Text.ToString() == "")
                    {
                        cbSelectApp1.SelectedIndex = 0;
                    }
                    else if (Convert.ToInt64(cbSelectApp1.SelectedIndex) - 1 >= 0)
                    {
                        cbSelectApp1.SelectedIndex -= 1;
                    }
                });
            }
            if (dataFromArduino.StartsWith("NP1"))
            {
                this.Invoke((MethodInvoker)delegate ()
                {
                    if (cbSelectApp1.Text.ToString() == "")
                    {
                        cbSelectApp1.SelectedIndex = 0;

                    }
                    else if (Convert.ToInt64(cbSelectApp1.SelectedIndex) + 1 < cbNumItems)
                    {
                        cbSelectApp1.SelectedIndex++;
                    }
                });
            }

            //2
            if (dataFromArduino.StartsWith("LP2"))
            {
                this.Invoke((MethodInvoker)delegate ()
                {
                    if (cbSelectApp2.Text.ToString() == "")
                    {
                        cbSelectApp2.SelectedIndex = 0;
                    }
                    else if (Convert.ToInt64(cbSelectApp2.SelectedIndex) - 1 >= 0)
                    {
                        cbSelectApp2.SelectedIndex -= 1;
                    }
                });
            }
            if (dataFromArduino.StartsWith("NP2"))
            {
                this.Invoke((MethodInvoker)delegate ()
                {
                    if (cbSelectApp2.Text.ToString() == "")
                    {
                        cbSelectApp2.SelectedIndex = 0;

                    }
                    else if (Convert.ToInt64(cbSelectApp2.SelectedIndex) + 1 < cbNumItems)
                    {
                        cbSelectApp2.SelectedIndex++;
                    }
                });
            }

            //3
            if (dataFromArduino.StartsWith("LP3"))
            {
                this.Invoke((MethodInvoker)delegate ()
                {
                    if (cbSelectApp3.Text.ToString() == "")
                    {
                        cbSelectApp3.SelectedIndex = 0;
                    }
                    else if (Convert.ToInt64(cbSelectApp3.SelectedIndex) - 1 >= 0)
                    {
                        cbSelectApp3.SelectedIndex -= 1;
                    }
                });
            }
            if (dataFromArduino.StartsWith("NP3"))
            {
                this.Invoke((MethodInvoker)delegate ()
                {
                    if (cbSelectApp3.Text.ToString() == "")
                    {
                        cbSelectApp3.SelectedIndex = 0;

                    }
                    else if (Convert.ToInt64(cbSelectApp3.SelectedIndex) + 1 < cbNumItems)
                    {
                        cbSelectApp3.SelectedIndex++;
                    }
                });
            }

            if (dataFromArduino.StartsWith("B0"))
            {
                if (radbutPrev0.Checked == true)
                {
                    // Jump to previous track
                    keybd_event(VK_MEDIA_PREV_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                }
                if (radbutPauPlay0.Checked == true)
                {
                    // Play or Pause music
                    keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                }
                if (radbutNext0.Checked == true)
                {
                    // Jump to next track
                    keybd_event(VK_MEDIA_NEXT_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                }


            }
            if (dataFromArduino.StartsWith("B1"))
            {
                if (radbutPrev1.Checked == true)
                {
                    // Jump to previous track
                    keybd_event(VK_MEDIA_PREV_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                }
                if (radbutPauPlay1.Checked == true)
                {
                    // Play or Pause music
                    keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                }
                if (radbutNext1.Checked == true)
                {
                    // Jump to next track
                    keybd_event(VK_MEDIA_NEXT_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                }
            }
            if (dataFromArduino.StartsWith("B2"))
            {
                if (radbutPrev2.Checked == true)
                {
                    // Jump to previous track
                    keybd_event(VK_MEDIA_PREV_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                }
                if (radbutPauPlay2.Checked == true)
                {
                    // Play or Pause music
                    keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                }
                if (radbutNext2.Checked == true)
                {
                    // Jump to next track
                    keybd_event(VK_MEDIA_NEXT_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                }
            }
            if (dataFromArduino.StartsWith("B3"))
            {
                if (radbutPrev3.Checked == true)
                {
                    // Jump to previous track
                    keybd_event(VK_MEDIA_PREV_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                }
                if (radbutPauPlay3.Checked == true)
                {
                    // Play or Pause music
                    keybd_event(VK_MEDIA_PLAY_PAUSE, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                }
                if (radbutNext3.Checked == true)
                {
                    // Jump to next track
                    keybd_event(VK_MEDIA_NEXT_TRACK, 0, KEYEVENTF_EXTENTEDKEY, IntPtr.Zero);
                }
            }

            //foreach (var process1 in Process.GetProcessesByName("cmd"))
            //{
            //    process1.Kill();
            //}
            else
            {

            }
            
        }

        private void showWindow()
        {
            this.Show();
        }

        private void cbSelectApp0_DropDown(object sender, EventArgs e)
        {
            cbSelectApp0.Items.Clear();
            foreach (AudioSession session in AudioUtilities.GetAllSessions())
            {
                if (session.Process != null) //&& session.Process.MainWindowTitle != "")
                {
                    if (session.Process.ProcessName.Contains("Discord"))
                    {
                        if (session.Process.MainWindowTitle == "")
                        {
                            cbSelectApp0.Items.Add($"{session.Process.ProcessName}.exe (Voice):{session.Process.Id}");
                        }
                        else if (!session.Process.MainWindowTitle.Contains(":"))
                        {
                            cbSelectApp0.Items.Add($"{session.Process.ProcessName}.exe (TTS):{session.Process.Id}");
                        }
                    }
                    else
                    {
                        cbSelectApp0.Items.Add($"{session.Process.ProcessName}.exe:{session.Process.Id}");
                    }
                }
            }
            cbNumItems = cbSelectApp0.Items.Count;

        }
        private void cbSelectApp1_DropDown(object sender, EventArgs e)
        {
            cbSelectApp1.Items.Clear();
            foreach (AudioSession session in AudioUtilities.GetAllSessions())
            {
                if (session.Process != null) //&& session.Process.MainWindowTitle != "")
                {
                    if (session.Process.ProcessName.Contains("Discord"))
                    {
                        if (session.Process.MainWindowTitle == "")
                        {
                            cbSelectApp1.Items.Add($"{session.Process.ProcessName}.exe (Voice):{session.Process.Id}");
                        }
                        else
                        {
                            cbSelectApp1.Items.Add($"{session.Process.ProcessName}.exe (TTS):{session.Process.Id}");
                        }
                    }
                    else
                    {
                        cbSelectApp1.Items.Add($"{session.Process.ProcessName}.exe:{session.Process.Id}");
                    }
                }
            }
            cbNumItems = cbSelectApp1.Items.Count;
        }
        private void cbSelectApp2_DropDown(object sender, EventArgs e)
        {
            cbSelectApp2.Items.Clear();
            foreach (AudioSession session in AudioUtilities.GetAllSessions())
            {
                if (session.Process != null) //&& session.Process.MainWindowTitle != "")
                {
                    if (session.Process.ProcessName.Contains("Discord"))
                    {
                        if (session.Process.MainWindowTitle == "")
                        {
                            cbSelectApp2.Items.Add($"{session.Process.ProcessName}.exe (Voice):{session.Process.Id}");
                        }
                        else if (!session.Process.MainWindowTitle.Contains(":"))
                        {
                            cbSelectApp2.Items.Add($"{session.Process.ProcessName}.exe (TTS):{session.Process.Id}");
                        }
                    }
                    else
                    {
                        cbSelectApp2.Items.Add($"{session.Process.ProcessName}.exe:{session.Process.Id}");
                    }
                }
            }
            cbNumItems = cbSelectApp2.Items.Count;

        }
        private void cbSelectApp3_DropDown(object sender, EventArgs e)
        {
            cbSelectApp3.Items.Clear();
            foreach (AudioSession session in AudioUtilities.GetAllSessions())
            {
                if (session.Process != null) //&& session.Process.MainWindowTitle != "")
                {
                    if (session.Process.ProcessName.Contains("Discord"))
                    {
                        if (session.Process.MainWindowTitle == "")
                        {
                            cbSelectApp3.Items.Add($"{session.Process.ProcessName}.exe (Voice):{session.Process.Id}");
                        }
                        else if (!session.Process.MainWindowTitle.Contains(":"))
                        {
                            cbSelectApp3.Items.Add($"{session.Process.ProcessName}.exe (TTS):{session.Process.Id}");
                        }
                    }
                    else
                    {
                        cbSelectApp3.Items.Add($"{session.Process.ProcessName}.exe:{session.Process.Id}");
                    }
                }
            }
            cbNumItems = cbSelectApp3.Items.Count;
        }

        private void cbSelectApp0_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbSelectApp0.Text.ToString() == "Main Volume")
            {
                Console.WriteLine(cbSelectApp0.Text.ToString());
            }
            else
            {
                strSelectedApp0 = cbSelectApp0.Text.ToString(); //.Substring(comboBox1.SelectedValue.ToString().IndexOf(":") + 1);
                strSelectedAppID0 = strSelectedApp0.Substring(strSelectedApp0.IndexOf(":") + 1);
                //Console.WriteLine(strSelectedAppID0);
                //intSelectedAppID0 = Convert.ToInt32(strSelectedAppID0);
                intSelectedAppID0 = Convert.ToInt32(strSelectedAppID0);
                //Console.WriteLine(intSelectedAppID0);
                

                //Console.WriteLine("P0:" + cbSelectApp0.SelectedItem.ToString() + '\n');
                //activeSerial.Write("P0:" + cbSelectApp0.SelectedItem.ToString() + '\n');
                activeSerial.WriteLine(cbSelectApp0.SelectedItem.ToString().Substring(0, cbSelectApp0.SelectedItem.ToString().LastIndexOf(".exe")) + '\n');
                Console.Write("Sending \"" + cbSelectApp0.SelectedItem.ToString().Substring(0, cbSelectApp0.SelectedItem.ToString().LastIndexOf(".exe")) + "\" to arduino." + '\n');
            }
        }
        private void cbSelectApp1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbSelectApp1.SelectedText == "Main Volume")
            {

            }
            else
            {
                strSelectedApp1 = cbSelectApp1.Text.ToString(); //.Substring(comboBox1.SelectedValue.ToString().IndexOf(":") + 1);
                strSelectedAppID1 = strSelectedApp1.Substring(strSelectedApp1.IndexOf(":") + 1);
                Console.WriteLine(strSelectedAppID1);

                intSelectedAppID1 = Convert.ToInt32(strSelectedAppID1);
                Console.WriteLine(intSelectedAppID1);
                Console.WriteLine(strSelectedApp1);
            }
        }
        private void cbSelectApp2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbSelectApp2.SelectedText == "Main Volume")
            {

            }
            else
            {
                strSelectedApp2 = cbSelectApp2.Text.ToString(); //.Substring(comboBox1.SelectedValue.ToString().IndexOf(":") + 1);
                strSelectedAppID2 = strSelectedApp2.Substring(strSelectedApp2.IndexOf(":") + 1);
                Console.WriteLine(strSelectedAppID2);

                intSelectedAppID2 = Convert.ToInt32(strSelectedAppID2);
            }
        }
        private void cbSelectApp3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbSelectApp3.SelectedText == "Main Volume")
            {

            }
            else
            {
            strSelectedApp3 = cbSelectApp3.Text.ToString(); //.Substring(comboBox1.SelectedValue.ToString().IndexOf(":") + 1);
            strSelectedAppID3 = strSelectedApp3.Substring(strSelectedApp3.IndexOf(":") + 1);
            Console.WriteLine(strSelectedAppID3);

            intSelectedAppID3 = Convert.ToInt32(strSelectedAppID3);
            }
        }
        

        private void btnApply_Click(object sender, EventArgs e)
        {
            

        }

        private void chkSysVol_CheckedChanged(object sender, EventArgs e)
        {
            if (chkSysVol.Checked == true)
            {
                cbSelectApp0.Enabled = false;
                //cbSelectApp0.Items.Clear();
            }
            else
            {
                cbSelectApp0.Enabled = true;
            }
        }


        private void btnCustButtons_Click(object sender, EventArgs e)
        {
            if (!customizingButtons)
            {
                this.Size = new Size(385, 493);
                btnCustButtons.Location = new Point(12, 410);
                //Location = new Point((Screen.PrimaryScreen.Bounds.Size.Width / 2) - (Size.Width / 2), (Screen.PrimaryScreen.Bounds.Size.Height / 2) - (Size.Height / 2));
                customizingButtons = true;
                btnCustButtons.Text = "Main Menu";
                gb0.Visible = false;
                gb1.Visible = false;
                gb2.Visible = false;
                gb3.Visible = false;
                gbButton1.Visible = true;
                gbButton2.Visible = true;
                gbButton3.Visible = true;
                gbButton4.Visible = true;
            }
            else
            {
                Size = new Size(269, 450);
                //Location = new Point((Screen.PrimaryScreen.Bounds.Size.Width / 2) - (Size.Width / 2), (Screen.PrimaryScreen.Bounds.Size.Height / 2) - (Size.Height / 2));
                btnCustButtons.Location = new Point(12, 374);
                btnCustButtons.Text = "Customize Buttons";
                customizingButtons = false;
                gb0.Visible = true;
                gb1.Visible = true;
                gb2.Visible = true;
                gb3.Visible = true;
                gbButton1.Visible = false;
                gbButton2.Visible = false;
                gbButton3.Visible = false;
                gbButton4.Visible = false;
            }
        }

        public void button1_Click(object sender, EventArgs e)
        {

            //foreach (AudioSession session in AudioUtilities.GetAllSessions())
            //{
            //    if (session.Process != null && session.Process.MainWindowTitle != "")
            //    {

            //    }
            //}
            //Process[] process = Process.GetProcesses();

            //intSelectedAppID0 = Convert.ToInt32(strSelectedAppID0);
            ////var curvol = Convert.ToInt32(VolumeMixer.GetApplicationVolume(intSelectedAppID0));
            //var curvol = Convert.ToInt32(GetApplicationVolume(14204));
            ////VolumeMixer.SetApplicationVolume(intSelectedAppID0, curvol + 10);
            //SetApplicationVolume(17376, curvol + 10);
            //Console.WriteLine(curvol);


            //CoreAudioDevice defaultPlaybackDevice = new CoreAudioController().DefaultPlaybackDevice;
            //Debug.WriteLine("Current Volume:" + defaultPlaybackDevice.Volume);
            //defaultPlaybackDevice.Volume = 100;
        }

        public void button2_Click(object sender, EventArgs e)
        {
            //foreach (AudioSession session in AudioUtilities.GetAllSessions())
            //{
            //    if (session.Process != null && session.Process.MainWindowTitle != "")
            //    {

            //    }
            //}
            //Process[] process = Process.GetProcesses();

            //intSelectedAppID0 = Convert.ToInt32(strSelectedAppID0);
            ////var curvol = Convert.ToInt32(VolumeMixer.GetApplicationVolume(intSelectedAppID0));
            //var curvol = Convert.ToInt32(GetApplicationVolume(14204));
            //// VolumeMixer.SetApplicationVolume(intSelectedAppID0, curvol - 10);
            //SetApplicationVolume(17376, curvol - 10);
            //Console.WriteLine(curvol);


            //CoreAudioDevice defaultPlaybackDevice = new CoreAudioController().DefaultPlaybackDevice;
            //Debug.WriteLine("Current Volume:" + defaultPlaybackDevice.Volume);
            //defaultPlaybackDevice.Volume = 0;
        }




        public static float? GetApplicationVolume(int pid)
        {
            ISimpleAudioVolume volume = GetVolumeObject(pid);
            if (volume == null)
                return null;

            float level;
            volume.GetMasterVolume(out level);
            Marshal.ReleaseComObject(volume);
            //Console.WriteLine("GetApplicationVolume()");

            return level * 100;
        }

        public static bool? GetApplicationMute(int pid)
        {
            ISimpleAudioVolume volume = GetVolumeObject(pid);
            if (volume == null)
                return null;

            bool mute;
            volume.GetMute(out mute);
            Marshal.ReleaseComObject(volume);

            //Console.WriteLine("GetApplicationMute()");
            return mute;
        }

        public static void SetApplicationVolume(int pid, float level)
        {
            ISimpleAudioVolume volume = GetVolumeObject(pid);
            if (volume == null)
                return;

            Guid guid = Guid.Empty;
            volume.SetMasterVolume(level / 100, ref guid);
            Marshal.ReleaseComObject(volume);


            //Console.WriteLine("SetApplicationVolume()");
        }

        public static void SetApplicationMute(int pid, bool mute)
        {
            ISimpleAudioVolume volume = GetVolumeObject(pid);
            if (volume == null)
                return;

            Guid guid = Guid.Empty;
            volume.SetMute(mute, ref guid);
            Marshal.ReleaseComObject(volume);

            //Console.WriteLine("SetApplicationMute()");
        }

        private static ISimpleAudioVolume GetVolumeObject(int pid)
        {
            // get the speakers (1st render + multimedia) device
            //IMMDeviceEnumerator deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
            IMMDeviceEnumerator deviceEnumerator = MMDeviceEnumeratorFactory.CreateInstance();
            //deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
            IMMDevice speakers;
            deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out speakers);

            // activate the session manager. we need the enumerator
            Guid IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
            object o;
            speakers.Activate(ref IID_IAudioSessionManager2, 0, IntPtr.Zero, out o);
            IAudioSessionManager2 mgr = (IAudioSessionManager2)o;

            // enumerate sessions for on this device
            IAudioSessionEnumerator sessionEnumerator;
            mgr.GetSessionEnumerator(out sessionEnumerator);
            int count;
            sessionEnumerator.GetCount(out count);

            // search for an audio session with the required name
            // NOTE: we could also use the process id instead of the app name (with IAudioSessionControl2)
            ISimpleAudioVolume volumeControl = null;
            for (int i = 0; i < count; i++)
            {
                IAudioSessionControl2 ctl;
                sessionEnumerator.GetSession(i, out ctl);
                int cpid;
                ctl.GetProcessId(out cpid);

                if (cpid == pid)
                {
                    volumeControl = ctl as ISimpleAudioVolume;
                    break;
                }
                Marshal.ReleaseComObject(ctl);
            }
            Marshal.ReleaseComObject(sessionEnumerator);
            Marshal.ReleaseComObject(mgr);
            Marshal.ReleaseComObject(speakers);
            Marshal.ReleaseComObject(deviceEnumerator);
            //Console.WriteLine("GetVolumeObject()");
            return volumeControl;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            notifyIcon1.Visible = false;
            if (FormWindowState.Minimized == this.WindowState)
            {
                notifyIcon1.Visible = true;
                //notifyIcon1.ShowBalloonTip(500);
                this.Hide();
            }

            else if (FormWindowState.Normal == this.WindowState)
            {
                notifyIcon1.Visible = false;
            }
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Show();
            notifyIcon1.Visible = false;
            this.WindowState = FormWindowState.Normal;
        }

        private void radbutPrev0_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radbutNext3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radbutPrev0_CheckedChanged_1(object sender, EventArgs e)
        {

        }
    }
    public static class Constants
    {
        //windows message id for hotkey
        public const int WM_HOTKEY_MSG_ID = 0x0312;
    }

    public class KeyHandler
    {
        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private int key;
        private IntPtr hWnd;
        private int id;

        public KeyHandler(Keys key, Form form)
        {
            this.key = (int)key;
            this.hWnd = form.Handle;
            id = this.GetHashCode();
        }

        public override int GetHashCode()
        {
            return key ^ hWnd.ToInt32();
        }

        public bool Register()
        {
            return RegisterHotKey(hWnd, id, 0, key);
        }

        public bool Unregiser()
        {
            return UnregisterHotKey(hWnd, id);
        }
    }


    //public class VolumeMixer
    //{
    //    public static float? GetApplicationVolume(int pid)
    //    {
    //        ISimpleAudioVolume volume = GetVolumeObject(pid);
    //        if (volume == null)
    //            return null;

    //        float level;
    //        volume.GetMasterVolume(out level);
    //        Marshal.ReleaseComObject(volume);
    //        return level * 100;
    //    }

    //    public static bool? GetApplicationMute(int pid)
    //    {
    //        ISimpleAudioVolume volume = GetVolumeObject(pid);
    //        if (volume == null)
    //            return null;

    //        bool mute;
    //        volume.GetMute(out mute);
    //        Marshal.ReleaseComObject(volume);
    //        return mute;
    //    }

    //    public static void SetApplicationVolume(int pid, float level)
    //    {
    //        ISimpleAudioVolume volume = GetVolumeObject(pid);
    //        if (volume == null)
    //            return;

    //        Guid guid = Guid.Empty;
    //        volume.SetMasterVolume(level / 100, ref guid);
    //        Marshal.ReleaseComObject(volume);
    //    }

    //    public static void SetApplicationMute(int pid, bool mute)
    //    {
    //        ISimpleAudioVolume volume = GetVolumeObject(pid);
    //        if (volume == null)
    //            return;

    //        Guid guid = Guid.Empty;
    //        volume.SetMute(mute, ref guid);
    //        Marshal.ReleaseComObject(volume);
    //    }

    //    private static ISimpleAudioVolume GetVolumeObject(int pid)
    //    {
    //        // get the speakers (1st render + multimedia) device
    //        //IMMDeviceEnumerator deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
    //        IMMDevice speakers;
    //        Form1.deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out speakers);

    //        // activate the session manager. we need the enumerator
    //        Guid IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
    //        object o;
    //        speakers.Activate(ref IID_IAudioSessionManager2, 0, IntPtr.Zero, out o);
    //        IAudioSessionManager2 mgr = (IAudioSessionManager2)o;

    //        // enumerate sessions for on this device
    //        IAudioSessionEnumerator sessionEnumerator;
    //        mgr.GetSessionEnumerator(out sessionEnumerator);
    //        int count;
    //        sessionEnumerator.GetCount(out count);

    //        // search for an audio session with the required name
    //        // NOTE: we could also use the process id instead of the app name (with IAudioSessionControl2)
    //        ISimpleAudioVolume volumeControl = null;
    //        for (int i = 0; i < count; i++)
    //        {
    //            IAudioSessionControl2 ctl;
    //            sessionEnumerator.GetSession(i, out ctl);
    //            int cpid;
    //            ctl.GetProcessId(out cpid);

    //            if (cpid == pid)
    //            {
    //                volumeControl = ctl as ISimpleAudioVolume;
    //                break;
    //            }
    //            Marshal.ReleaseComObject(ctl);
    //        }
    //        Marshal.ReleaseComObject(sessionEnumerator);
    //        Marshal.ReleaseComObject(mgr);
    //        Marshal.ReleaseComObject(speakers);
    //        Marshal.ReleaseComObject(Form1.deviceEnumerator);
    //        return volumeControl;
    //    }
    //}

    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    internal class MMDeviceEnumerator
    {
    }

    [ComImport]
    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceEnumerator
    {
        int NotImpl1();
        
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);

        // the rest is not implemented
    }


    public static class MMDeviceEnumeratorFactory
    {
        private static readonly Guid MMDeviceEnumerator = new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E");

        public static IMMDeviceEnumerator CreateInstance()
        {
            var type = Type.GetTypeFromCLSID(MMDeviceEnumerator);
            return (IMMDeviceEnumerator)Activator.CreateInstance(type);
        }
    }

    public enum EDataFlow
    {
        eRender,
        eCapture,
        eAll,
        EDataFlow_enum_count
    }

    public enum ERole
    {
        eConsole,
        eMultimedia,
        eCommunications,
        ERole_enum_count
    }

    //[Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    //internal interface IMMDeviceEnumerator
    //{
    //    int NotImpl1();

    //    [PreserveSig]
    //    int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);

    //    // the rest is not implemented
    //}

    [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDevice
    {
        [PreserveSig]
        int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);

        // the rest is not implemented
    }

    [Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionManager2
    {
        int NotImpl1();
        int NotImpl2();

        [PreserveSig]
        int GetSessionEnumerator(out IAudioSessionEnumerator SessionEnum);

        // the rest is not implemented
    }

    [Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionEnumerator
    {
        [PreserveSig]
        int GetCount(out int SessionCount);

        [PreserveSig]
        int GetSession(int SessionCount, out IAudioSessionControl2 Session);
    }

    [Guid("87CE5498-68D6-44E5-9215-6DA47EF883D8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ISimpleAudioVolume
    {
        [PreserveSig]
        int SetMasterVolume(float fLevel, ref Guid EventContext);

        [PreserveSig]
        int GetMasterVolume(out float pfLevel);

        [PreserveSig]
        int SetMute(bool bMute, ref Guid EventContext);

        [PreserveSig]
        int GetMute(out bool pbMute);
    }

    [Guid("bfb7ff88-7239-4fc9-8fa2-07c950be9c6d"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionControl2
    {
        // IAudioSessionControl
        [PreserveSig]
        int NotImpl0();

        [PreserveSig]
        int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

        [PreserveSig]
        int SetDisplayName([MarshalAs(UnmanagedType.LPWStr)]string Value, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

        [PreserveSig]
        int GetIconPath([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

        [PreserveSig]
        int SetIconPath([MarshalAs(UnmanagedType.LPWStr)] string Value, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

        [PreserveSig]
        int GetGroupingParam(out Guid pRetVal);

        [PreserveSig]
        int SetGroupingParam([MarshalAs(UnmanagedType.LPStruct)] Guid Override, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

        [PreserveSig]
        int NotImpl1();

        [PreserveSig]
        int NotImpl2();

        // IAudioSessionControl2
        [PreserveSig]
        int GetSessionIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

        [PreserveSig]
        int GetSessionInstanceIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

        [PreserveSig]
        int GetProcessId(out int pRetVal);

        [PreserveSig]
        int IsSystemSoundsSession();

        [PreserveSig]
        int SetDuckingPreference(bool optOut);
    }

    // audio utilities
    public static class AudioUtilities
    {
        private static IAudioSessionManager2 GetAudioSessionManager()
        {
            IMMDevice speakers = GetSpeakers();
            if (speakers == null)
                return null;

            // win7+ only
            object o;
            if (speakers.Activate(typeof(IAudioSessionManager2).GUID, CLSCTX.CLSCTX_ALL, IntPtr.Zero, out o) != 0 || o == null)
                return null;

            return o as IAudioSessionManager2;
        }

        public static AudioDevice GetSpeakersDevice()
        {
            return CreateDevice(GetSpeakers());
        }

        private static AudioDevice CreateDevice(IMMDevice dev)
        {
            if (dev == null)
                return null;

            string id;
            dev.GetId(out id);
            DEVICE_STATE state;
            dev.GetState(out state);
            Dictionary<string, object> properties = new Dictionary<string, object>();
            IPropertyStore store;
            dev.OpenPropertyStore(STGM.STGM_READ, out store);
            if (store != null)
            {
                int propCount;
                store.GetCount(out propCount);
                for (int j = 0; j < propCount; j++)
                {
                    PROPERTYKEY pk;
                    if (store.GetAt(j, out pk) == 0)
                    {
                        PROPVARIANT value = new PROPVARIANT();
                        int hr = store.GetValue(ref pk, ref value);
                        object v = value.GetValue();
                        try
                        {
                            if (value.vt != VARTYPE.VT_BLOB) // for some reason, this fails?
                            {
                                PropVariantClear(ref value);
                            }
                        }
                        catch
                        {
                        }
                        string name = pk.ToString();
                        properties[name] = v;
                    }
                }
            }
            return new AudioDevice(id, (AudioDeviceState)state, properties);
        }

        public static IList<AudioDevice> GetAllDevices()
        {
            List<AudioDevice> list = new List<AudioDevice>();
            IMMDeviceEnumerator deviceEnumerator = null;
            try
            {
                deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
            }
            catch
            {
            }
            if (deviceEnumerator == null)
                return list;

            IMMDeviceCollection collection;
            deviceEnumerator.EnumAudioEndpoints(EDataFlow.eAll, DEVICE_STATE.MASK_ALL, out collection);
            if (collection == null)
                return list;

            int count;
            collection.GetCount(out count);
            for (int i = 0; i < count; i++)
            {
                IMMDevice dev;
                collection.Item(i, out dev);
                if (dev != null)
                {
                    list.Add(CreateDevice(dev));
                }
            }
            return list;
        }

        private static IMMDevice GetSpeakers()
        {
            // get the speakers (1st render + multimedia) device
            try
            {
                IMMDeviceEnumerator deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
                IMMDevice speakers;
                deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out speakers);
                return speakers;
            }
            catch
            {
                return null;
            }
        }

        public static IList<AudioSession> GetAllSessions()
        {
            List<AudioSession> list = new List<AudioSession>();
            IAudioSessionManager2 mgr = GetAudioSessionManager();
            if (mgr == null)
                return list;

            IAudioSessionEnumerator sessionEnumerator;
            mgr.GetSessionEnumerator(out sessionEnumerator);
            int count;
            sessionEnumerator.GetCount(out count);

            for (int i = 0; i < count; i++)
            {
                IAudioSessionControl ctl;
                sessionEnumerator.GetSession(i, out ctl);
                if (ctl == null)
                    continue;

                IAudioSessionControl2 ctl2 = ctl as IAudioSessionControl2;
                if (ctl2 != null)
                {
                    list.Add(new AudioSession(ctl2));
                }
            }
            Marshal.ReleaseComObject(sessionEnumerator);
            Marshal.ReleaseComObject(mgr);
            return list;
        }

        public static AudioSession GetProcessSession()
        {
            int id = Process.GetCurrentProcess().Id;
            foreach (AudioSession session in GetAllSessions())
            {
                if (session.ProcessId == id)
                    return session;

                session.Dispose();
            }
            return null;
        }

        [DllImport("ole32.dll")]
        private static extern int PropVariantClear(ref PROPVARIANT pvar);

        [ComImport]
        [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        private class MMDeviceEnumerator
        {
        }

        [Flags]
        private enum CLSCTX
        {
            CLSCTX_INPROC_SERVER = 0x1,
            CLSCTX_INPROC_HANDLER = 0x2,
            CLSCTX_LOCAL_SERVER = 0x4,
            CLSCTX_REMOTE_SERVER = 0x10,
            CLSCTX_ALL = CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER | CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER
        }

        private enum STGM
        {
            STGM_READ = 0x00000000,
        }

        private enum EDataFlow
        {
            eRender,
            eCapture,
            eAll,
        }

        private enum ERole
        {
            eConsole,
            eMultimedia,
            eCommunications,
        }

        private enum DEVICE_STATE
        {
            ACTIVE = 0x00000001,
            DISABLED = 0x00000002,
            NOTPRESENT = 0x00000004,
            UNPLUGGED = 0x00000008,
            MASK_ALL = 0x0000000F
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct PROPERTYKEY
        {
            public Guid fmtid;
            public int pid;

            public override string ToString()
            {
                return fmtid.ToString("B") + " " + pid;
            }
        }

        // NOTE: we only define what we handle
        [Flags]
        private enum VARTYPE : short
        {
            VT_I4 = 3,
            VT_BOOL = 11,
            VT_UI4 = 19,
            VT_LPWSTR = 31,
            VT_BLOB = 65,
            VT_CLSID = 72,
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct PROPVARIANT
        {
            public VARTYPE vt;
            public ushort wReserved1;
            public ushort wReserved2;
            public ushort wReserved3;
            public PROPVARIANTunion union;

            public object GetValue()
            {
                switch (vt)
                {
                    case VARTYPE.VT_BOOL:
                        return union.boolVal != 0;

                    case VARTYPE.VT_LPWSTR:
                        return Marshal.PtrToStringUni(union.pwszVal);

                    case VARTYPE.VT_UI4:
                        return union.lVal;

                    case VARTYPE.VT_CLSID:
                        return (Guid)Marshal.PtrToStructure(union.puuid, typeof(Guid));

                    default:
                        return vt.ToString() + ":?";
                }
            }
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct PROPVARIANTunion
        {
            [FieldOffset(0)]
            public int lVal;
            [FieldOffset(0)]
            public ulong uhVal;
            [FieldOffset(0)]
            public short boolVal;
            [FieldOffset(0)]
            public IntPtr pwszVal;
            [FieldOffset(0)]
            public IntPtr puuid;
        }

        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDeviceEnumerator
        {
            [PreserveSig]
            int EnumAudioEndpoints(EDataFlow dataFlow, DEVICE_STATE dwStateMask, out IMMDeviceCollection ppDevices);

            [PreserveSig]
            int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);

            [PreserveSig]
            int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string pwstrId, out IMMDevice ppDevice);

            [PreserveSig]
            int RegisterEndpointNotificationCallback(IMMNotificationClient pClient);

            [PreserveSig]
            int UnregisterEndpointNotificationCallback(IMMNotificationClient pClient);
        }

        [Guid("7991EEC9-7E89-4D85-8390-6C703CEC60C0"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMNotificationClient
        {
            void OnDeviceStateChanged([MarshalAs(UnmanagedType.LPWStr)] string pwstrDeviceId, DEVICE_STATE dwNewState);
            void OnDeviceAdded([MarshalAs(UnmanagedType.LPWStr)] string pwstrDeviceId);
            void OnDeviceRemoved([MarshalAs(UnmanagedType.LPWStr)] string deviceId);
            void OnDefaultDeviceChanged(EDataFlow flow, ERole role, string pwstrDefaultDeviceId);
            void OnPropertyValueChanged([MarshalAs(UnmanagedType.LPWStr)] string pwstrDeviceId, PROPERTYKEY key);
        }

        [Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDeviceCollection
        {
            [PreserveSig]
            int GetCount(out int pcDevices);

            [PreserveSig]
            int Item(int nDevice, out IMMDevice ppDevice);
        }

        [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDevice
        {
            [PreserveSig]
            int Activate([MarshalAs(UnmanagedType.LPStruct)] Guid riid, CLSCTX dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);

            [PreserveSig]
            int OpenPropertyStore(STGM stgmAccess, out IPropertyStore ppProperties);

            [PreserveSig]
            int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);

            [PreserveSig]
            int GetState(out DEVICE_STATE pdwState);
        }

        [Guid("6f79d558-3e96-4549-a1d1-7d75d2288814"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPropertyDescription
        {
            [PreserveSig]
            int GetPropertyKey(out PROPERTYKEY pkey);

            [PreserveSig]
            int GetCanonicalName(out IntPtr ppszName);

            [PreserveSig]
            int GetPropertyType(out short pvartype);

            [PreserveSig]
            int GetDisplayName(out IntPtr ppszName);

            // WARNING: the rest is undefined. you *can't* implement it, only use it.
        }

        [Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPropertyStore
        {
            [PreserveSig]
            int GetCount(out int cProps);

            [PreserveSig]
            int GetAt(int iProp, out PROPERTYKEY pkey);

            [PreserveSig]
            int GetValue(ref PROPERTYKEY key, ref PROPVARIANT pv);

            [PreserveSig]
            int SetValue(ref PROPERTYKEY key, ref PROPVARIANT propvar);

            [PreserveSig]
            int Commit();
        }

        [Guid("BFA971F1-4D5E-40BB-935E-967039BFBEE4"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAudioSessionManager
        {
            [PreserveSig]
            int GetAudioSessionControl([MarshalAs(UnmanagedType.LPStruct)] Guid AudioSessionGuid, int StreamFlags, out IAudioSessionControl SessionControl);

            [PreserveSig]
            int GetSimpleAudioVolume([MarshalAs(UnmanagedType.LPStruct)] Guid AudioSessionGuid, int StreamFlags, ISimpleAudioVolume AudioVolume);
        }

        [Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAudioSessionManager2
        {
            [PreserveSig]
            int GetAudioSessionControl([MarshalAs(UnmanagedType.LPStruct)] Guid AudioSessionGuid, int StreamFlags, out IAudioSessionControl SessionControl);

            [PreserveSig]
            int GetSimpleAudioVolume([MarshalAs(UnmanagedType.LPStruct)] Guid AudioSessionGuid, int StreamFlags, ISimpleAudioVolume AudioVolume);

            [PreserveSig]
            int GetSessionEnumerator(out IAudioSessionEnumerator SessionEnum);

            [PreserveSig]
            int RegisterSessionNotification(IAudioSessionNotification SessionNotification);

            [PreserveSig]
            int UnregisterSessionNotification(IAudioSessionNotification SessionNotification);

            int RegisterDuckNotificationNotImpl();
            int UnregisterDuckNotificationNotImpl();
        }

        [Guid("641DD20B-4D41-49CC-ABA3-174B9477BB08"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAudioSessionNotification
        {
            void OnSessionCreated(IAudioSessionControl NewSession);
        }

        [Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAudioSessionEnumerator
        {
            [PreserveSig]
            int GetCount(out int SessionCount);

            [PreserveSig]
            int GetSession(int SessionCount, out IAudioSessionControl Session);
        }

        [Guid("bfb7ff88-7239-4fc9-8fa2-07c950be9c6d"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IAudioSessionControl2
        {
            // IAudioSessionControl
            [PreserveSig]
            int GetState(out AudioSessionState pRetVal);

            [PreserveSig]
            int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

            [PreserveSig]
            int SetDisplayName([MarshalAs(UnmanagedType.LPWStr)]string Value, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

            [PreserveSig]
            int GetIconPath([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

            [PreserveSig]
            int SetIconPath([MarshalAs(UnmanagedType.LPWStr)] string Value, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

            [PreserveSig]
            int GetGroupingParam(out Guid pRetVal);

            [PreserveSig]
            int SetGroupingParam([MarshalAs(UnmanagedType.LPStruct)] Guid Override, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

            [PreserveSig]
            int RegisterAudioSessionNotification(IAudioSessionEvents NewNotifications);

            [PreserveSig]
            int UnregisterAudioSessionNotification(IAudioSessionEvents NewNotifications);

            // IAudioSessionControl2
            [PreserveSig]
            int GetSessionIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

            [PreserveSig]
            int GetSessionInstanceIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

            [PreserveSig]
            int GetProcessId(out int pRetVal);

            [PreserveSig]
            int IsSystemSoundsSession();

            [PreserveSig]
            int SetDuckingPreference(bool optOut);
        }

        [Guid("F4B1A599-7266-4319-A8CA-E70ACB11E8CD"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IAudioSessionControl
        {
            [PreserveSig]
            int GetState(out AudioSessionState pRetVal);

            [PreserveSig]
            int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

            [PreserveSig]
            int SetDisplayName([MarshalAs(UnmanagedType.LPWStr)]string Value, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

            [PreserveSig]
            int GetIconPath([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

            [PreserveSig]
            int SetIconPath([MarshalAs(UnmanagedType.LPWStr)] string Value, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

            [PreserveSig]
            int GetGroupingParam(out Guid pRetVal);

            [PreserveSig]
            int SetGroupingParam([MarshalAs(UnmanagedType.LPStruct)] Guid Override, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

            [PreserveSig]
            int RegisterAudioSessionNotification(IAudioSessionEvents NewNotifications);

            [PreserveSig]
            int UnregisterAudioSessionNotification(IAudioSessionEvents NewNotifications);
        }

        [Guid("24918ACC-64B3-37C1-8CA9-74A66E9957A8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IAudioSessionEvents
        {
            void OnDisplayNameChanged([MarshalAs(UnmanagedType.LPWStr)] string NewDisplayName, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);
            void OnIconPathChanged([MarshalAs(UnmanagedType.LPWStr)] string NewIconPath, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);
            void OnSimpleVolumeChanged(float NewVolume, bool NewMute, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);
            void OnChannelVolumeChanged(int ChannelCount, IntPtr NewChannelVolumeArray, int ChangedChannel, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);
            void OnGroupingParamChanged([MarshalAs(UnmanagedType.LPStruct)] Guid NewGroupingParam, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);
            void OnStateChanged(AudioSessionState NewState);
            void OnSessionDisconnected(AudioSessionDisconnectReason DisconnectReason);
        }

        [Guid("87CE5498-68D6-44E5-9215-6DA47EF883D8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface ISimpleAudioVolume
        {
            [PreserveSig]
            int SetMasterVolume(float fLevel, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

            [PreserveSig]
            int GetMasterVolume(out float pfLevel);

            [PreserveSig]
            int SetMute(bool bMute, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

            [PreserveSig]
            int GetMute(out bool pbMute);
        }
    }

    public sealed class AudioSession : IDisposable
    {
        private AudioUtilities.IAudioSessionControl2 _ctl;
        private Process _process;

        internal AudioSession(AudioUtilities.IAudioSessionControl2 ctl)
        {
            _ctl = ctl;
        }

        public Process Process
        {
            get
            {
                if (_process == null && ProcessId != 0)
                {
                    try
                    {
                        _process = Process.GetProcessById(ProcessId);
                    }
                    catch
                    {
                        // do nothing
                    }
                }
                return _process;
            }
        }

        public int ProcessId
        {
            get
            {
                CheckDisposed();
                int i;
                _ctl.GetProcessId(out i);
                return i;
            }
        }

        public string Identifier
        {
            get
            {
                CheckDisposed();
                string s;
                _ctl.GetSessionIdentifier(out s);
                return s;
            }
        }

        public string InstanceIdentifier
        {
            get
            {
                CheckDisposed();
                string s;
                _ctl.GetSessionInstanceIdentifier(out s);
                return s;
            }
        }

        public AudioSessionState State
        {
            get
            {
                CheckDisposed();
                AudioSessionState s;
                _ctl.GetState(out s);
                return s;
            }
        }

        public Guid GroupingParam
        {
            get
            {
                CheckDisposed();
                Guid g;
                _ctl.GetGroupingParam(out g);
                return g;
            }
            set
            {
                CheckDisposed();
                _ctl.SetGroupingParam(value, Guid.Empty);
            }
        }

        public string DisplayName
        {
            get
            {
                CheckDisposed();
                string s;
                _ctl.GetDisplayName(out s);
                return s;
            }
            set
            {
                CheckDisposed();
                string s;
                _ctl.GetDisplayName(out s);
                if (s != value)
                {
                    _ctl.SetDisplayName(value, Guid.Empty);
                }
            }
        }

        public string IconPath
        {
            get
            {
                CheckDisposed();
                string s;
                _ctl.GetIconPath(out s);
                return s;
            }
            set
            {
                CheckDisposed();
                string s;
                _ctl.GetIconPath(out s);
                if (s != value)
                {
                    _ctl.SetIconPath(value, Guid.Empty);
                }
            }
        }

        private void CheckDisposed()
        {
            if (_ctl == null)
                throw new ObjectDisposedException("Control");
        }

        public override string ToString()
        {
            string s = DisplayName;
            if (!string.IsNullOrEmpty(s))
                return "DisplayName: " + s;

            if (Process != null)
                return "Process: " + Process.ProcessName;

            return "Pid: " + ProcessId;
        }

        public void Dispose()
        {
            if (_ctl != null)
            {
                Marshal.ReleaseComObject(_ctl);
                _ctl = null;
            }
        }
    }

    public sealed class AudioDevice
    {
        internal AudioDevice(string id, AudioDeviceState state, IDictionary<string, object> properties)
        {
            Id = id;
            State = state;
            Properties = properties;
        }

        public string Id { get; private set; }
        public AudioDeviceState State { get; private set; }
        public IDictionary<string, object> Properties { get; private set; }

        public string Description
        {
            get
            {
                const string PKEY_Device_DeviceDesc = "{a45c254e-df1c-4efd-8020-67d146a850e0} 2";
                object value;
                Properties.TryGetValue(PKEY_Device_DeviceDesc, out value);
                return string.Format("{0}", value);
            }
        }

        public string ContainerId
        {
            get
            {
                const string PKEY_Devices_ContainerId = "{8c7ed206-3f8a-4827-b3ab-ae9e1faefc6c} 2";
                object value;
                Properties.TryGetValue(PKEY_Devices_ContainerId, out value);
                return string.Format("{0}", value);
            }
        }

        public string EnumeratorName
        {
            get
            {
                const string PKEY_Device_EnumeratorName = "{a45c254e-df1c-4efd-8020-67d146a850e0} 24";
                object value;
                Properties.TryGetValue(PKEY_Device_EnumeratorName, out value);
                return string.Format("{0}", value);
            }
        }

        public string InterfaceFriendlyName
        {
            get
            {
                const string DEVPKEY_DeviceInterface_FriendlyName = "{026e516e-b814-414b-83cd-856d6fef4822} 2";
                object value;
                Properties.TryGetValue(DEVPKEY_DeviceInterface_FriendlyName, out value);
                return string.Format("{0}", value);
            }
        }

        public string FriendlyName
        {
            get
            {
                const string DEVPKEY_Device_FriendlyName = "{a45c254e-df1c-4efd-8020-67d146a850e0} 14";
                object value;
                Properties.TryGetValue(DEVPKEY_Device_FriendlyName, out value);
                return string.Format("{0}", value);
            }
        }

        public override string ToString()
        {
            return FriendlyName;
        }
    }

    public enum AudioSessionState
    {
        Inactive = 0,
        Active = 1,
        Expired = 2
    }

    public enum AudioDeviceState
    {
        Active = 0x1,
        Disabled = 0x2,
        NotPresent = 0x4,
        Unplugged = 0x8,
    }

    public enum AudioSessionDisconnectReason
    {
        DisconnectReasonDeviceRemoval = 0,
        DisconnectReasonServerShutdown = 1,
        DisconnectReasonFormatChanged = 2,
        DisconnectReasonSessionLogoff = 3,
        DisconnectReasonSessionDisconnected = 4,
        DisconnectReasonExclusiveModeOverride = 5
    }

}
