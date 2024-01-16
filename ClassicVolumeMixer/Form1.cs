using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections;
using System.IO;
using System.Text.Json;
using CoreAudio;
using System.Collections.Generic;

namespace ClassicVolumeMixer
{
    public class Options
    {
        public bool closeClick { get; set; }
        public bool adjustWidth { get; set; }
        public bool hideMixer { get; set; }
    }


    public partial class Form1 : Form
    {
        // it's better to use the Windows Directory directly, because it can change and no be "Windows".
        // private static String drive = System.Environment.GetEnvironmentVariable("SystemDrive");
        private static String WinDir = System.Environment.GetEnvironmentVariable("SystemRoot");  //location of windows installation
        private String mixerPath = WinDir + "\\Sysnative\\sndvol.exe";
        private String controlPanelPath = WinDir + "\\Sysnative\\control.exe";
        private String soundPanelArgument = "mmsys.cpl";
        private String saveFile = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\ClassicVolumeMixerSettings.json";
        private NotifyIcon notifyIcon = new NotifyIcon(new System.ComponentModel.Container());
        private ContextMenuStrip contextMenu = new System.Windows.Forms.ContextMenuStrip();
        private ToolStripMenuItem openClassic = new System.Windows.Forms.ToolStripMenuItem();
        private ToolStripMenuItem sounds = new System.Windows.Forms.ToolStripMenuItem();
        private ToolStripMenuItem closeClick = new System.Windows.Forms.ToolStripMenuItem();
        private ToolStripMenuItem adjustWidth = new System.Windows.Forms.ToolStripMenuItem();
        private ToolStripMenuItem hideMixer = new System.Windows.Forms.ToolStripMenuItem();
        private ToolStripMenuItem exit = new System.Windows.Forms.ToolStripMenuItem();
        private Process process;
        private Timer timer = new Timer();
        private IntPtr taskbar = IntPtr.Zero;
        Point rightClickPosition = new Point();
        Stopwatch stopwatch = Stopwatch.StartNew();
        IntPtr handle; // the handle of the mixer window
        bool isVisible;
        private Options options = new Options { adjustWidth = true, closeClick = true, hideMixer = false };
        String defaultAudioDevice;

        public Form1()
        {
            InitializeComponent();
            MMDeviceEnumerator test = new CoreAudio.MMDeviceEnumerator(new Guid());
            foreach (var item in new CoreAudio.MMDeviceEnumerator(new Guid()).EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
            {
                Console.WriteLine(item.DeviceFriendlyName);
                Console.WriteLine(item.DeviceFriendlyName.Remove(item.DeviceFriendlyName.Length - item.DeviceInterfaceFriendlyName.Length - 3));
            };


        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.ShowInTaskbar = false;
            this.Visible = false;

            notifyIcon.Icon = Icon.ExtractAssociatedIcon(mixerPath);
            notifyIcon.Text = "Classic Mixer";
            notifyIcon.Visible = true;
            notifyIcon.MouseClick += new MouseEventHandler(notifyIcon_Click);
            loadContextMenu();

            taskbar = FindWindow("Shell_TrayWnd", null);

            timer.Interval = 100;  //if the Mixer takes too long to close after losing focus lower this value
            timer.Tick += new EventHandler(timer_Tick);


            if (File.Exists(saveFile))
            {
                readOptions();
            }
            else
            {
                writeOptions();
            }
        }

        private void loadContextMenu()
        {
            contextMenu = new System.Windows.Forms.ContextMenuStrip();

            contextMenu.Opening += ContextMenu_Opening;
            contextMenu.Closing += ContextMenu_Closing;

            foreach (MMDevice device in new CoreAudio.MMDeviceEnumerator(new Guid()).EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active))
            {
                ToolStripMenuItem audioMenuItem = new ToolStripMenuItem(device.DeviceFriendlyName);
                contextMenu.Items.Add(audioMenuItem);
                if (device.Selected)
                {
                    audioMenuItem.Checked = true;
                }

                audioMenuItem.Click += (sender2, e2) => setDefaultAudioDevice(device);
            }

            contextMenu.Items.Add(new ToolStripSeparator());

            foreach (MMDevice device in new CoreAudio.MMDeviceEnumerator(new Guid()).EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
            {
                ToolStripMenuItem audioMenuItem = new ToolStripMenuItem(device.DeviceFriendlyName);
                contextMenu.Items.Add(audioMenuItem);
                if (device.Selected)
                {
                    audioMenuItem.Checked = true;
                }

                audioMenuItem.Click += (sender2, e2) => setDefaultAudioDevice(device);
            }

            contextMenu.Items.Add(new ToolStripSeparator());

            contextMenu.Items.AddRange(new
                System.Windows.Forms.ToolStripMenuItem[] {
                     openClassic,
                     sounds,
                     closeClick,
                     adjustWidth,
                     hideMixer,
                     exit
        });

            openClassic.Text = "Open Classic Volume Mixer";
            openClassic.Click += new System.EventHandler(openClassic_Click);

            sounds.Text = "Sound";
            sounds.Click += new System.EventHandler(openSoundControl);

            closeClick.Text = "Close by clicking outside the window";
            closeClick.Checked = true;
            closeClick.Click += new System.EventHandler(closeClickToggle);

            adjustWidth.Text = "Dynamically adjust window width";
            adjustWidth.Checked = true;
            adjustWidth.Click += new System.EventHandler(adjustWidthToggle);

            hideMixer.Text = "Hide mixer instead of closing it";
            hideMixer.Checked = false;
            hideMixer.Click += new System.EventHandler(hideMixerToggle);

            exit.Text = "Exit";
            exit.Click += new System.EventHandler(exit_Click);

            notifyIcon.ContextMenuStrip = contextMenu;
        }

        private void setDefaultAudioDevice(MMDevice device)
        {
            new CoreAudio.CPolicyConfigVistaClient().SetDefaultDevice(device.ID);
            loadContextMenu();
        }

        /**
        * reads the options from a json file and adjusts the checkboxes acordingly
        */
        private void readOptions()
        {
            options = JsonSerializer.Deserialize<Options>(File.ReadAllText(saveFile));
            closeClick.Checked = options.closeClick;
            adjustWidth.Checked = options.adjustWidth;
            hideMixer.Checked = options.hideMixer;
        }

        /**
         * writes the options to a json file
         */
        private void writeOptions()
        {
            File.WriteAllText(saveFile, JsonSerializer.Serialize(options));
        }

        private void hideMixerToggle(object sender, EventArgs e)
        {
            hideMixer.Checked = !hideMixer.Checked;
            options.hideMixer = !options.hideMixer;
            writeOptions();
        }

        private void adjustWidthToggle(object sender, EventArgs e)
        {
            adjustWidth.Checked = !adjustWidth.Checked;
            options.adjustWidth = !options.adjustWidth;
            writeOptions();
        }

        private void ContextMenu_Closing(object sender, ToolStripDropDownClosingEventArgs e)
        {
            if (isVisible)
            {
                timer.Start();
            }
        }

        private void ContextMenu_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            timer.Stop();
        }

        private void closeClickToggle(object sender, EventArgs e)
        {
            closeClick.Checked = !closeClick.Checked;
            SetForegroundWindow(handle);
            options.closeClick = !options.closeClick;
            writeOptions();
        }

        private void openSoundControl(object sender, EventArgs e)
        {
            Process soundProcess = new Process();
            soundProcess.StartInfo.FileName = controlPanelPath;
            soundProcess.StartInfo.Arguments = soundPanelArgument;
            soundProcess.StartInfo.UseShellExecute = true;
            soundProcess.Start();
        }

        private void notifyIcon_Click(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                //check if the mixer is currently open. 
                if (this.process.HasExited && stopwatch.ElapsedMilliseconds > 100)
                {
                    openClassicMixer();
                    isVisible = true;
                    timer.Start();
                }
                else
                {
                    if (isVisible)
                    {
                        closeMixer();
                        timer.Stop();
                    }
                    else if (stopwatch.ElapsedMilliseconds > 100)
                    {
                        ShowWindowAsync(handle, 1);
                        SetForegroundWindow(handle);
                        setMixerPositionAndSize();
                        timer.Start();
                        isVisible = true;
                    }
                }
            }
        }

        private void openClassic_Click(object sender, EventArgs e)
        {
            if (this.process.HasExited)
            {
                openClassicMixer();
            }
            else
            {
                ShowWindowAsync(handle, 1);
                SetForegroundWindow(handle);
            }
            isVisible = true;
            timer.Start();
        }
        private void exit_Click(object sender, EventArgs e)
        {
            if (!this.process.HasExited)
            {
                this.process.Kill();
            }
            this.Close();
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            IntPtr foregroundWindow = GetForegroundWindow();
            if ((foregroundWindow != handle) && closeClick.Checked)
            {
                closeMixer();
                stopwatch.Restart();
                timer.Stop();
            }
        }

        private void closeMixer()
        {
            if (hideMixer.Checked)
            {
                ShowWindowAsync(handle, 0);
                isVisible = false;
            }
            else
            {
                if (!this.process.HasExited)
                {
                    this.process.Kill();
                }
            }
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hwnd, ref Rect rectangle);

        [DllImport("user32.dll")]
        static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr window, EnumedWindow callback, ArrayList lParam);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);



        [StructLayout(LayoutKind.Sequential)]
        public struct Rect
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }

        private delegate bool EnumedWindow(IntPtr handleWindow, ArrayList handles);

        private void openClassicMixer()
        {
            this.process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            this.process.Start();
            this.process.WaitForInputIdle();

            Process[] processes = Process.GetProcessesByName("SndVol");
            foreach (Process process in processes)
            {
                while (process.MainWindowHandle == IntPtr.Zero) { } //busy waiting until the window is open
                this.handle = process.MainWindowHandle;
                setMixerPositionAndSize();
            }
            SetForegroundWindow(this.handle);
        }

        //sets the mixers position to bottom right of the PrimaryScreen and adjusts the window width depending on the number of active sound application
        private void setMixerPositionAndSize()
        {
            Rectangle screenArea = Screen.PrimaryScreen.WorkingArea;
            Rect corners = new Rect();
            GetWindowRect(handle, ref corners);

            ArrayList windowHandles = new ArrayList();
            EnumedWindow callBackPtr = GetWindowHandle;
            EnumChildWindows(handle, callBackPtr, windowHandles);
            int appCount = 3;
            if (adjustWidth.Checked)
            {
                appCount = (windowHandles.Count - 12) / 7;
            }
            GetWindowRect(handle, ref corners);
            MoveWindow(this.handle, screenArea.Width - (160 + 110 * appCount), screenArea.Height - (corners.Bottom - corners.Top), 160 + 110 * appCount, 350, true);
        }

        private static bool GetWindowHandle(IntPtr windowHandle, ArrayList windowHandles)
        {
            windowHandles.Add(windowHandle);
            return true;
        }

        protected override void WndProc(ref Message m)
        {
            //WM_DEVICECHANGE = 0x0219;
            if (m.Msg == 0x0219)
            {
                loadContextMenu();
            }
            base.WndProc(ref m);
        }
    }
}
