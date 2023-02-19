using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static ClassicVolumeMixer.Form1;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using System.Windows.Input;

namespace ClassicVolumeMixer
{
    public partial class Form1 : Form
    {
        // it's better to use the Windows Directory directly, because it can change and no be "Windows".
        // private static String drive = System.Environment.GetEnvironmentVariable("SystemDrive");
        private static String WinDir = System.Environment.GetEnvironmentVariable("SystemRoot");  //location of windows installation
        private String mixerPath = WinDir + "\\System32\\sndvol.exe";
        private String soundControlPath = WinDir + "\\System32\\mmsys.cpl";
        private NotifyIcon notifyIcon = new NotifyIcon(new System.ComponentModel.Container());
        private ContextMenuStrip contextMenu = new System.Windows.Forms.ContextMenuStrip();
        private ToolStripMenuItem openClassic = new System.Windows.Forms.ToolStripMenuItem();
        private ToolStripMenuItem sounds = new System.Windows.Forms.ToolStripMenuItem();
        private ToolStripMenuItem closeClick = new System.Windows.Forms.ToolStripMenuItem();
        private ToolStripMenuItem exit = new System.Windows.Forms.ToolStripMenuItem();
        private Process process;
        private Timer timer = new Timer();
        Stopwatch stopwatch = Stopwatch.StartNew();
        IntPtr handle; // the handle of the mixer window
        bool isVisible;

        public Form1()
        {
            InitializeComponent();
            Process[] processlist = Process.GetProcesses();

            foreach (Process process in processlist)
            {
                if (!String.IsNullOrEmpty(process.MainWindowTitle))
                {
                    Console.WriteLine("Process: {0} ID: {1} Window title: {2}", process.ProcessName, process.Id, process.MainWindowTitle);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.ShowInTaskbar = false;
            this.Visible = false;

            notifyIcon.Icon = Icon.ExtractAssociatedIcon(mixerPath);
            notifyIcon.Text = "Classic Mixer";
            notifyIcon.Visible = true;
            notifyIcon.MouseClick += new MouseEventHandler(notifyIcon_Click);
            notifyIcon.MouseMove += new MouseEventHandler(notifyIcon_MouseMove);
            notifyIcon.ContextMenuStrip = contextMenu;

            contextMenu.Opening += ContextMenu_Opening;
            contextMenu.Closing += ContextMenu_Closing;

            contextMenu.Items.AddRange(new
                System.Windows.Forms.ToolStripMenuItem[] {
                     openClassic,
                     sounds,
                     closeClick,
                     exit
            });

            openClassic.Text = "Open Classic Volume Mixer";
            openClassic.Click += new System.EventHandler(openClassic_Click);

            sounds.Text = "Sound";
            sounds.Click += new System.EventHandler(openSoundControl);

            closeClick.Text = "Close by clicking outside the window";
            closeClick.Checked = true;
            closeClick.Click += new System.EventHandler(closeClickToggle);

            exit.Text = "Exit";
            exit.Click += new System.EventHandler(exit_Click);

            timer.Interval = 100;  //if the Mixer takes too long to close after losing focus lower this value
            timer.Tick += new EventHandler(timer_Tick);

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
        }

        private void openSoundControl(object sender, EventArgs e)
        {
            Process soundProcess = new Process();
            soundProcess.StartInfo.FileName = soundControlPath;
            soundProcess.StartInfo.UseShellExecute = true;
            soundProcess.Start();
        }

        private void notifyIcon_MouseMove(object sender, MouseEventArgs e)
        {
            stopwatch.Restart();
        }

        private void notifyIcon_Click(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                //check if the mixer is currently open. 
                if (this.process.HasExited)
                {
                    openClassicMixer();
                    isVisible = true;
                    timer.Start();
                }
                else
                {
                    Console.WriteLine(stopwatch.ElapsedMilliseconds > 1000);
                    if (isVisible)
                    {
                        ShowWindowAsync(handle, 0);
                        timer.Stop();
                    }
                    else
                    {
                        ShowWindowAsync(handle, 1);
                        SetForegroundWindow(handle);
                        timer.Start();
                    }
                    isVisible = !isVisible;
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
            if ((GetForegroundWindow() != handle) && (stopwatch.ElapsedMilliseconds > 1000) && closeClick.Checked)
            {
                ShowWindowAsync(handle, 0);
                isVisible = false;
                timer.Stop();
            }
        }

        [DllImport("user32.dll")]
        public static extern void SetWindowText(int hWnd, String text);

        [DllImport("user32.dll", EntryPoint = "SetWindowPos")]
        public static extern IntPtr SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int Y, int cx, int cy, int wFlags);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hwnd, ref Rect rectangle);

        [DllImport("user32.dll")]
        static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [StructLayout(LayoutKind.Sequential)]
        public struct Rect
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }

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
                Rect corners = new Rect();
                if (GetWindowRect(handle, ref corners)) //get window dimensions
                {
                    Rectangle screenArea = Screen.PrimaryScreen.WorkingArea;
                    //set window position to bottom right of the PrimaryScreen
                    SetWindowPos(handle, 0, screenArea.Width - (corners.Right - corners.Left), screenArea.Height - (corners.Bottom - corners.Top), 0, 0, 0x0041);
                }
            }
        }
    }
}
