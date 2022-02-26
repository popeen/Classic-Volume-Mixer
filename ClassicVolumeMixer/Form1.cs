﻿using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ClassicVolumeMixer
{
    public partial class Form1 : Form
    {

        // it's better to use the Windows Directory directly, because it can change and no be Windows.
        // private static String drive = System.Environment.GetEnvironmentVariable("SystemDrive");
        private static String WinDir = System.Environment.GetEnvironmentVariable("SystemRoot");  //location of windows installation
        private String mixerPath = WinDir + "\\System32\\sndvol.exe";
        private String soundControlPath = WinDir + "\\System32\\mmsys.cpl";
        private NotifyIcon notifyIcon = new NotifyIcon(new System.ComponentModel.Container());
        private ContextMenu contextMenu = new System.Windows.Forms.ContextMenu();
        private MenuItem openClassic = new System.Windows.Forms.MenuItem();
        private MenuItem sounds = new System.Windows.Forms.MenuItem();
        private MenuItem exit = new System.Windows.Forms.MenuItem();
        private Process process;
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
            notifyIcon.ContextMenu = contextMenu;

            contextMenu.MenuItems.AddRange(new
                System.Windows.Forms.MenuItem[] {
                     openClassic,
                     sounds,
                     exit
            });

            openClassic.Index = 0;
            openClassic.Text = "Open Classic Volume Mixer";
            openClassic.Click += new System.EventHandler(openClassic_Click);

            sounds.Index = 1;
            sounds.Text = "Sound";
            sounds.Click += new System.EventHandler(openSoundControl);

            exit.Index = 2;
            exit.Text = "Exit";
            exit.Click += new System.EventHandler(exit_Click);

        }

        private void openSoundControl(object sender, EventArgs e)
        {
            Process soundProcess = new Process();
            soundProcess.StartInfo.FileName = soundControlPath;
            soundProcess.StartInfo.UseShellExecute = true;
            soundProcess.Start();
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
                }
                else
                { 
                    if (isVisible)
                    {
                        ShowWindowAsync(handle, 0);
                    }
                    else {
                        ShowWindowAsync(handle, 1);
                        SetForegroundWindow(handle);
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
            else {
                ShowWindowAsync(handle, 1);
                SetForegroundWindow(handle);
            }
            isVisible = true;
        }
        private void exit_Click(object sender, EventArgs e)
        {
            if (!this.process.HasExited)
            {
                this.process.Kill();
            }
            this.Close();
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
