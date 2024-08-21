using System;
using System.Collections;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Forms;
using CoreAudio;
using Microsoft.Win32;

namespace ClassicVolumeMixer
{
    public class Options
    {
        public bool CloseClick { get; set; }
        public bool AdjustWidth { get; set; }
        public bool HideMixer { get; set; }
    }

    public partial class Form1 : Form
    {
        private static readonly string WinDir = Environment.GetEnvironmentVariable("SystemRoot");
        private readonly string mixerPath = Path.Combine(WinDir, "Sysnative", "sndvol.exe");
        private readonly string controlPanelPath = Path.Combine(WinDir, "Sysnative", "control.exe");
        private readonly string soundPanelArgument = "mmsys.cpl";
        private readonly string soundIconsPath = Path.Combine(WinDir, "Sysnative", "SndVolSSO.dll");
        private readonly string saveFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ClassicVolumeMixerSettings.json");

        private readonly NotifyIcon notifyIcon = new NotifyIcon(new System.ComponentModel.Container());
        private ContextMenuStrip contextMenu = new ContextMenuStrip();
        private readonly ToolStripMenuItem openClassic = new ToolStripMenuItem();
        private readonly ToolStripMenuItem sounds = new ToolStripMenuItem();
        private readonly ToolStripMenuItem closeClick = new ToolStripMenuItem();
        private readonly ToolStripMenuItem adjustWidth = new ToolStripMenuItem();
        private readonly ToolStripMenuItem hideMixer = new ToolStripMenuItem();
        private readonly ToolStripMenuItem exit = new ToolStripMenuItem();

        private Process process = new Process();
        private readonly Timer timer = new Timer();
        private readonly Timer volumeChangeTimer = new Timer();
        private readonly Stopwatch stopwatch = Stopwatch.StartNew();
        private IntPtr handle;

        private bool isVisible;
        private Options options = new Options { AdjustWidth = true, CloseClick = true, HideMixer = false };
        private readonly Icon[] icons = new Icon[6];
        private bool showNoAudioDeviceWarning = true;

        public Form1()
        {
            InitializeComponent();
            EnumerateAudioDevices();
        }

        private void EnumerateAudioDevices()
        {
            foreach (var item in new MMDeviceEnumerator(new Guid()).EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active))
            {
                Console.WriteLine(item.DeviceFriendlyName);
                Console.WriteLine(item.DeviceFriendlyName.Remove(item.DeviceFriendlyName.Length - item.DeviceInterfaceFriendlyName.Length - 3));
            }
        }

        [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

        private Icon ExtractIcon(string sFile, int iIndex, bool flipColors)
        {
            ExtractIconEx(sFile, iIndex, out IntPtr intPtr, out _, 1);
            Icon icon = Icon.FromHandle(intPtr);

            if (flipColors)
            {
                icon = FlipIconColors(icon);
            }
            return icon;
        }

        private Icon FlipIconColors(Icon icon)
        {
            Bitmap bitmap = icon.ToBitmap();
            for (int y = 0; y < bitmap.Height; y++)
            {
                for (int x = 0; x < bitmap.Width; x++)
                {
                    Color pixelColor = bitmap.GetPixel(x, y);
                    Color flippedColor = Color.FromArgb(pixelColor.A, 255 - pixelColor.R, 255 - pixelColor.G, 255 - pixelColor.B);
                    bitmap.SetPixel(x, y, flippedColor);
                }
            }
            return Icon.FromHandle(bitmap.GetHicon());
        }

        public static bool ShouldUseDarkIcon()
        {
            const string keyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath))
            {
                if (key != null)
                {
                    object value = key.GetValue("SystemUsesLightTheme");
                    if (value is int systemUsesLightTheme)
                    {
                        return systemUsesLightTheme == 1;
                    }
                }
            }
            return false;
        }

        private void SetIcons()
        {
            bool shouldUseDark = ShouldUseDarkIcon();
            icons[0] = ExtractIcon(soundIconsPath, 3, shouldUseDark); // one bar
            icons[1] = ExtractIcon(soundIconsPath, 4, shouldUseDark); // two bars
            icons[2] = icons[3] = ExtractIcon(soundIconsPath, 5, shouldUseDark); // three bars
            icons[4] = ExtractIcon(soundIconsPath, 1, shouldUseDark); // mute
            icons[5] = ExtractIcon(soundIconsPath, 2, shouldUseDark); // zero bars
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.ShowInTaskbar = false;
            this.Visible = false;

            SetIcons();

            timer.Interval = 100;  //if the Mixer takes too long to close after losing focus lower this value
            timer.Tick += Timer_Tick;

            volumeChangeTimer.Interval = 100;
            volumeChangeTimer.Tick += VolumeChangeTimer_Tick;
            volumeChangeTimer.Start();

            notifyIcon.Text = "Classic Mixer";
            notifyIcon.Visible = true;
            notifyIcon.MouseClick += NotifyIcon_Click;
            LoadContextMenu();

            if (File.Exists(saveFile))
            {
                ReadOptions();
            }
            else
            {
                WriteOptions();
            }
        }

        private void VolumeChangeTimer_Tick(object sender, EventArgs e)
        {
            ChangeTrayIconVolume();
        }

        private void ChangeTrayIconVolume()
        {
            try
            {
                MMDevice defaultAudioDevice = new MMDeviceEnumerator(new Guid()).GetDefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);
                int volume = (int)(defaultAudioDevice.AudioEndpointVolume.MasterVolumeLevelScalar * 100);
                if (defaultAudioDevice.AudioEndpointVolume.Mute)
                {
                    notifyIcon.Icon = icons[4];
                }
                else if (volume == 0)
                {
                    notifyIcon.Icon = icons[5];
                }
                else
                {
                    notifyIcon.Icon = icons[((volume - 1) / 33)];
                }
            }
            catch
            {
                if (showNoAudioDeviceWarning)
                {
                    showNoAudioDeviceWarning = false;
                    MessageBox.Show("There is no audio device on the system. Classic Volume Mixer will close.", "Classic Volume Mixer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Application.Exit();
                }
            }
        }

        private void LoadContextMenu()
        {
            contextMenu = new ContextMenuStrip();

            contextMenu.Opening += ContextMenu_Opening;
            contextMenu.Closing += ContextMenu_Closing;

            AddAudioDevicesToContextMenu(DataFlow.Capture);
            contextMenu.Items.Add(new ToolStripSeparator());
            AddAudioDevicesToContextMenu(DataFlow.Render);
            contextMenu.Items.Add(new ToolStripSeparator());

            contextMenu.Items.AddRange(new ToolStripMenuItem[] {
                openClassic,
                sounds,
                closeClick,
                adjustWidth,
                hideMixer,
                exit
            });

            openClassic.Text = "Open Classic Volume Mixer";
            openClassic.Click += OpenClassic_Click;

            sounds.Text = "Sound";
            sounds.Click += OpenSoundControl;

            closeClick.Text = "Close by clicking outside the window";
            closeClick.Checked = options.CloseClick;
            closeClick.Click += CloseClickToggle;

            adjustWidth.Text = "Dynamically adjust window width";
            adjustWidth.Checked = options.AdjustWidth;
            adjustWidth.Click += AdjustWidthToggle;

            hideMixer.Text = "Hide mixer instead of closing it";
            hideMixer.Checked = options.HideMixer;
            hideMixer.Click += HideMixerToggle;

            exit.Text = "Exit";
            exit.Click += Exit_Click;

            notifyIcon.ContextMenuStrip = contextMenu;
        }

        private void AddAudioDevicesToContextMenu(DataFlow dataFlow)
        {
            foreach (MMDevice device in new MMDeviceEnumerator(new Guid()).EnumerateAudioEndPoints(dataFlow, DeviceState.Active))
            {
                ToolStripMenuItem audioMenuItem = new ToolStripMenuItem(device.DeviceFriendlyName);
                contextMenu.Items.Add(audioMenuItem);
                if (device.Selected)
                {
                    audioMenuItem.Checked = true;
                }

                audioMenuItem.Click += (sender, e) => SetDefaultAudioDevice(device);
            }
        }

        private void SetDefaultAudioDevice(MMDevice device)
        {
            new CoreAudio.CPolicyConfigVistaClient().SetDefaultDevice(device.ID);
            ChangeTrayIconVolume();
            LoadContextMenu();
        }

        private void ReadOptions()
        {
            options = JsonSerializer.Deserialize<Options>(File.ReadAllText(saveFile));
            closeClick.Checked = options.CloseClick;
            adjustWidth.Checked = options.AdjustWidth;
            hideMixer.Checked = options.HideMixer;
        }

        private void WriteOptions()
        {
            File.WriteAllText(saveFile, JsonSerializer.Serialize(options));
        }

        private void HideMixerToggle(object sender, EventArgs e)
        {
            hideMixer.Checked = !hideMixer.Checked;
            options.HideMixer = !options.HideMixer;
            WriteOptions();
        }

        private void AdjustWidthToggle(object sender, EventArgs e)
        {
            adjustWidth.Checked = !adjustWidth.Checked;
            options.AdjustWidth = !options.AdjustWidth;
            WriteOptions();
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

        private void CloseClickToggle(object sender, EventArgs e)
        {
            closeClick.Checked = !closeClick.Checked;
            SetForegroundWindow(handle);
            options.CloseClick = !options.CloseClick;
            WriteOptions();
        }

        private void OpenSoundControl(object sender, EventArgs e)
        {
            Process soundProcess = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = controlPanelPath,
                    Arguments = soundPanelArgument,
                    UseShellExecute = true
                }
            };
            soundProcess.Start();
        }

        private void NotifyIcon_Click(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (process.HasExited && stopwatch.ElapsedMilliseconds > 100)
                {
                    OpenClassicMixer();
                    isVisible = true;
                    timer.Start();
                }
                else
                {
                    if (isVisible)
                    {
                        CloseMixer();
                        timer.Stop();
                    }
                    else if (stopwatch.ElapsedMilliseconds > 100)
                    {
                        ShowWindowAsync(handle, 1);
                        SetForegroundWindow(handle);
                        SetMixerPositionAndSize();
                        timer.Start();
                        isVisible = true;
                    }
                }
            }
        }

        private void OpenClassic_Click(object sender, EventArgs e)
        {
            if (process.HasExited)
            {
                OpenClassicMixer();
            }
            else
            {
                ShowWindowAsync(handle, 1);
                SetForegroundWindow(handle);
            }
            isVisible = true;
            timer.Start();
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            if (!process.HasExited)
            {
                process.Kill();
            }
            this.Close();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            IntPtr foregroundWindow = GetForegroundWindow();
            if ((foregroundWindow != handle) && closeClick.Checked)
            {
                CloseMixer();
                stopwatch.Restart();
                timer.Stop();
            }
        }

        private void CloseMixer()
        {
            if (hideMixer.Checked)
            {
                ShowWindowAsync(handle, 0);
                isVisible = false;
            }
            else
            {
                if (!process.HasExited)
                {
                    process.Kill();
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

        private void OpenClassicMixer()
        {
            process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            process.Start();
            process.WaitForInputIdle();

            Process[] processes = Process.GetProcessesByName("SndVol");
            foreach (Process process in processes)
            {
                while (process.MainWindowHandle == IntPtr.Zero) { } //busy waiting until the window is open
                handle = process.MainWindowHandle;
                SetMixerPositionAndSize();
            }
            SetForegroundWindow(handle);
        }

        private void SetMixerPositionAndSize()
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
            MoveWindow(handle, screenArea.Width - (160 + 110 * appCount), screenArea.Height - (corners.Bottom - corners.Top), 160 + 110 * appCount, 350, true);
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
                LoadContextMenu();
            }
            base.WndProc(ref m);
        }
    }
}
