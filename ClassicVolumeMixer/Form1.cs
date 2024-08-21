using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows.Forms;
using ClassicVolumeMixer.Helpers;
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
        private readonly ToolStripMenuItem openClassicVolumeMixer = new ToolStripMenuItem();
        private readonly ToolStripMenuItem openClassicSounds = new ToolStripMenuItem();
        private readonly ToolStripMenuItem openModernSounds = new ToolStripMenuItem();
        private readonly ToolStripMenuItem openModernVolumeMixer = new ToolStripMenuItem();
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
        }

        private void SetIcons()
        {
            bool shouldUseDarkIcon = ThemeHelper.SystemUsesLightTheme();
            icons[0] = IconHelper.ExtractIcon(soundIconsPath, 3, shouldUseDarkIcon); // one bar
            icons[1] = IconHelper.ExtractIcon(soundIconsPath, 4, shouldUseDarkIcon); // two bars
            icons[2] = icons[3] = IconHelper.ExtractIcon(soundIconsPath, 5, shouldUseDarkIcon); // three bars
            icons[4] = IconHelper.ExtractIcon(soundIconsPath, 1, shouldUseDarkIcon); // mute
            icons[5] = IconHelper.ExtractIcon(soundIconsPath, 2, shouldUseDarkIcon); // zero bars
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
            if (!AudioDeviceHelper.IsAudioDeviceAvailable())
            {
                if (showNoAudioDeviceWarning)
                {
                    showNoAudioDeviceWarning = false;
                    MessageBox.Show("There is no audio device on the system. Classic Volume Mixer will close.", "Classic Volume Mixer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Application.Exit();
                }
                return;
            }
                
            int volume = AudioDeviceHelper.GetVolumeLevel();
            if (AudioDeviceHelper.IsMuted())
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

        private void LoadContextMenu()
        {
            contextMenu = new ContextMenuStrip();

            contextMenu.Opening += ContextMenu_Opening;
            contextMenu.Closing += ContextMenu_Closing;

            AddAudioDevicesToContextMenu(DataFlow.Capture);
            contextMenu.Items.Add(new ToolStripSeparator());
            AddAudioDevicesToContextMenu(DataFlow.Render);
            contextMenu.Items.Add(new ToolStripSeparator());

            AddMenuItems();
            AddOptionsMenuItems();
            AddExitMenuItem();

            notifyIcon.ContextMenuStrip = contextMenu;
        }

        private void AddMenuItems()
        {
            contextMenu.Items.AddRange(new ToolStripMenuItem[] {
                openClassicVolumeMixer,
                openClassicSounds,
                openModernVolumeMixer,
                openModernSounds
            });
            contextMenu.Items.Add(new ToolStripSeparator());

            openClassicVolumeMixer.Text = "Open Classic Volume Mixer";
            openClassicVolumeMixer.Click += OpenClassicMixer_Click;

            openClassicSounds.Text = "Open Classic Sound Settings";
            openClassicSounds.Click += (sender, e) => StartProcess(controlPanelPath, soundPanelArgument);

            openModernVolumeMixer.Text = "Open Modern Volume Mixer";
            openModernVolumeMixer.Click += (sender, e) => OpenSettingsPage("apps-volume");

            openModernSounds.Text = "Open Modern Sound Settings";
            openModernSounds.Click += (sender, e) => OpenSettingsPage("sound");
        }

        private void AddOptionsMenuItems()
        {
            contextMenu.Items.AddRange(new ToolStripMenuItem[] {
                closeClick,
                adjustWidth,
                hideMixer,
            });
            contextMenu.Items.Add(new ToolStripSeparator());

            closeClick.Text = "Close by clicking outside the window";
            closeClick.Checked = options.CloseClick;
            closeClick.Click += CloseClickToggle;

            adjustWidth.Text = "Dynamically adjust window width";
            adjustWidth.Checked = options.AdjustWidth;
            adjustWidth.Click += AdjustWidthToggle;

            hideMixer.Text = "Hide mixer instead of closing it";
            hideMixer.Checked = options.HideMixer;
            hideMixer.Click += HideMixerToggle;
        }

        private void AddExitMenuItem()
        {
            exit.Text = "Exit";
            exit.Click += Exit_Click;
            contextMenu.Items.Add(exit);
        }

        private void AddAudioDevicesToContextMenu(DataFlow dataFlow)
        {
            foreach (MMDevice device in AudioDeviceHelper.GetAudioDevices(dataFlow))
            {
                ToolStripMenuItem audioMenuItem = new ToolStripMenuItem(device.DeviceFriendlyName);
                contextMenu.Items.Add(audioMenuItem);
                if (device.Selected)
                {
                    audioMenuItem.Checked = true;
                }

                audioMenuItem.Click += (sender, e) =>
                {
                    AudioDeviceHelper.SetDefaultAudioDevice(device);
                    ChangeTrayIconVolume();
                    LoadContextMenu();
                };
            }
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
            WindowHelper.SetForegroundWindow(handle);
            options.CloseClick = !options.CloseClick;
            WriteOptions();
        }

        private void StartProcess(String filename, String arguments = "")
        {
            Process process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = filename,
                    Arguments = arguments,
                    UseShellExecute = true
                }
            };
            process.Start();
        }
        private void OpenSettingsPage(string page)
        {
            StartProcess($"ms-settings:{page}");
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
                        WindowHelper.ShowWindowAsync(handle, 1);
                        WindowHelper.SetForegroundWindow(handle);
                        SetMixerPositionAndSize();
                        timer.Start();
                        isVisible = true;
                    }
                }
            }
        }

        private void OpenClassicMixer_Click(object sender, EventArgs e)
        {
            if (process.HasExited)
            {
                OpenClassicMixer();
            }
            else
            {
                WindowHelper.ShowWindowAsync(handle, 1);
                WindowHelper.SetForegroundWindow(handle);
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
            IntPtr foregroundWindow = WindowHelper.GetForegroundWindow();
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
                WindowHelper.ShowWindowAsync(handle, 0);
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
            WindowHelper.SetForegroundWindow(handle);
        }

        private void SetMixerPositionAndSize()
        {
            Rectangle screenArea = Screen.PrimaryScreen.WorkingArea;
            WindowHelper.Rect corners = new WindowHelper.Rect();
            WindowHelper.GetWindowRect(handle, ref corners);

            ArrayList windowHandles = new ArrayList();
            WindowHelper.EnumedWindow callBackPtr = WindowHelper.GetWindowHandle;
            WindowHelper.EnumChildWindows(handle, callBackPtr, windowHandles);
            int appCount = 3;
            if (adjustWidth.Checked)
            {
                appCount = (windowHandles.Count - 12) / 7;
            }
            WindowHelper.GetWindowRect(handle, ref corners);
            WindowHelper.MoveWindow(handle, screenArea.Width - (160 + 110 * appCount), screenArea.Height - (corners.Bottom - corners.Top), 160 + 110 * appCount, 350, true);
        }

    }
}
