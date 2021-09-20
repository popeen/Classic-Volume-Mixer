using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace ClassicVolumeMixer
{
    public partial class Form1 : Form
    {
        private String mixerPath = "C:\\windows\\System32\\sndvol.exe";

        private NotifyIcon notifyIcon = new NotifyIcon(new System.ComponentModel.Container());
        private ContextMenu contextMenu = new System.Windows.Forms.ContextMenu();
        private MenuItem openClassic = new System.Windows.Forms.MenuItem();
        private MenuItem exit = new System.Windows.Forms.MenuItem();

        public Form1()
        {
            InitializeComponent();
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
                     exit
            });
            
            openClassic.Index = 0;
            openClassic.Text = "Open Classic Volume Mixer";
            openClassic.Click += new System.EventHandler(openClassic_Click);

            exit.Index = 1;
            exit.Text = "Exit";
            exit.Click += new System.EventHandler(exit_Click);

        }

        private void notifyIcon_Click(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                openClassicMixer();
            }
        }

        private void openClassic_Click(object sender, EventArgs e)
        {
            openClassicMixer();
        }
        private void exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void openClassicMixer()
        {
            ProcessStartInfo proc = new ProcessStartInfo();
            proc.UseShellExecute = true;
            proc.FileName = mixerPath;
            Process.Start(proc);
        }
    }
}
