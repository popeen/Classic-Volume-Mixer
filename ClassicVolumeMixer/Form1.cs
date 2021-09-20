using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace ClassicVolumeMixer
{
    public partial class Form1 : Form
    {
        String mixerPath = "C:\\windows\\System32\\sndvol.exe";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.ShowInTaskbar = false;
            this.Visible = false;

            notifyIcon1.Icon = Icon.ExtractAssociatedIcon(mixerPath);
            notifyIcon1.Text = "Classic Mixer";
            notifyIcon1.Visible = true;
        }

        private void notifyIcon1_Click(object sender, EventArgs e)
        {
            ProcessStartInfo proc = new ProcessStartInfo();
            proc.UseShellExecute = true;
            proc.FileName = mixerPath;
            Process.Start(proc);
        }
    }
}
