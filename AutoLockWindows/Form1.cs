using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using Timer = System.Windows.Forms.Timer;

namespace AutoLockWindows
{
    public partial class Form1 : Form
    {
        private static Timer _timer;
        private static Timer _secondTimer;
        private DateTime _nextLockTime;
        private const string RegKey = @"SOFTWARE\WOW6432Node\Compute\AutoLockService";
        private readonly string _fullRegKey = @$"HKEY_LOCAL_MACHINE\{RegKey}";
        private const int DefaultInterval = 900000; //15 MINUTES
        private readonly string _logPath;

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool LockWorkStation();

        [DllImport("advapi32.dll")]
        static extern bool LogonUser(string userName, string domainName, string password, int logonType, int logonProvider, ref IntPtr phToken);


        public Form1()
        {
            _logPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty, "log.txt");
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Log("Starting up...", Color.GreenYellow);
            InitRegistry();
            InitFromRegistry();
            CreateTimer();
            ShowLoginPanel();
        }

        private void ShowLoginPanel()
        {
            this.pnlLogin.Width = this.ClientRectangle.Width;
            this.pnlLogin.Height = this.ClientRectangle.Height - this.lblTimeLeft.Height;
            this.pnlLogin.Location = new Point(0, this.lblTimeLeft.Height);
            this.pnlLogin.Visible = true;
            this.txtPassword.Text = "";
            this.Focus();
            this.txtPassword.Focus();
            this.txtPassword.Select();
        }

        private void InitRegistry()
        {
            using RegistryKey sk = Registry.LocalMachine.OpenSubKey(RegKey);
            if (sk == null)
            {
                Log("Registry key not found.");
                using RegistryKey skn = Registry.LocalMachine.CreateSubKey(RegKey);
                if (skn == null)
                {
                    Error("Could not create registry key");
                    return;
                }
                Log("Registry key created.");
                skn.SetValue("PreventAppNames", "", RegistryValueKind.String);
                Log("Registry value PreventAppNames set.");
                skn.SetValue("TimerInterval", DefaultInterval, RegistryValueKind.DWord);
                Log("Registry value TimerInterval set.");
            }
            else
            {
                Log("Registry key found.");
            }
        }

        private void InitFromRegistry()
        {
            string preventApps = (string)Registry.GetValue(_fullRegKey, "PreventAppNames", null);
            if (preventApps == null)
            {
                Log("Registry value PreventAppNames not found.");
                throw new Exception($"{_fullRegKey}\\PreventAppNames not found");
            }
            else
            {
                if (preventApps.Length > 0)
                {
                    Log("Registry value PreventAppNames found. Value: " + preventApps);
                    foreach (string app in preventApps.Split(','))
                    {
                        switch (app.ToLower())
                        {
                            case "winword.exe":
                                this.cbWord.Checked = true;
                                break;
                            case "excel.exe":
                                this.cbExcel.Checked = true;
                                break;
                        }
                    }
                }
            }

            int interval = (int)Registry.GetValue(_fullRegKey, "TimerInterval", 0);
            if (interval == 0)
            {
                Log("Registry value TimerInterval not found.");
                throw new Exception($"{_fullRegKey}\\TimerInterval not found");
            }

            //The whole unnecessary conversion thing is just to keep ReSharper quiet....
            this.numericUpDown1.Value = Math.Floor((decimal)(interval / 60000.0)); 
            Log($"Registry value TimerInterval found. Value: {interval}");
        }

        private void btRun_Click(object sender, EventArgs e)
        {
            if (_timer != null)
            {
                Warn("Timer service is already running.");
            }
            else
            {
                this.CreateTimer();
                Log("Timer Service was successfully started.", Color.Green);
            }
        }

        private void btStop_Click(object sender, EventArgs e)
        {
            if (_timer == null)
            {
                Warn("Timer service is already stopped.");
            }
            else
            {
                this.DestroyTimer();
                Log("Timer Service was successfully stopped.", Color.OrangeRed);
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveToRegistry();
        }

        private void SaveToRegistry()
        {
            string preventApps = "";
            if (this.cbWord.Checked)
            {
                preventApps += (preventApps.Length > 0 ? "," : "") + "winword.exe";
            }

            if (cbExcel.Checked)
            {
                preventApps += (preventApps.Length > 0 ? "," : "") + "excel.exe";
            }

            Registry.SetValue(_fullRegKey, "PreventAppNames", preventApps, RegistryValueKind.String);
            Log("Saved PreventAppNames registry key. Value: " + preventApps, Color.MediumPurple);
            Registry.SetValue(_fullRegKey, "TimerInterval", this.numericUpDown1.Value * 60000, RegistryValueKind.DWord);
            Log("Saved TimerInterval registry key. Value: " + (this.numericUpDown1.Value * 6000), Color.MediumPurple);

            if (_timer != null)
            {
                CreateTimer();
                Log("Successfully restarted AutoLockService service.", Color.Green);
            }
        }

        private void Log(string txt, Color color)
        {
            var lvi = new ListViewItem { Text = txt, ForeColor = color };
            this.listView1.Items.Add(lvi);
            this.listView1.Refresh();
            this.listView1.EnsureVisible(this.listView1.Items.IndexOf(lvi));
            File.AppendAllText(_logPath, $@"{DateTime.Now:G} [App] {txt}{Environment.NewLine}");
        }

        private void Log(string txt)
        {
            this.Log(txt, this.listView1.ForeColor);
        }
        private void Warn(string txt)
        {
            this.Log(txt, Color.Orange);
        }
        private void Error(string txt)
        {
            this.Log(txt, Color.Red);
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            switch (this.WindowState)
            {
                case FormWindowState.Minimized:
                    this.Hide();
                    break;
                case FormWindowState.Normal:
                    break;
            }
        }

        private void notifyIcon1_Click(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.ShowLoginPanel();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LockWorkStation();
        }

        private void OnTimedEvent(object sender, EventArgs e)
        {
            Log($"Starting OnTimedEvent");
            this.ShowLoginPanel();
            string preventApps = (string)Registry.GetValue(_fullRegKey, "PreventAppNames", null);
            if (preventApps == null)
            {
                Log($"In OnTimedEvent: Registry value {_fullRegKey}\\PreventAppNames is null");
                return;
            }
            else
            {
                if (preventApps.Length > 0)
                {
                    foreach (string app in preventApps.Split(','))
                    {
                        Process[] processes = Process.GetProcessesByName(app);
                        if (processes.Length > 0)
                        {
                            Log($"In OnTimedEvent: {app} is running. Exiting run.");
                            return;
                        }
                    }
                }
            }

            if (!LockWorkStation())
            {
                try
                {
                    int lastError = Marshal.GetLastWin32Error();
                    throw new Win32Exception(lastError); // or any other thing
                }
                catch (Exception exception)
                {
                    Log($"In OnTimedEvent: LockWorkStation call failed. Exception: {exception.Message}");
                    throw;
                }
            }
            else
            {
                Log($"In OnTimedEvent: LockWorkStation called successfully.");
                this.Hide();
            }
        }
        private void CreateTimer()
        {
            DestroyTimer();

            int interval = (int)Registry.GetValue(_fullRegKey, "TimerInterval", 0);
            if (interval == 0)
            {
                Log($"In CreateTimer: Registry key value {_fullRegKey}\\TimerInterval, was not found. Timer can not be set.");
                return;
            }
            _timer = new Timer() { Interval = interval };
            _timer.Tick += OnTimedEvent;
            _timer.Start();
            _secondTimer = new Timer() {Interval = 1000};
            _secondTimer.Tick += SecondTimerTick;
            _secondTimer.Start();
            _nextLockTime = DateTime.Now.AddMilliseconds(interval);
            this.lblTimeLeft.Visible = true;
            Log($"In CreateTimer: Screen will be locked at {_nextLockTime:T}.");
        }

        private void SecondTimerTick(object sender, EventArgs e)
        {
            TimeSpan ts = (_nextLockTime - DateTime.Now);
            this.lblTimeLeft.Text = $@"הנעילה הבאה בעוד {ts.Hours} שעות {ts.Minutes} דקות {ts.Seconds} שניות";
            this.notifyIcon1.Text = this.lblTimeLeft.Text;
            
            if ((int)ts.TotalSeconds == 30)
            {
                this.notifyIcon1.BalloonTipText = "נעילה אוטומטית בעוד 30 שניות";
                this.notifyIcon1.ShowBalloonTip(30000);
            }
        }

        private void DestroyTimer()
        {
            if (_timer == null)
                return;
            _timer.Stop();
            _timer.Dispose();
            _timer = null;
            _secondTimer.Stop();
            _secondTimer.Dispose();
            _secondTimer = null;
            this.lblTimeLeft.Visible = false;
            Log($"In CreateTimer: Timer has been destroyed.");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DestroyTimer();
            notifyIcon1.Visible = false;
            notifyIcon1.Dispose();
        }

        private bool LogInUser(string password)
        {
            IntPtr tokenHandler = IntPtr.Zero;
            System.Security.Principal.WindowsIdentity currentUser = System.Security.Principal.WindowsIdentity.GetCurrent();
            string userName = currentUser.Name.Split("\\")[1];
            return LogonUser(userName, Environment.UserDomainName, password, 2, 0, ref tokenHandler);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (LogInUser(this.txtPassword.Text.Trim()))
            {
                this.pnlLogin.Visible = false;
            }
        }

        private void txtPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.button3.PerformClick();
            }
        }
    }
}
