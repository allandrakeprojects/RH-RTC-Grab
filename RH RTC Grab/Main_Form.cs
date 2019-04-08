using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace RH_RTC_Grab
{
    public partial class Main_Form : Form
    {
        private JObject __jo;
        private bool m_aeroEnabled;
        private bool __isClose;
        private int __secho;
        private int __display_length = 5000;
        private int __result_count_json;
        private int __total_page;
        private int __i = 0;
        private bool __isLogin = false;
        private bool __isStart = false;
        private bool __isBreak = false;
        private bool __is_send = false;
        private string __playerlist_cn;
        private string __playerlist_ea;
        private string __playerlist_qq;
        private string __player_id;
        private string __start_time;
        private string __end_time;
        private string __brand_code = "RH";
        private string __brand_color = "#9B8435";
        private string __app = "RTC Grab";
        private string __app_type = "0";
        private int __send = 0;
        Form __mainFormHandler;

        // Deposit
        private JObject __jo_deposit;
        private bool __isBreak_deposit = false;
        private bool __isInsert_deposit = false;
        private int __secho_deposit;
        private int __total_page_deposit;
        private int __result_count_json_deposit;
        private bool __detectInsert_deposit = false;

        // Drag Header to Move
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        // ----- Drag Header to Move

        // Form Shadow
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,
            int nTopRect,
            int nRightRect,
            int nBottomRect,
            int nWidthEllipse,
            int nHeightEllipse
        );
        [DllImport("dwmapi.dll")]
        public static extern int DwmExtendFrameIntoClientArea(IntPtr hWnd, ref MARGINS pMarInset);
        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
        [DllImport("dwmapi.dll")]
        public static extern int DwmIsCompositionEnabled(ref int pfEnabled);
        private const int CS_DROPSHADOW = 0x00020000;
        private const int WM_NCPAINT = 0x0085;
        private const int WM_ACTIVATEAPP = 0x001C;
        private const int WM_NCHITTEST = 0x84;
        private const int HTCLIENT = 0x1;
        private const int HTCAPTION = 0x2;
        private const int WS_MINIMIZEBOX = 0x20000;
        private const int CS_DBLCLKS = 0x8;
        public struct MARGINS
        {
            public int leftWidth;
            public int rightWidth;
            public int topHeight;
            public int bottomHeight;
        }
        protected override CreateParams CreateParams
        {
            get
            {
                m_aeroEnabled = CheckAeroEnabled();

                CreateParams cp = base.CreateParams;
                if (!m_aeroEnabled)
                    cp.ClassStyle |= CS_DROPSHADOW;

                cp.Style |= WS_MINIMIZEBOX;
                cp.ClassStyle |= CS_DBLCLKS;
                return cp;
            }
        }
        private bool CheckAeroEnabled()
        {
            if (Environment.OSVersion.Version.Major >= 6)
            {
                int enabled = 0;
                DwmIsCompositionEnabled(ref enabled);
                return (enabled == 1) ? true : false;
            }
            return false;
        }
        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_NCPAINT:
                    if (m_aeroEnabled)
                    {
                        var v = 2;
                        DwmSetWindowAttribute(Handle, 2, ref v, 4);
                        MARGINS margins = new MARGINS()
                        {
                            bottomHeight = 1,
                            leftWidth = 0,
                            rightWidth = 0,
                            topHeight = 0
                        };
                        DwmExtendFrameIntoClientArea(Handle, ref margins);

                    }
                    break;
                default:
                    break;
            }
            base.WndProc(ref m);

            if (m.Msg == WM_NCHITTEST && (int)m.Result == HTCLIENT)
                m.Result = (IntPtr)HTCAPTION;
        }
        // ----- Form Shadow

        public Main_Form()
        {
            InitializeComponent();

            timer_landing.Start();
        }

        // Drag to Move
        private void panel_header_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void label_title_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void panel2_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void pictureBox_loader_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void label_brand_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void label_player_last_registered_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void panel_landing_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void pictureBox_landing_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void pictureBox_header_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        // ----- Drag to Move

        // Click Close
        private void pictureBox_close_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Exit the program?", "RH RTC Grab", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                __isClose = true;
                Environment.Exit(0);
            }
        }

        // Click Minimize
        private void pictureBox_minimize_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        // Form Closing
        private void Main_Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!__isClose)
            {
                DialogResult dr = MessageBox.Show("Exit the program?", "RH RTC Grab", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dr == DialogResult.No)
                {
                    e.Cancel = true;
                }
                else
                {
                    Environment.Exit(0);
                }
            }

            Environment.Exit(0);
        }

        [DllImport("winmm.dll")]
        public static extern int waveOutGetVolume(IntPtr h, out uint dwVolume);

        // Mute Sounds
        [DllImport("winmm.dll")]
        public static extern int waveOutSetVolume(IntPtr h, uint dwVolume);

        // Form Load
        private void Main_Form_Load(object sender, EventArgs e)
        {
            int NewVolume = ((ushort.MaxValue / 10) * 100);
            uint NewVolumeAllChannels = (((uint)NewVolume & 0x0000ffff) | ((uint)NewVolume << 16));
            waveOutSetVolume(IntPtr.Zero, NewVolumeAllChannels);

            webBrowser.Navigate("http://rh893sh3d.7799779.com/_xLhPnNlm9H0ifrRhPnNlm9H0ifrrZZ/default.aspx");
        }

        static int LineNumber([System.Runtime.CompilerServices.CallerLineNumber] int lineNumber = 0)
        {
            return lineNumber;
        }

        // WebBrowser
        private async void webBrowser_DocumentCompletedAsync(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (webBrowser.ReadyState == WebBrowserReadyState.Complete)
            {
                if (e.Url == webBrowser.Url)
                {
                    try
                    {
                        if (webBrowser.Url.ToString().Equals("http://rh893sh3d.7799779.com/_xLhPnNlm9H0ifrRhPnNlm9H0ifrrZZ/default.aspx"))
                        {
                            if (__isStart)
                            {
                                label_brand.Visible = false;
                                pictureBox_loader.Visible = false;
                                label_player_last_registered.Visible = false;
                                label_page_count.Visible = false;
                                label_currentrecord.Visible = false;
                                __mainFormHandler = Application.OpenForms[0];
                                __mainFormHandler.Size = new Size(466, 468);

                                SendITSupport("The application have been logout, please re-login again.");
                                SendMyBot("The application have been logout, please re-login again.");
                                __send = 0;

                                if (__is_send)
                                {
                                    __isClose = false;
                                    Environment.Exit(0);
                                }
                            }

                            __isLogin = false;
                            __isStart = false;
                            timer.Stop();
                            label_player_last_registered.Text = "-";
                            HtmlElementCollection htmlcol = webBrowser.Document.GetElementsByTagName("input");
                            for (int i = 0; i < htmlcol.Count; i++)
                            {
                                if (htmlcol[i].OuterHtml.Contains("username"))
                                {
                                    htmlcol[i].SetAttribute("value", "devteam_@dm1n");
                                }

                                if (htmlcol[i].OuterHtml.Contains("password"))
                                {
                                    htmlcol[i].SetAttribute("value", "$$rhr0ngh0@$$");
                                }
                            }
                            webBrowser.Document.Window.ScrollTo(185, webBrowser.Document.Body.ScrollRectangle.Height);
                            webBrowser.Document.Body.Style = "zoom:.99";
                            webBrowser.Document.GetElementsByTagName("input").GetElementsByName("acode")[0].Focus();
                            webBrowser.Visible = true;
                            label_brand.Visible = false;
                            pictureBox_loader.Visible = false;
                            label_player_last_registered.Visible = false;
                            webBrowser.WebBrowserShortcutsEnabled = true;
                        }
                        
                        if (webBrowser.Url.ToString().Contains("adm/default.aspx"))
                        {
                            label_brand.Visible = true;
                            pictureBox_loader.Visible = true;
                            label_player_last_registered.Visible = true;
                            label_page_count.Visible = true;
                            label_currentrecord.Visible = true;
                            __mainFormHandler = Application.OpenForms[0];
                            __mainFormHandler.Size = new Size(466, 168);

                            __isLogin = true;

                            if (!__isStart)
                            {
                                __isStart = true;
                                webBrowser.Visible = false;
                                label_brand.Visible = true;
                                pictureBox_loader.Visible = true;
                                label_player_last_registered.Visible = true;
                                webBrowser.WebBrowserShortcutsEnabled = false;
                                await ___PlayerLastRegisteredAsync();
                                await ___GetPlayerListsRequest();
                                ___GetPlayerListsRequest_Deposit();
                            }
                        }
                    }
                    catch (Exception err)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());
                        __send = 0;

                        __isClose = false;
                        Environment.Exit(0);
                    }
                }
            }
        }

        // asdasdasd
        // ------ Functions
        private async Task ___GetPlayerListsRequest()
        {
            __isBreak = false;
            List<string> player_info = new List<string>();
            string _player_last_username = "";

            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                byte[] result = await wc.DownloadDataTaskAsync("http://rh893sh3d.7799779.com/_xLhPnNlm9H0ifrRhPnNlm9H0ifrRZZ/adm/player/xml_player_search.aspx?str=&by=email&orderby=joindate&sortorder=desc&fromrow=0&torow=5000 ");
                string responsebody = Encoding.UTF8.GetString(result);
                var xDoc = XDocument.Parse(responsebody);

                var emptyElements = xDoc.Descendants("search");
                int i = 0;

                foreach (var xe in emptyElements)
                {
                    string player_id = xe.Attribute("id").Value.Trim();
                    string name = xe.Attribute("fullname").Value.Trim();
                    string username = xe.Attribute("username").Value.Trim();
                    string datetime_register = xe.Attribute("joined").Value.Trim();
                    string player_ldd = await ___PlayerListLastDepositAsync(player_id);
                    string playerlist_cn = xe.Attribute("contact").Value.Replace("+", "").Trim();
                    if (!String.IsNullOrEmpty(playerlist_cn.ToString()))
                    {
                        if (playerlist_cn.Substring(0, 2) == "84")
                        {
                            playerlist_cn = playerlist_cn.Substring(2);
                        }
                    }
                    string playerlist_ea = xe.Attribute("email").Value.Trim();
                    string agent = "";
                    string playerlist_qq = "";

                    if (username.ToLower() != Properties.Settings.Default.______last_registered_player.ToLower())
                    {
                        if (i == 0)
                        {
                            _player_last_username = username;
                        }

                        player_info.Add(username + "*|*" + name + "*|*" + datetime_register + "*|*" + player_ldd + "*|*" + playerlist_cn + "*|*" + playerlist_ea + "*|*" + agent + "*|*" + playerlist_qq);
                    }
                    else
                    {
                        // send to api
                        if (player_info.Count != 0)
                        {
                            player_info.Reverse();
                            string player_info_get = String.Join(",", player_info);
                            string[] values = player_info_get.Split(',');
                            foreach (string value in values)
                            {
                                Application.DoEvents();
                                string[] values_inner = value.Split(new string[] { "*|*" }, StringSplitOptions.None);
                                int count = 0;
                                string _username = "";
                                string _name = "";
                                string _date_register = "";
                                string _date_deposit = "";
                                string _cn = "";
                                string _email = "";
                                string _agent = "";
                                string _qq = "";

                                foreach (string value_inner in values_inner)
                                {
                                    count++;

                                    // Username
                                    if (count == 1)
                                    {
                                        _username = value_inner;
                                    }
                                    // Name
                                    else if (count == 2)
                                    {
                                        _name = value_inner;
                                    }
                                    // Register Date
                                    else if (count == 3)
                                    {
                                        _date_register = value_inner;
                                    }
                                    // Last Deposit Date
                                    else if (count == 4)
                                    {
                                        if (!String.IsNullOrEmpty(value_inner))
                                        {
                                            _date_deposit = value_inner;
                                        }
                                        else
                                        {
                                            _date_deposit = "";
                                        }
                                    }
                                    // Contact Number
                                    else if (count == 5)
                                    {
                                        _cn = value_inner;
                                    }
                                    // Email
                                    else if (count == 6)
                                    {
                                        _email = value_inner;
                                    }
                                    // Agent
                                    else if (count == 7)
                                    {
                                        _agent = value_inner;
                                    }
                                    // QQ
                                    else if (count == 8)
                                    {
                                        _qq = value_inner;
                                    }
                                }

                                // ----- Insert Data
                                //using (StreamWriter file = new StreamWriter(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\rtcgrab_rh.txt", true, Encoding.UTF8))
                                //{
                                //    file.WriteLine(_username + "*|*" + _name + "*|*" + _date_register + "*|*" + _date_deposit + "*|*" + _cn + "*|*" + _email + "*|*" + _agent + "*|*" + _qq + "*|*" + __brand_code);
                                //}
                                // update 01/11/2018
                                try
                                {
                                    if (!__isSending)
                                    {
                                        if (File.Exists(Path.GetTempPath() + @"\rtcgrab_temp_sending_rh.txt"))
                                        {
                                            File.Delete(Path.GetTempPath() + @"\rtcgrab_temp_sending_rh.txt");
                                        }
                                        Properties.Settings.Default.______count_player = Properties.Settings.Default.______count_player++;
                                        Properties.Settings.Default.Save();
                                        using (StreamWriter file = new StreamWriter(Path.GetTempPath() + @"\rtcgrab_sending_rh.txt", true, Encoding.UTF8))
                                        {
                                            file.WriteLine(_username + "*|*" + _name + "*|*" + _date_register + "*|*" + _date_deposit + "*|*" + _cn + "*|*" + _email + "*|*" + _agent + "*|*" + _qq + "*|*" + __brand_code);
                                        }
                                        ___Send();
                                    }
                                    else
                                    {
                                        Properties.Settings.Default.______count_player = Properties.Settings.Default.______count_player++;
                                        Properties.Settings.Default.Save();
                                        using (StreamWriter file = new StreamWriter(Path.GetTempPath() + @"\rtcgrab_temp_sending_rh.txt", true, Encoding.UTF8))
                                        {
                                            file.WriteLine(_username + "*|*" + _name + "*|*" + _date_register + "*|*" + _date_deposit + "*|*" + _cn + "*|*" + _email + "*|*" + _agent + "*|*" + _qq + "*|*" + __brand_code);
                                        }
                                    }
                                }
                                catch (Exception err)
                                {
                                    SendMyBot(err.ToString());
                                }

                                ___InsertData(_username, _name, _date_register, _date_deposit, _cn, _email, _agent, _qq, __brand_code);

                                __send = 0;

                                __playerlist_cn = "";
                                __playerlist_ea = "";
                                __playerlist_qq = "";
                            }
                        }

                        if (!String.IsNullOrEmpty(_player_last_username.Trim()))
                        {
                            ___SavePlayerLastRegistered(_player_last_username);

                            Invoke(new Action(() =>
                            {
                                label_player_last_registered.Text = "Last Registered: " + Properties.Settings.Default.______last_registered_player;
                            }));
                        }

                        player_info.Clear();
                        timer.Start();
                        __isBreak = true;
                        break;
                    }

                    i++;
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    await ___GetPlayerListsRequest();
                }
            }
        }

        private async void timer_TickAsync(object sender, EventArgs e)
        {
            timer.Stop();
            await ___GetPlayerListsRequest();

            if (__isInsert_deposit)
            {
                __isInsert_deposit = false;
                ___GetPlayerListsRequest_Deposit();
            }
        }

        private void ___InsertData(string username, string name, string date_register, string date_deposit, string contact, string email, string agent, string qq, string brand_code)
        {
            try
            {
                string password = username.ToLower() + date_register + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["username"] = username,
                        ["name"] = name,
                        ["date_register"] = date_register,
                        ["date_deposit"] = date_deposit,
                        ["contact"] = contact,
                        ["email"] = email,
                        ["agent"] = agent,
                        ["qq"] = qq,
                        ["wc"] = "",
                        ["brand_code"] = brand_code,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://192.168.10.252:8080/API/sendRTC", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);

                    using (StreamWriter file = new StreamWriter(Path.GetTempPath() + @"\rtcgrab_rh.txt", true, Encoding.UTF8))
                    {
                        file.WriteLine(username + "*|*" + name + "*|*" + date_register + "*|*" + date_deposit + "*|*" + contact + "*|*" + email + "*|*" + agent + "*|*" + qq + "*|*" + __brand_code + " ----- " + responseInString);
                    }
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        ____InsertData2(username, name, date_register, date_deposit, contact, email, agent, qq, brand_code);
                    }
                }
            }
        }

        private void ____InsertData2(string username, string name, string date_register, string date_deposit, string contact, string email, string agent, string qq, string brand_code)
        {
            try
            {
                string password = username.ToLower() + date_register + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["username"] = username,
                        ["name"] = name,
                        ["date_register"] = date_register,
                        ["date_deposit"] = date_deposit,
                        ["contact"] = contact,
                        ["email"] = email,
                        ["agent"] = agent,
                        ["qq"] = qq,
                        ["wc"] = "",
                        ["brand_code"] = brand_code,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus.ssitex.com:8080/API/sendRTC", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);

                    using (StreamWriter file = new StreamWriter(Path.GetTempPath() + @"\rtcgrab_rh.txt", true, Encoding.UTF8))
                    {
                        file.WriteLine(username + "*|*" + name + "*|*" + date_register + "*|*" + date_deposit + "*|*" + contact + "*|*" + email + "*|*" + agent + "*|*" + qq + "*|*" + __brand_code + " ----- " + responseInString);
                    }
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        ___InsertData(username, name, date_register, date_deposit, contact, email, agent, qq, brand_code);
                    }
                }
            }
        }

        private async Task<string> ___PlayerListLastDepositAsync(string id)
        {
            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();
                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                string date_to = DateTime.Now.ToString("yyyy-MM-dd");
                string date_from = DateTime.Now.AddMonths(-1).ToString("yyyy-MM-dd");
                string responsebody = await wc.DownloadStringTaskAsync("http://rh893sh3d.7799779.com/_xLhPnNlm9H0ifrRhPnNlm9H0ifrRZZ/adm/player/inc_player_deposit_withdrawal.aspx?type=deposit&id=" + id + "&dateFrom=" + date_from + "&dateTo=" + date_to);
                if (responsebody.ToLower().Contains("no record"))
                {
                    return null;
                }
                else
                {
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(@responsebody);
                    var rows = doc.DocumentNode.SelectNodes("//*[@id='tblTrans']/tr");
                    int count = 0;
                    foreach (var row in rows)
                    {
                        string td = row.InnerHtml;
                        if (count != 0)
                        {
                            if (!td.ToLower().Contains("total"))
                            {
                                string date = row.SelectSingleNode("td[4]").InnerText;
                                string status = row.SelectSingleNode("td[5]").InnerText;
                                if (status.ToLower() == "processed")
                                {
                                    return date.Trim();
                                }
                                else
                                {
                                    return null;
                                }
                            }
                        }
                        count++;
                    }
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    await ___PlayerListLastDepositAsync(id);
                }
            }

            return null;
        }

        private async Task ___PlayerLastRegisteredAsync()
        {
            Properties.Settings.Default.______last_registered_player = "";
            Properties.Settings.Default.______last_registered_player_deposit = "";

            try
            {
                if (Properties.Settings.Default.______last_registered_player == "" && Properties.Settings.Default.______last_registered_player_deposit == "")
                {
                    await ___GetLastRegisteredPlayerAsync();
                }

                label_player_last_registered.Text = "Last Registered: " + Properties.Settings.Default.______last_registered_player;
            }
            catch (Exception err)
            {
                __send++;
                if (__send == 5)
                {
                    SendITSupport("There's a problem to the server, please re-open the application.");
                    SendMyBot(err.ToString());

                    __isClose = false;
                    Environment.Exit(0);
                }
                else
                {
                    ___WaitNSeconds(10);
                    await ___PlayerLastRegisteredAsync();
                }
            }
        }

        private void ___SavePlayerLastRegistered(string username)
        {
            Properties.Settings.Default.______last_registered_player = username.ToLower();
            Properties.Settings.Default.Save();
        }

        private void timer_landing_Tick(object sender, EventArgs e)
        {
            panel_landing.Visible = false;
            timer_landing.Stop();
        }

        // Deposit
        private async void ___GetPlayerListsRequest_Deposit()
        {
            List<string> player_info = new List<string>();
            string path = Path.GetTempPath() + @"\rtcgrab_rh_deposit.txt";

            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                byte[] result = await wc.DownloadDataTaskAsync("http://rh893sh3d.7799779.com/_xLhPnNlm9H0ifrRhPnNlm9H0ifrRZZ/adm/player/xml_player_search.aspx?str=&by=email&orderby=joindate&sortorder=desc&fromrow=0&torow=5000 ");
                string responsebody = Encoding.UTF8.GetString(result);
                var xDoc = XDocument.Parse(responsebody);

                var emptyElements = xDoc.Descendants("search");
                int i = 0;

                foreach (var xe in emptyElements)
                {
                    if (!File.Exists(path))
                    {
                        using (StreamWriter file = new StreamWriter(path, true, Encoding.UTF8))
                        {
                            file.WriteLine("test123*|*");
                            file.Close();
                        }
                    }
                    
                    string player_id = xe.Attribute("id").Value.Trim();
                    string username = xe.Attribute("username").Value.Trim();
                    string player_ldd = "";

                    if (username.ToLower() == Properties.Settings.Default.______last_registered_player.ToLower())
                    {
                        __detectInsert_deposit = true;
                    }
                    
                    bool isInsert = false;

                    if (__detectInsert_deposit)
                    {
                        using (StreamReader sr = File.OpenText(path))
                        {
                            string s = String.Empty;
                            while ((s = sr.ReadLine()) != null)
                            {
                                Application.DoEvents();

                                if (s == username)
                                {
                                    isInsert = true;
                                    break;
                                }
                                else
                                {
                                    isInsert = false;
                                }
                            }
                            sr.Close();
                        }

                        if (!isInsert)
                        {
                            player_ldd = await ___PlayerListLastDeposit_DepositAsync(player_id);
                        }
                    }
                    
                    if (username.ToLower() != Properties.Settings.Default.______last_registered_player_deposit.ToLower())
                    {
                        if (__detectInsert_deposit)
                        {
                            if (!isInsert)
                            {
                                if (!String.IsNullOrEmpty(player_ldd))
                                {
                                    player_info.Add(username + "*|*" + player_ldd);

                                    using (StreamWriter file = new StreamWriter(path, true, Encoding.UTF8))
                                    {
                                        file.WriteLine(username);
                                        file.Close();
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (__detectInsert_deposit)
                        {
                            if (!isInsert)
                            {
                                if (!String.IsNullOrEmpty(player_ldd))
                                {
                                    player_info.Add(username + "*|*" + player_ldd);

                                    using (StreamWriter file = new StreamWriter(path, true, Encoding.UTF8))
                                    {
                                        file.WriteLine(username);
                                        file.Close();
                                    }
                                }
                            }
                        }

                        if (player_info.Count != 0)
                        {
                            player_info.Reverse();
                            string player_info_get = String.Join(",", player_info);
                            string[] values = player_info_get.Split(',');
                            foreach (string value in values)
                            {
                                Application.DoEvents();
                                string[] values_inner = value.Split(new string[] { "*|*" }, StringSplitOptions.None);
                                int count = 0;
                                string _username = "";
                                string _date_deposit = "";

                                foreach (string value_inner in values_inner)
                                {
                                    count++;

                                    // Username
                                    if (count == 1)
                                    {
                                        _username = value_inner;
                                    }
                                    // Last Deposit Date
                                    else if (count == 2)
                                    {
                                        if (!String.IsNullOrEmpty(value_inner))
                                        {
                                            _date_deposit = value_inner;
                                        }
                                        else
                                        {
                                            _date_deposit = "";
                                        }
                                    }
                                }

                                Thread t = new Thread(delegate () { ___InsertData_Deposit(_username, _date_deposit, __brand_code); });
                                t.Start();

                                __send = 0;
                            }

                        }

                        player_info.Clear();
                        __isBreak_deposit = true;
                        __detectInsert_deposit = false;
                        break;
                    }
                }

                ___DepositLastRegistered();
                __isInsert_deposit = true;
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    ___GetPlayerListsRequest_Deposit();
                }
            }
        }

        private async Task<string> ___PlayerListLastDeposit_DepositAsync(string id)
        {
            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();
                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                string date_to = DateTime.Now.ToString("yyyy-MM-dd");
                string date_from = DateTime.Now.AddMonths(-1).ToString("yyyy-MM-dd");
                string responsebody = await wc.DownloadStringTaskAsync("http://rh893sh3d.7799779.com/_xLhPnNlm9H0ifrRhPnNlm9H0ifrRZZ/adm/player/inc_player_deposit_withdrawal.aspx?type=deposit&id=" + id + "&dateFrom=" + date_from + "&dateTo=" + date_to);
                if (responsebody.ToLower().Contains("no record"))
                {
                    return null;
                }
                else
                {
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(@responsebody);
                    var rows = doc.DocumentNode.SelectNodes("//*[@id='tblTrans']/tr");
                    int count = 0;
                    foreach (var row in rows)
                    {
                        string td = row.InnerHtml;
                        if (count != 0)
                        {
                            if (!td.ToLower().Contains("total"))
                            {
                                string date = row.SelectSingleNode("td[4]").InnerText;
                                string status = row.SelectSingleNode("td[5]").InnerText;
                                if (status.ToLower() == "processed")
                                {
                                    return date;
                                }
                                else
                                {
                                    return null;
                                }
                            }
                        }
                        count++;
                    }
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    await ___PlayerListLastDeposit_DepositAsync(id);
                }
            }

            return null;
        }

        private void ___InsertData_Deposit(string username, string last_deposit_date, string brand)
        {
            try
            {
                string password = username.ToLower() + last_deposit_date + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["username"] = username,
                        ["date_deposit"] = last_deposit_date,
                        ["brand_code"] = brand,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://192.168.10.252:8080/API/sendRTCdep", "POST", data);
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        ___InsertData2_Deposit(username, last_deposit_date, brand);
                    }
                }
            }
        }

        private void ___InsertData2_Deposit(string username, string last_deposit_date, string brand)
        {
            try
            {
                string password = username.ToLower() + last_deposit_date + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["username"] = username,
                        ["date_deposit"] = last_deposit_date,
                        ["brand_code"] = brand,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus.ssitex.com:8080/API/sendRTCdep", "POST", data);
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        ___InsertData_Deposit(username, last_deposit_date, brand);
                    }
                }
            }
        }

        private void ___DepositLastRegistered()
        {
            string path = Path.GetTempPath() + @"\rtcgrab_rh_deposit.txt";
            if (label_player_last_registered.Text != "-" && label_player_last_registered.Text.Trim() != "")
            {
                if (Properties.Settings.Default.______detect_deposit == "")
                {
                    DateTime today = DateTime.Now;
                    DateTime date = today.AddDays(1);
                    Properties.Settings.Default.______detect_deposit = date.ToString("yyyy-MM-dd 23");
                    Properties.Settings.Default.Save();
                }
                else
                {
                    DateTime today = DateTime.Now;
                    if (Properties.Settings.Default.______detect_deposit == today.ToString("yyyy-MM-dd HH"))
                    {
                        Properties.Settings.Default.______detect_deposit = "";
                        Properties.Settings.Default.______last_registered_player_deposit = label_player_last_registered.Text.Replace("Last Registered: ", "");
                        Properties.Settings.Default.Save();

                        if (File.Exists(path))
                        {
                            File.Delete(path);
                        }
                    }
                    else
                    {
                        string start_datetime = today.ToString("yyyy-MM-dd HH");
                        DateTime start = DateTime.ParseExact(start_datetime, "yyyy-MM-dd HH", CultureInfo.InvariantCulture);

                        string end_datetime = Properties.Settings.Default.______detect_deposit;
                        DateTime end = DateTime.ParseExact(end_datetime, "yyyy-MM-dd HH", CultureInfo.InvariantCulture);

                        if (start > end)
                        {
                            Properties.Settings.Default.______detect_deposit = "";
                            Properties.Settings.Default.______last_registered_player_deposit = label_player_last_registered.Text.Replace("Last Registered: ", "");
                            Properties.Settings.Default.Save();

                            if (File.Exists(path))
                            {
                                File.Delete(path);
                            }
                        }
                    }
                }
            }
        }

        private void SendMyBot(string message)
        {
            try
            {
                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                string urlString = "https://api.telegram.org/bot{0}/sendMessage?chat_id={1}&text={2}";
                string apiToken = "772918363:AAHn2ufmP3ocLEilQ1V-IHcqYMcSuFJHx5g";
                string chatId = "@allandrake";
                string text = "-----" + __brand_code + " " + __app + "-----%0A%0AIP:%20" + Properties.Settings.Default.______server_ip + "%0ALocation:%20" + Properties.Settings.Default.______server_location + "%0ADate%20and%20Time:%20[" + datetime + "]%0AMessage:%20" + message;
                urlString = String.Format(urlString, apiToken, chatId, text);
                WebRequest request = WebRequest.Create(urlString);
                Stream rs = request.GetResponse().GetResponseStream();
                StreamReader reader = new StreamReader(rs);
                string line = "";
                StringBuilder sb = new StringBuilder();
                while (line != null)
                {
                    line = reader.ReadLine();
                    if (line != null)
                        sb.Append(line);
                }
            }
            catch (Exception err)
            {
                if (err.ToString().ToLower().Contains("hexadecimal"))
                {
                    string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                    string urlString = "https://api.telegram.org/bot{0}/sendMessage?chat_id={1}&text={2}";
                    string apiToken = "772918363:AAHn2ufmP3ocLEilQ1V-IHcqYMcSuFJHx5g";
                    string chatId = "@allandrake";
                    string text = "-----" + __brand_code + " " + __app + "-----%0A%0AIP:%20192.168.10.60%0ALocation:%20192.168.10.60%0ADate%20and%20Time:%20[" + datetime + "]%0AMessage:%20" + message;
                    urlString = String.Format(urlString, apiToken, chatId, text);
                    WebRequest request = WebRequest.Create(urlString);
                    Stream rs = request.GetResponse().GetResponseStream();
                    StreamReader reader = new StreamReader(rs);
                    string line = "";
                    StringBuilder sb = new StringBuilder();
                    while (line != null)
                    {
                        line = reader.ReadLine();
                        if (line != null)
                            sb.Append(line);
                    }

                    __isClose = false;
                    Environment.Exit(0);
                }
                else
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        SendMyBot(message);
                    }
                }
            }
        }

        private void SendITSupport(string message)
        {
            if (__is_send)
            {
                try
                {
                    string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                    string urlString = "https://api.telegram.org/bot{0}/sendMessage?chat_id={1}&text={2}";
                    string apiToken = "612187347:AAE9doWWcStpWrDrfpOod89qGSxCJ5JwQO4";
                    string chatId = "@it_support_ssi";
                    string text = "-----" + __brand_code + " " + __app + "-----%0A%0AIP:%20" + Properties.Settings.Default.______server_ip + "%0ALocation:%20" + Properties.Settings.Default.______server_location + "%0ADate%20and%20Time:%20[" + datetime + "]%0AMessage:%20" + message;
                    urlString = String.Format(urlString, apiToken, chatId, text);
                    WebRequest request = WebRequest.Create(urlString);
                    Stream rs = request.GetResponse().GetResponseStream();
                    StreamReader reader = new StreamReader(rs);
                    string line = "";
                    StringBuilder sb = new StringBuilder();
                    while (line != null)
                    {
                        line = reader.ReadLine();
                        if (line != null)
                        {
                            sb.Append(line);
                        }
                    }
                }
                catch (Exception err)
                {
                    if (err.ToString().ToLower().Contains("hexadecimal"))
                    {
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        string urlString = "https://api.telegram.org/bot{0}/sendMessage?chat_id={1}&text={2}";
                        string apiToken = "612187347:AAE9doWWcStpWrDrfpOod89qGSxCJ5JwQO4";
                        string chatId = "@it_support_ssi";
                        string text = "-----" + __brand_code + " " + __app + "-----%0A%0AIP:%20192.168.10.60%0ALocation:%20192.168.10.60%0ADate%20and%20Time:%20[" + datetime + "]%0AMessage:%20" + message;
                        urlString = String.Format(urlString, apiToken, chatId, text);
                        WebRequest request = WebRequest.Create(urlString);
                        Stream rs = request.GetResponse().GetResponseStream();
                        StreamReader reader = new StreamReader(rs);
                        string line = "";
                        StringBuilder sb = new StringBuilder();
                        while (line != null)
                        {
                            line = reader.ReadLine();
                            if (line != null)
                            {
                                sb.Append(line);
                            }
                        }

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        __send++;
                        if (__send == 5)
                        {
                            SendITSupport("There's a problem to the server, please re-open the application.");
                            SendMyBot(err.ToString());

                            __isClose = false;
                            Environment.Exit(0);
                        }
                        else
                        {
                            ___WaitNSeconds(10);
                            SendITSupport(message);
                        }
                    }
                }
            }
        }

        private async Task ___GetLastRegisteredPlayerAsync()
        {
            try
            {
                string password = __brand_code.ToString() + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["token"] = token
                    };

                    byte[] result = await wb.UploadValuesTaskAsync("http://192.168.10.252:8080/API/lastRTCrecord", "POST", data);
                    string responsebody = Encoding.UTF8.GetString(result);
                    var deserializeObject = JsonConvert.DeserializeObject(responsebody);
                    JObject jo = JObject.Parse(deserializeObject.ToString());
                    JToken plr = jo.SelectToken("$.msg");
                    Properties.Settings.Default.______last_registered_player = plr.ToString();
                    Properties.Settings.Default.______last_registered_player_deposit = plr.ToString();
                    Properties.Settings.Default.Save();
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        await ___GetLastRegisteredPlayer2Async();
                    }
                }
            }
        }

        private async Task ___GetLastRegisteredPlayer2Async()
        {
            try
            {
                string password = __brand_code + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["token"] = token
                    };

                    var result = await wb.UploadValuesTaskAsync("http://zeus.ssitex.com:8080/API/lastRTCrecord", "POST", data);
                    string responsebody = Encoding.UTF8.GetString(result);
                    var deserializeObject = JsonConvert.DeserializeObject(responsebody);
                    JObject jo = JObject.Parse(deserializeObject.ToString());
                    JToken plr = jo.SelectToken("$.msg");

                    Properties.Settings.Default.______last_registered_player = plr.ToString();
                    Properties.Settings.Default.______last_registered_player_deposit = plr.ToString();
                    Properties.Settings.Default.Save();
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        await ___GetLastRegisteredPlayerAsync();
                    }
                }
            }
        }

        private void timer_flush_memory_Tick(object sender, EventArgs e)
        {
            FlushMemory();
        }

        public static void FlushMemory()
        {
            Process prs = Process.GetCurrentProcess();
            try
            {
                prs.MinWorkingSet = (IntPtr)(300000);
            }
            catch (Exception err)
            {
                // leave blank
            }
        }

        private double __total_records_mb;
        private double __display_length_mb = 5000;
        private int __total_page_mb;
        private JObject __jo_mb;
        private int __result_count_json_mb;
        private bool __inserted_in_excel_mb = true;
        private bool __detect_mb = false;
        private bool __isSending = false;
        private int __i_mb = 0;
        private int __ii_mb = 0;
        private int __pages_count_display_mb = 0;
        private int __test_gettotal_count_record_mb;
        private int __get_ii_mb = 1;
        private int __get_ii_display_mb = 1;
        private int __pages_count_mb = 0;
        private string __shared_path = "\\\\192.168.10.22\\ssi-reporting\\";
        private string __file_name = "";
        private string __task_id = "";
        StringBuilder __csv_mb = new StringBuilder();

        private async void ___GetMABListsAsync()
        {
            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                byte[] result = await wc.DownloadDataTaskAsync("http://rh893sh3d.7799779.com/_xLhPnNlm9H0ifrRhPnNlm9H0ifrRZZ/adm/player/xml_player_search.aspx?str=&by=email&orderby=joindate&sortorder=desc&fromrow=0&torow=500000");
                string responsebody = Encoding.UTF8.GetString(result);
                var xDoc = XDocument.Parse(responsebody);

                var emptyElements = xDoc.Descendants("search");
                int count = 0;

                foreach (var xe in emptyElements)
                {
                    string player_id = xe.Attribute("id").Value.Trim();
                    string username = xe.Attribute("username").Value.Trim();
                    string result_count = xe.Attribute("resultCount").Value.Trim();
                    string mab = await ___PlayerListMABAsync(player_id);

                    count++;

                    if (count == 1)
                    {
                        var header = string.Format("{0},{1},{2}", "Brand", "Username", "Main Account Balance");
                        __csv_mb.AppendLine(header);
                    }

                    var newLine = string.Format("{0},{1},{2}", __brand_code, "\"" + username + "\"", "\"" + mab + "\"");
                    __csv_mb.AppendLine(newLine);

                    label_currentrecord.Text = Convert.ToInt32(count).ToString("N0") + " of " + Convert.ToInt32(result_count).ToString("N0");
                    label_currentrecord.Invalidate();
                    label_currentrecord.Update();
                }

                ___PlayerListInsertDoneMABAsync();
            }
            catch (Exception err)
            {
                __send++;
                if (__send == 5)
                {
                    SendITSupport("There's a problem to the server, please re-open the application.");
                    SendMyBot(err.ToString());

                    __isClose = false;
                    Environment.Exit(0);
                }
                else
                {
                    ___WaitNSeconds(10);
                    ___GetMABListsAsync();
                }
            }
        }

        private async Task<string> ___PlayerListMABAsync(string id)
        {
            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();
                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                string date_to = DateTime.Now.ToString("yyyy-MM-dd");
                string date_from = DateTime.Now.AddMonths(-1).ToString("yyyy-MM-dd");
                string responsebody = await wc.DownloadStringTaskAsync("http://rh893sh3d.7799779.com/_xLhPnNlm9H0ifrRhPnNlm9H0ifrRZZ/adm/player/inc_player_profile.aspx?id=" + id);
                
                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(@responsebody);
                var rows = doc.DocumentNode.SelectNodes("//*[@class='listTable2']/tr");

                foreach (var row in rows)
                {
                    string td = row.InnerHtml;
                    string mab_innettext = row.SelectSingleNode("td[1]").InnerText;
                    if (mab_innettext.ToLower().Trim() == "account balance:")
                    {
                        string mab = row.SelectSingleNode("td[2]").InnerText.Replace("&nbsp;", "").Trim();
                        return mab;
                    }
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    await ___PlayerListMABAsync(id);
                }
            }

            return null;
        }
        
        private async void ___PlayerListInsertDoneMABAsync()
        {
            try
            {
                string _path = "\\\\192.168.10.252\\Balance$\\";
                string _current_datetime = DateTime.Now.ToString("yyyy-MM-ddHHmm");
                __file_name = __brand_code + "_" + _current_datetime;
                string _folder_path_result = _path + __brand_code + "_" + _current_datetime + ".txt";
                string _folder_path_result_xlsx = _path + __brand_code + "_" + _current_datetime + ".xlsx";

                if (File.Exists(_folder_path_result))
                {
                    File.Delete(_folder_path_result);
                }

                if (File.Exists(_folder_path_result_xlsx))
                {
                    File.Delete(_folder_path_result_xlsx);
                }

                __csv_mb.ToString().Reverse();
                File.WriteAllText(_folder_path_result, __csv_mb.ToString(), Encoding.UTF8);

                Excel.Application app = new Excel.Application();
                Excel.Workbook wb = app.Workbooks.Open(_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Worksheet worksheet = wb.ActiveSheet;
                worksheet.Activate();
                worksheet.Application.ActiveWindow.SplitRow = 1;
                worksheet.Application.ActiveWindow.FreezePanes = true;
                Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                firstRow.AutoFilter(1,
                                    Type.Missing,
                                    Excel.XlAutoFilterOperator.xlAnd,
                                    Type.Missing,
                                    true);
                Excel.Range usedRange = worksheet.UsedRange;
                Excel.Range rows = usedRange.Rows;
                int count = 0;
                foreach (Excel.Range row in rows)
                {
                    if (count == 0)
                    {
                        Excel.Range firstCell = row.Cells[1];

                        string firstCellValue = firstCell.Value as String;

                        if (!string.IsNullOrEmpty(firstCellValue))
                        {
                            row.Interior.Color = Color.FromArgb(155, 132, 53);
                            row.Font.Color = Color.FromArgb(255, 255, 255);
                        }

                        break;
                    }

                    count++;
                }
                int i;
                for (i = 1; i <= 3; i++)
                {
                    worksheet.Columns[i].ColumnWidth = 22;
                }
                wb.SaveAs(_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wb.Close();
                app.Quit();
                Marshal.ReleaseComObject(app);

                if (File.Exists(_folder_path_result))
                {
                    File.Delete(_folder_path_result);
                }

                __csv_mb.Clear();
                __total_records_mb = 0;
                __display_length_mb = 5000;
                __total_page_mb = 0;
                __result_count_json_mb = 0;
                __inserted_in_excel_mb = true;
                __detect_mb = false;
                __i_mb = 0;
                __ii_mb = 0;
                __pages_count_display_mb = 0;
                __test_gettotal_count_record_mb = 0;
                __get_ii_mb = 1;
                __get_ii_display_mb = 1;
                __pages_count_mb = 0;
                label_currentrecord.Text = "";
                label_page_count.Text = "";

                // send
                await ___SetTaskStatusAsync(__task_id, __file_name);
                timer_mb_detect.Start();
            }
            catch (Exception err)
            {
                __send++;
                if (__send == 5)
                {
                    SendITSupport("There's a problem to the server, please re-open the application.");
                    SendMyBot(err.ToString());

                    __isClose = false;
                    Environment.Exit(0);
                }
                else
                {
                    ___WaitNSeconds(10);
                    ___GetTaskStatus();
                }
            }
        }

        private void timer_mb_detect_Tick(object sender, EventArgs e)
        {
            ___GetTaskStatus();
        }

        private void ___GetTaskStatus()
        {
            try
            {
                timer_mb_detect.Stop();
                string password = __brand_code + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://192.168.10.252:8080/API/getBalanceTaskStatus", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                    var deserializeObject = JsonConvert.DeserializeObject(responseInString);
                    JObject jo_mb = JObject.Parse(deserializeObject.ToString());
                    JToken status = jo_mb.SelectToken("$.status");
                    JToken task_id = jo_mb.SelectToken("$.task_id");
                    __task_id = task_id.ToString();

                    if (status.ToString() == "1")
                    {
                        if (webBrowser.Url.ToString().Contains("adm/default.aspx"))
                        {
                            // start
                            timer_mb_detect.Stop();
                            ___UpdateTaskStatus();
                            ___GetMABListsAsync();
                        }
                        else
                        {
                            timer_mb_detect.Start();
                        }
                    }
                    else
                    {
                        timer_mb_detect.Start();
                    }
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        ___GetTaskStatus2();
                    }
                }
            }
        }

        private void ___GetTaskStatus2()
        {
            try
            {
                timer_mb_detect.Stop();
                string password = __brand_code + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus.ssitex.com:8080/API/getBalanceTaskStatus", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                    var deserializeObject = JsonConvert.DeserializeObject(responseInString);
                    JObject jo_mb = JObject.Parse(deserializeObject.ToString());
                    JToken status = jo_mb.SelectToken("$.status");
                    JToken task_id = jo_mb.SelectToken("$.task_id");
                    __task_id = task_id.ToString();

                    if (status.ToString() == "1")
                    {
                        if (webBrowser.Url.ToString().Contains("adm/default.aspx"))
                        {
                            // start
                            timer_mb_detect.Stop();
                            ___UpdateTaskStatus();
                            ___GetMABListsAsync();
                        }
                        else
                        {
                            timer_mb_detect.Start();
                        }
                    }
                    else
                    {
                        timer_mb_detect.Start();
                    }
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        ___GetTaskStatus();
                    }
                }
            }
        }

        private async Task ___SetTaskStatusAsync(string task_id, string file_name)
        {
            try
            {
                string password = file_name + task_id + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["task_id"] = task_id,
                        ["filename"] = file_name,
                        ["token"] = token
                    };

                    var response = await wb.UploadValuesTaskAsync("http://192.168.10.252:8080/API/setBalanceTaskStatus", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);

                    __file_name = "";
                    __task_id = "";
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        await ___SetTaskStatus2Async(task_id, file_name);
                    }
                }
            }
        }

        private async Task ___SetTaskStatus2Async(string task_id, string file_name)
        {
            try
            {
                string password = file_name + task_id + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["task_id"] = task_id,
                        ["filename"] = file_name,
                        ["token"] = token
                    };

                    var response = await wb.UploadValuesTaskAsync("http://zeus.ssitex.com:8080/API/setBalanceTaskStatus", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);

                    __file_name = "";
                    __task_id = "";
                    timer_mb_detect.Start();
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        await ___SetTaskStatusAsync(task_id, file_name);
                    }
                }
            }
        }

        private void ___UpdateTaskStatus()
        {
            try
            {
                string password = __brand_code + __task_id + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["task_id"] = __task_id,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://192.168.10.252:8080/API/updBalanceTaskStatus", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        ___UpdateTaskStatus2();
                    }
                }
            }
        }

        private void ___UpdateTaskStatus2()
        {
            try
            {
                string password = __brand_code + __task_id + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["task_id"] = __task_id,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus.ssitex.com:8080/API/updBalanceTaskStatus", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        ___UpdateTaskStatus();
                    }
                }
            }
        }

        private void ___Send()
        {
            try
            {
                string path = Path.GetTempPath() + @"\rtcgrab_sending_rh.txt";
                string path_temp = Path.GetTempPath() + @"\rtcgrab_temp_sending_rh.txt";
                if (Properties.Settings.Default.______count_player == 5)
                {
                    __isSending = true;

                    Properties.Settings.Default.______count_player = 0;
                    Properties.Settings.Default.Save();
                    string line;
                    char[] split = "*|*".ToCharArray();
                    StreamReader file = new StreamReader(path, Encoding.UTF8);
                    while ((line = file.ReadLine()) != null)
                    {
                        if (line.Trim() != "")
                        {
                            string[] data = line.Split(split);
                            ___InsertData(data[0], data[3], data[6], data[9], data[12], data[15], data[18], data[21], data[24]);
                            __send = 0;
                        }
                    }

                    File.WriteAllText(path, string.Empty);

                    if (File.Exists(path_temp))
                    {
                        File.Copy(path_temp, path);
                    }

                    __isSending = false;
                }
            }
            catch (Exception err)
            {
                SendMyBot(err.ToString());
            }
        }

        private void panel1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (__is_send)
            {
                __is_send = false;
                MessageBox.Show("Telegram Notification is Disabled.", __brand_code + " " + __app, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                __is_send = true;
                MessageBox.Show("Telegram Notification is Enabled.", __brand_code + " " + __app, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void timer_detect_running_Tick(object sender, EventArgs e)
        {
            //___DetectRunning();
        }

        private void ___DetectRunning()
        {
            try
            {
                string datetime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string password = __brand_code + datetime + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["app_type"] = __app_type,
                        ["last_update"] = datetime,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://192.168.10.252:8080/API/updateAppStatus", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());
                        __send = 0;

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        ___DetectRunning2();
                    }
                }
            }
        }

        private void ___DetectRunning2()
        {
            try
            {
                string datetime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string password = __brand_code + datetime + "youdieidie";
                byte[] encodedPassword = new UTF8Encoding().GetBytes(password);
                byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedPassword);
                string token = BitConverter.ToString(hash)
                   .Replace("-", string.Empty)
                   .ToLower();

                using (var wb = new WebClient())
                {
                    var data = new NameValueCollection
                    {
                        ["brand_code"] = __brand_code,
                        ["app_type"] = __app_type,
                        ["last_update"] = datetime,
                        ["token"] = token
                    };

                    var response = wb.UploadValues("http://zeus.ssitex.com:8080/API/updateAppStatus", "POST", data);
                    string responseInString = Encoding.UTF8.GetString(response);
                }
            }
            catch (Exception err)
            {
                if (__isLogin)
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __isClose = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        ___DetectRunning();
                    }
                }
            }
        }

        private void ___WaitNSeconds(int sec)
        {
            if (sec < 1) return;
            DateTime _desired = DateTime.Now.AddSeconds(sec);
            while (DateTime.Now < _desired)
            {
                Application.DoEvents();
            }
        }
    }
}