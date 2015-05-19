using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Windows.Forms;
using ExcelLibrary.SpreadSheet;
using Newtonsoft.Json.Linq;
using Quobject.SocketIoClientDotNet.Client;

namespace StreamTipUpdater
{
    internal class Tip
    {
        public string Id;
        public string Processor;
        public string TransactionId;
        public string FirstName;
        public string LastName;
        public string Username;
        public string Email;
        public string CurrencyCode;
        public string CurrencySymbol;
        public int Cents;
        public string Note;
        public bool Pending;
        public bool Reversed;
        public bool Deleted;
        public DateTime Date;
        public decimal Amount;

        public Tip(JToken tip)
        {
            try
            {
                Id = tip.Value<string>("id");
                Processor = tip.Value<string>("processor");
                TransactionId = tip.Value<string>("transactionId");
                FirstName = tip.Value<string>("firstName");
                LastName = tip.Value<string>("lastName");
                Username = tip.Value<string>("username");
                Email = tip.Value<string>("email");
                CurrencyCode = tip.Value<string>("currentCode");
                CurrencySymbol = tip.Value<string>("currencySymbol");
                Cents = tip.Value<int>("cents");
                Note = tip.Value<string>("note");
                Pending = tip.Value<bool>("pending");
                Reversed = tip.Value<bool>("reversed");
                Deleted = tip.Value<bool>("deleted");
                Date = tip.Value<DateTime>("date");
                Amount = tip.Value<decimal>("amount");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Exception parsing Tip: " + ex.Message, "Exception", MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
            }
        }
    }

    internal class StreamTipUpdater : Form
    {
        private readonly NotifyIcon  trayIcon;

        private readonly string _clientId;
        private readonly string _accessToken;

        private readonly Dictionary<string, Tip> _tips = new Dictionary<string, Tip>();

        private readonly object _lockObject = new object();

        private readonly string _donationXls = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "donations.xls");

        private DateTime FirstOfMonth 
        {
            get
            {
                var now = DateTime.Now;
                return new DateTime(now.Year, now.Month, 1);
            }
        }

        private DateTime LastOfMonth
        {
            get
            {
                var now = DateTime.Now;
                var firstDayOfMonth = new DateTime(now.Year, now.Month, 1);
                return firstDayOfMonth.AddMonths(1);
            }
        }

        public StreamTipUpdater()
        {
            // Create a simple tray menu with only one item.
            ContextMenu trayMenu = new ContextMenu();
            trayMenu.MenuItems.Add("Exit", OnExit);
 
            // Create a tray icon. In this example we use a
            // standard system icon for simplicity, but you
            // can of course use your own custom icon too.
            trayIcon = new NotifyIcon
            {
                Text = "Stream Tip",
                //Icon = Icon.FromHandle(Properties.Resources.ResourceManager.)
                Icon = Properties.Resources.TrayIcon,
                ContextMenu = trayMenu,
                Visible = true
            };

            trayMenu.MenuItems.Add("Export to Excel", (sender, args) =>
            {
                GenerateExcelFile();
                try
                {
                    Process.Start("explorer.exe", _donationXls);
                }
                catch (Win32Exception ex)
                {
                }
            });

            var clientId = ConfigurationManager.AppSettings["client_id"];
            var accessToken = ConfigurationManager.AppSettings["access_token"];

            if (string.IsNullOrWhiteSpace(clientId) || string.IsNullOrWhiteSpace(accessToken))
            {
                MessageBox.Show(
                    "Please make sure your client ID and access Token are setup correctly in the application config file.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            _clientId = clientId.Trim();
            _accessToken = accessToken.Trim();
            
            // Add menu to tray icon and show it.
        }

        private void AddTip(JToken jTip)
        {
            lock (_lockObject)
            {
                var tip = new Tip(jTip);

                _tips.Add(tip.Id, tip);

            }
        }

        protected override void OnLoad(EventArgs e)
        {
            Visible = false; // Hide form window.
            ShowInTaskbar = false; // Remove from taskbar.
            
            if (string.IsNullOrWhiteSpace(_clientId) || string.IsNullOrWhiteSpace(_accessToken))
            {
                Application.Exit();
                return;
            }

            base.OnLoad(e);

            InitTips();
            WriteObs();
            StartSocket();
        }

        private void GenerateExcelFile()
        {
            //create new xls file
            var workbook = new Workbook();

            var allTips = _tips.Values
                .OrderByDescending(t => t.Date)
                .GroupBy(t => new DateTime(t.Date.Year, t.Date.Month, 1));

            foreach (var grouping in allTips)
            {
                var worksheet = new Worksheet(string.Format("{0:MMMM yyyy}", grouping.Key));

                var tips = grouping
                    .GroupBy(t => t.Username)
                    .Select(grp => new
                    {
                        Username = grp.Key,
                        Amount = grp.Sum(t => t.Amount)
                    })
                    .ToArray();

                for (var i = 0; i < tips.Length; i++)
                {
                    worksheet.Cells[i, 0] = new Cell(tips[i].Username);
                    worksheet.Cells[i, 1] = new Cell(tips[i].Amount, new CellFormat(CellFormatType.Currency, "$0.00"));
                }
                worksheet.Cells.ColumnWidth[0, 1] = 3000;

                workbook.Worksheets.Add(worksheet);
            }
            try
            {
                workbook.Save(_donationXls);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An exception occurred trying to save the excel file: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void WriteObs()
        {
            var directory = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "obs");

            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            lock (_lockObject)
            {
                File.WriteAllText(Path.Combine(directory, "last25.txt"), String.Join(" ",
                    _tips.Values
                        .Where(t => !t.Deleted && !t.Reversed)
                        .OrderByDescending(t => t.Date)
                        .Take(25)
                        .Select(t => string.Format("{0}: {1}{2}", t.Username, t.CurrencySymbol, t.Amount))
                    ));

                var topTip = _tips.Values
                    .Where(t => t.Date > FirstOfMonth && t.Date < LastOfMonth)
                    .GroupBy(t => t.Username)
                    .Select(grp => new
                    {
                        Username = grp.Key,
                        Amount = grp.Sum(t => t.Amount)
                    })
                    .OrderByDescending(a => a.Amount)
                    .FirstOrDefault();

                var topTipString = topTip != null ? string.Format("{0}: ${1}", topTip.Username, topTip.Amount) : ConfigurationManager.AppSettings["topTipNone"];

                File.WriteAllText(Path.Combine(directory, "topthismonth.txt"), topTipString);
            }
        }

        private void InitTips()
        {
            var offset = 0;

            while (true)
            {
                var uriBuilder = new UriBuilder("https://streamtip.com/api/tips");
                var query = HttpUtility.ParseQueryString(uriBuilder.Query);
                query["client_id"] = _clientId;
                query["access_token"] = _accessToken;
                query["limit"] = "100";
                query["offset"] = offset.ToString();
                query["sort_by"] = "date";
                query["direction"] = "desc";

                uriBuilder.Query = query.ToString();

                var req = WebRequest.CreateHttp(uriBuilder.Uri);

                using (var response = req.GetResponse())
                {
                    using (var stream = response.GetResponseStream())
                    {
                        using (var reader = new StreamReader(stream))
                        {
                            var json = reader.ReadToEnd();

                            var jo = JObject.Parse(json);

                            var tips = jo.Value<JArray>("tips");

                            if (!tips.Any())
                            {
                                return;
                            }

                            foreach (var tip in tips )
                            {
                                AddTip(tip);
                            }
                        }
                    }
                }

                offset += 100;
            }
        }

        private void StartSocket()
        {
            var socket = IO.Socket("https://streamtip.com", new IO.Options
            {
                Query = new Dictionary<string, string>
                {
                    {"client_id", _clientId},
                    {"access_token", _accessToken}
                }
            });

            socket.On(Socket.EVENT_CONNECT, () =>
            {
            });


            socket.On("authenticated",
                (data) =>
                    trayIcon.ShowBalloonTip(2000, "Connected", "Streamtip is connected and listening for new donations.",
                        ToolTipIcon.Info));

            socket.On("error", (err) => {
                if (err.ToString() == "401::Access Denied::")
                {
                    MessageBox.Show(
                        "Invalid client_id or access_token",
                        "Error", MessageBoxButtons.OK);
                }
                else
                {
                    MessageBox.Show(err.ToString(), "Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            });

            socket.On(Socket.EVENT_DISCONNECT, (data) =>
            {
            });

            socket.On("newTip", (data) =>
            {
                var jo = data as JToken;

                if (jo != null)
                {
                    AddTip(jo);
                    WriteObs();
                }
            });
        }

        private void OnExit(object sender, EventArgs e)
        {
            Application.Exit();
        }
 
        protected override void Dispose(bool isDisposing)
        {
            if (isDisposing)
            {
                // Release the icon resource.
                trayIcon.Dispose();
            }
 
            base.Dispose(isDisposing);
        }
    }
}
