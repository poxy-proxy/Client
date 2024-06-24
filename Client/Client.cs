using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using NetFwTypeLib;
using System.Net.NetworkInformation;
using System.Data;
using System.Linq;
using System.Management;


namespace Client
{
    public partial class Client : Form
    {
        private bool connected = false;
        private Thread client = null;
        static bool FlashIsEnable = true;
        static bool IndivFlashOption = false;
        private struct MyClient
        {
            public string username;
            public string key;
            public TcpClient client;
            public NetworkStream stream;
            public byte[] buffer;
            public StringBuilder data;
            public EventWaitHandle handle;
        };
        private MyClient obj;
        private Task send = null;
        private bool exit = false;
        private readonly Stopwatch stopwatch = new Stopwatch();

        public Client()
        {
            InitializeComponent();
            WqlEventQuery query = new WqlEventQuery("SELECT * FROM __InstanceOperationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_PnPEntity'");
            ManagementEventWatcher watcher = new ManagementEventWatcher(query);
            watcher.EventArrived += new EventArrivedEventHandler(DeviceChangedEvent);
            watcher.Start();
            AllowInPort();
            // делаем невидимой нашу иконку в трее
            notifyIcon1.Visible = false;
            // добавляем Эвент или событие по 2му клику мышки, 
            //вызывая функцию  notifyIcon1_MouseDoubleClick
            this.notifyIcon1.MouseDoubleClick += new MouseEventHandler(notifyIcon1_MouseDoubleClick);

            // добавляем событие на изменение окна
            this.Resize += new System.EventHandler(this.Form1_Resize);
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            // проверяем наше окно, и если оно было свернуто, делаем событие        
            if (WindowState == FormWindowState.Minimized)
            {
                // прячем наше окно из панели
                this.ShowInTaskbar = false;
                // делаем нашу иконку в трее активной
                notifyIcon1.Visible = true;
            }
        }


        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // делаем нашу иконку скрытой
            notifyIcon1.Visible = false;
            // возвращаем отображение окна в панели
            this.ShowInTaskbar = true;
            //разворачиваем окно
            WindowState = FormWindowState.Normal;
        }

        public static void AllowInPort(string port = "7000")
        {
            INetFwRule firewallRule = (INetFwRule)Activator.CreateInstance(
               Type.GetTypeFromProgID("HNetCfg.FWRule"));
            // firewallRule.RemoteAddresses = ra;
            firewallRule.Protocol = (int)ProtocolType.Tcp;
            firewallRule.LocalPorts = port;
            firewallRule.Action = NET_FW_ACTION_.NET_FW_ACTION_ALLOW;
            firewallRule.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN;
            firewallRule.Enabled = true;
            firewallRule.InterfaceTypes = "All";
            firewallRule.Name = "Allow port IN " + port;

            INetFwPolicy2 firewallPolicy = (INetFwPolicy2)Activator.CreateInstance(
                Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
            firewallPolicy.Rules.Add(firewallRule);
        }

        static void DeviceChangedEvent(object sender, EventArrivedEventArgs e)
        {
            ManagementBaseObject instance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
            string eventType = (string)e.NewEvent["__CLASS"];

            if (instance != null && instance["Caption"] != null)
            {
                var pars=String.Format("Name: {0}, Status: {1}, Instance Path: {2}, GUID: {3}",
                      instance.GetPropertyValue("Caption"),
                      instance.GetPropertyValue("Status"),
                      instance.GetPropertyValue("DeviceID"),
                      instance.GetPropertyValue("ClassGuid"));
                string deviceName = instance["Caption"].ToString();

                // Проверяем, является ли подключенное устройство USB-флеш-накопителем
                if (deviceName.Contains("Запоминающ") || deviceName.Contains("Запоминающее устройство для USB") || deviceName.Contains("ADB") || deviceName.Contains("память"))
                {
                    //if (eventType.Equals("__InstanceCreationEvent"))
                    //{
                    //    Console.WriteLine("USB-флеш-накопитель подключен: " + deviceName);
                    //    // Здесь можно вставить код блокировки или другой логики по обработке подключения
                    //}
                    //else if (eventType.Equals("__InstanceDeletionEvent"))
                    //{
                    //    Console.WriteLine("USB-флеш-накопитель отключен: " + deviceName);
                    //    // Здесь можно вставить код разблокировки или другой логики по обработке отключения
                    //}

                    //Console.WriteLine("USB-флеш-накопитель подключен: " + deviceName);
                    if (!FlashIsEnable)
                    {
                        DisableFlashUSB(deviceName);
                    }
                    else
                    {
                        EnableFlashUSB(deviceName);
                    }
                    //Console.WriteLine("Флеш-накопитель заблокирован");
                    //EnableFlashUSB(deviceName);
                    //Console.WriteLine("Флеш-накопитель разблокирован");
                }
            }
        }

        private void Log(string msg = "") // clear the log if message is not supplied or is empty
        {
            if (!exit)
            {
                logTextBox.Invoke((MethodInvoker)delegate
                {
                    if (msg.Length > 0)
                    {
                        logTextBox.AppendText(string.Format("[ {0} ] {1}{2}", DateTime.Now.ToString("HH:mm"), msg, Environment.NewLine));
                    }
                    else
                    {
                        logTextBox.Clear();
                    }
                });
            }
        }

        private string ErrorMsg(string msg)
        {
            return string.Format("ERROR: {0}", msg);
        }

        private string SystemMsg(string msg)
        {
            return string.Format("SYSTEM: {0}", msg);
        }

        private void Connected(bool status)
        {
            if (!exit)
            {
                connectButton.Invoke((MethodInvoker)delegate
                {
                    connected = status;
                    if (status)
                    {
                        addrTextBox.Enabled = false;
                        portTextBox.Enabled = false;
                        //usernameTextBox.Enabled = false;
                        //keyTextBox.Enabled = false;
                        connectButton.Text = "Отключиться";
                        Log(SystemMsg("Вы подключены"));
                    }
                    else
                    {
                        addrTextBox.Enabled = true;
                        portTextBox.Enabled = true;
                        //usernameTextBox.Enabled = true;
                        //keyTextBox.Enabled = true;
                        connectButton.Text = "Подключиться";
                        Log(SystemMsg("Вы отключены"));
                    }
                });
            }
        }

        private void Read(IAsyncResult result)
        {
            int bytes = 0;
            if (obj.client.Connected)
            {
                try
                {
                    bytes = obj.stream.EndRead(result);
                }
                catch (Exception ex)
                {
                    Log(ErrorMsg(ex.Message));
                }
            }
            if (bytes > 0)
            {
                obj.data.AppendFormat("{0}", Encoding.UTF8.GetString(obj.buffer, 0, bytes));
                try
                {
                    if (obj.stream.DataAvailable)
                    {
                        obj.stream.BeginRead(obj.buffer, 0, obj.buffer.Length, new AsyncCallback(Read), null);
                    }
                    else
                    {
                 
                        string logs = obj.data.ToString();
                        Log(logs);
                        if(logs.Contains("Запретить весь интернет на всех компьютерах"))
                        {
                            AllowSite("");
                        }
                        if(logs.Contains("Разрешить весь интернет на всех компьютерах"))
                        {
                            AllowInternet();
                        }
                        if (logs.Contains("Запретить доступ в интернет, кроме сайтов:"))
                        {
                            var message = logs.Split(':');
                            AllowSite(message[2]);
                        }
                        if (logs.Contains("Запретить подключение флеш-накопителей на всех комьютерах") && !IndivFlashOption)
                        {
                            DisableFlashUSB("Запоминающее устройство");

                            DisableFlashUSB("Запоминающее устройство для USB");

                            DisableFlashUSB("ADB");
                            DisableFlashUSB("память");

                            FlashIsEnable = false;

                            //DisableFlashUSB("Память");
                        }
                        if (logs.Contains("Разрешить подключение флеш-накопителей на всех комьютерах") && !IndivFlashOption)
                        {
                            EnableFlashUSB("Запоминающее устройство");
                            EnableFlashUSB("Запоминающее устройство для USB");
                            EnableFlashUSB("ADB");
                            EnableFlashUSB("память");
                            FlashIsEnable = true;

                        }
                        string host = Dns.GetHostName();
                        IPAddress[] addresses = Dns.GetHostAddresses(host);
                        foreach(var adr in addresses)
                        {
                            if (logs.Contains(adr.ToString()))
                            {
                                if(logs.Contains("кроме сайтов"))
                                {
                                    var message = logs.Split(':');
                                    AllowSite(message[2]);
                                }

                                if (logs.Contains("Интернет запрещён"))
                                {

                                    AllowSite("");
                                }

                                if (logs.Contains("Интернет разрешён"))
                                {

                                    AllowInternet();
                                }

                                if (logs.Contains("Флеш-накопители разрешены"))
                                {
                                    EnableFlashUSB("Запоминающее устройство");
                                    EnableFlashUSB("Запоминающее устройство для USB");
                                    EnableFlashUSB("ADB");
                                    EnableFlashUSB("память");
                                    FlashIsEnable = true;
                                }


                                if (logs.Contains("Флеш-накопители запрещены"))
                                {
                                    DisableFlashUSB("Запоминающее устройство");
                                    DisableFlashUSB("Запоминающее устройство для USB");
                                    DisableFlashUSB("ADB");
                                    DisableFlashUSB("память");
                                    FlashIsEnable = false;
                                }

                                if(logs.Contains("Индивидуальные опции"))
                                {
                                    IndivFlashOption = true;
                                }

                                if (logs.Contains("Опции по умолчанию"))
                                {
                                    IndivFlashOption = false;
                                }
                            }

                         
                        }
                        obj.data.Clear();
                        obj.handle.Set();
                    }
                }
                catch (Exception ex)
                {
                    obj.data.Clear();
                    Log(ErrorMsg(ex.Message));
                    obj.handle.Set();
                }
            }
            else
            {
                obj.client.Close();
                obj.handle.Set();
            }
        }



        static void DisableFlashUSB(string DeviceName)
        {
            
            ManagementObjectSearcher myDevices = new ManagementObjectSearcher("root\\CIMV2", @"SELECT * FROM Win32_PnPEntity where Name Like " + '"'+"%" + DeviceName +"%"+ '"');

            foreach (ManagementObject item in myDevices.Get())
            {
                ManagementBaseObject inParams = item.InvokeMethod("Disable", null, null);
            }
          
        }

        static void EnableFlashUSB(string DeviceName)
        {
          
            ManagementObjectSearcher myDevices = new ManagementObjectSearcher("root\\CIMV2", @"SELECT * FROM Win32_PnPEntity where Name Like " + '"' + "%" + DeviceName + "%" + '"');

            foreach (ManagementObject item in myDevices.Get())
            {
                ManagementBaseObject inParams = item.InvokeMethod("Enable", null, null);
            }

           
        }

        private void ReadAuth(IAsyncResult result)
        {
            int bytes = 0;
            if (obj.client.Connected)
            {
                try
                {
                    bytes = obj.stream.EndRead(result);
                }
                catch (Exception ex)
                {
                    Log(ErrorMsg(ex.Message));
                }
            }
            if (bytes > 0)
            {
                obj.data.AppendFormat("{0}", Encoding.UTF8.GetString(obj.buffer, 0, bytes));
                try
                {
                    if (obj.stream.DataAvailable)
                    {
                        obj.stream.BeginRead(obj.buffer, 0, obj.buffer.Length, new AsyncCallback(ReadAuth), null);
                    }
                    else
                    {
                        JavaScriptSerializer json = new JavaScriptSerializer(); // feel free to use JSON serializer
                        Dictionary<string, string> data = json.Deserialize<Dictionary<string, string>>(obj.data.ToString());
                        if (data.ContainsKey("status") && data["status"].Equals("authorized"))
                        {
                            Connected(true);
                        }
                        obj.data.Clear();
                        obj.handle.Set();
                    }
                }
                catch (Exception ex)
                {
                    obj.data.Clear();
                    Log(ErrorMsg(ex.Message));
                    obj.handle.Set();
                }
            }
            else
            {
                obj.client.Close();
                obj.handle.Set();
            }
        }

        private bool Authorize()
        {
            bool success = false;
            Dictionary<string, string> data = new Dictionary<string, string>();
            data.Add("username", obj.username);
            data.Add("key", obj.key);
            JavaScriptSerializer json = new JavaScriptSerializer(); // feel free to use JSON serializer
            Send(json.Serialize(data));
            while (obj.client.Connected)
            {
                try
                {
                    obj.stream.BeginRead(obj.buffer, 0, obj.buffer.Length, new AsyncCallback(ReadAuth), null);
                    obj.handle.WaitOne();
                    if (connected)
                    {
                        success = true;
                        break;
                    }
                }
                catch (Exception ex)
                {
                    Log(ErrorMsg(ex.Message));
                }
            }
            if (!connected)
            {
                Log(SystemMsg("Unauthorized"));
            }
            return success;
        }

        private void Connection(IPAddress ip, int port, string username, string key)
        {
            try
            {
                obj = new MyClient();
                obj.username = username;
                obj.key = key;
                obj.client = new TcpClient();
                obj.client.Connect(ip, port);
                obj.stream = obj.client.GetStream();
                obj.buffer = new byte[obj.client.ReceiveBufferSize];
                obj.data = new StringBuilder();
                obj.handle = new EventWaitHandle(false, EventResetMode.AutoReset);
                if (Authorize())
                {
                    while (obj.client.Connected)
                    {
                        try
                        {
                            obj.stream.BeginRead(obj.buffer, 0, obj.buffer.Length, new AsyncCallback(Read), null);
                            obj.handle.WaitOne();
                        }
                        catch (Exception ex)
                        {
                            Log(ErrorMsg(ex.Message));
                        }
                    }
                    obj.client.Close();
                    Connected(false);
                }
            }
            catch (Exception ex)
            {
                Log(ErrorMsg(ex.Message));
            }
        }

        private void ConnectButton_Click(object sender, EventArgs e)
        {
            if (connected)
            {
                obj.client.Close();
            }
            else if (client == null || !client.IsAlive)
            {
                string address = addrTextBox.Text.Trim();
                string number = portTextBox.Text.Trim();
                //string username = usernameTextBox.Text.Trim();
                bool error = false;
                IPAddress ip = null;
                if (address.Length < 1)
                {
                    error = true;
                    Log(SystemMsg("Address is required"));
                }
                else
                {
                    try
                    {
                        ip = Dns.Resolve(address).AddressList[0];
                    }
                    catch
                    {
                        error = true;
                        Log(SystemMsg("Address is not valid"));
                    }
                }
                int port = -1;
                if (number.Length < 1)
                {
                    error = true;
                    Log(SystemMsg("Port number is required"));
                }
                else if (!int.TryParse(number, out port))
                {
                    error = true;
                    Log(SystemMsg("Port number is not valid"));
                }
                else if (port < 0 || port > 65535)
                {
                    error = true;
                    Log(SystemMsg("Port number is out of range"));
                }
                //if (username.Length < 1)
                //{
                //    error = true;
                //    Log(SystemMsg("Username is required"));
                //}
                if (!error)
                {
                    // encryption key is optional
                    client = new Thread(() => Connection(ip, port, "user", "key"))
                    {
                        IsBackground = true
                    };
                    client.Start();
                }
            }
        }


        private void Connect()
        {
            if (!connected)
            {


             if (client == null || !client.IsAlive)
                {
                    string address = addrTextBox.Text.Trim();
                    string number = portTextBox.Text.Trim();
                   // string username = usernameTextBox.Text.Trim();
                    bool error = false;
                    IPAddress ip = null;
                    if (address.Length < 1)
                    {
                        error = true;
                        Log(SystemMsg("Address is required"));
                    }
                    else
                    {
                        try
                        {
                            ip = Dns.Resolve(address).AddressList[0];
                        }
                        catch
                        {
                            error = true;
                            Log(SystemMsg("Address is not valid"));
                        }
                    }
                    int port = -1;
                    if (number.Length < 1)
                    {
                        error = true;
                        Log(SystemMsg("Port number is required"));
                    }
                    else if (!int.TryParse(number, out port))
                    {
                        error = true;
                        Log(SystemMsg("Port number is not valid"));
                    }
                    else if (port < 0 || port > 65535)
                    {
                        error = true;
                        Log(SystemMsg("Port number is out of range"));
                    }
                    //if (username.Length < 1)
                    //{
                    //    error = true;
                    //    Log(SystemMsg("Username is required"));
                    //}
                    if (!error)
                    {
                        // encryption key is optional
                        client = new Thread(() => Connection(ip, port, "user", "key"))
                        {
                            IsBackground = true
                        };
                        client.Start();
                    }
                }
            }
        }

        private void Write(IAsyncResult result)
        {
            if (obj.client.Connected)
            {
                try
                {
                    obj.stream.EndWrite(result);
                }
                catch (Exception ex)
                {
                    Log(ErrorMsg(ex.Message));
                }
            }
        }

        private void BeginWrite(string msg)
        {
            byte[] buffer = Encoding.UTF8.GetBytes(msg);
            if (obj.client.Connected)
            {
                try
                {
                    obj.stream.BeginWrite(buffer, 0, buffer.Length, new AsyncCallback(Write), null);
                }
                catch (Exception ex)
                {
                    Log(ErrorMsg(ex.Message));
                }
            }
        }

        private void Send(string msg)
        {
            if (send == null || send.IsCompleted)
            {
                send = Task.Factory.StartNew(() => BeginWrite(msg));
            }
            else
            {
                send.ContinueWith(antecendent => BeginWrite(msg));
            }
        }

        //private void SendTextBox_KeyDown(object sender, KeyEventArgs e)
        //{
        //    if (e.KeyCode == Keys.Enter)
        //    {
        //        e.Handled = true;
        //        e.SuppressKeyPress = true;
        //        if (sendTextBox.Text.Length > 0)
        //        {
        //            string msg = sendTextBox.Text;
        //            sendTextBox.Clear();
        //            Log(string.Format("{0} (You): {1}", obj.username, msg));
        //            if (connected)
        //            {
        //                Send(msg);
        //            }
        //        }
        //    }
        //}

        private void Client_FormClosing(object sender, FormClosingEventArgs e)
        {
            exit = true;
            if (connected)
            {
                obj.client.Close();
            }
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            Log();
        }

        private void CheckBox_CheckedChanged(object sender, EventArgs e)
        {
            //if (keyTextBox.PasswordChar == '*')
            //{
            //    keyTextBox.PasswordChar = '\0';
            //}
            //else
            //{
            //    keyTextBox.PasswordChar = '*';
            //}
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Connect();
        }

        private void Client_Load(object sender, EventArgs e)
        {
            timer1.Start();
        }

        public static List<IPAddress> GetDefaultDns()
        {
            //var card = NetworkInterface.GetAllNetworkInterfaces().FirstOrDefault();
            var cards = NetworkInterface.GetAllNetworkInterfaces();
            if (cards == null) return null;
            List<IPAddress> address = new List<IPAddress>();
            foreach(var card in cards)
            {
                address.AddRange(card.GetIPProperties().DnsAddresses.ToList());
            }
            return address;
        }

        void AllowSite(string sites="")
        {
            AllowInternet();
            if (sites == "All")
            {
                return;
            }

  
            List<IPAddress> addresslist = new List<IPAddress>();
            if (sites != "")
            {
                var arSites = sites.Split(';');
                foreach (var site in arSites)
                {
                    addresslist.AddRange(Dns.GetHostAddresses(site).ToList());
                }
            }
            addresslist.Add(Dns.Resolve(addrTextBox.Text.Trim()).AddressList[0]);
            addresslist.AddRange(GetDefaultDns());


            // BanInternet();
            INetFwRule firewallRule = (INetFwRule)Activator.CreateInstance(
            Type.GetTypeFromProgID("HNetCfg.FWRule"));

            string RemoteAddresses = "";
            List<string> ipList = new List<string>();
            foreach (IPAddress theaddress in addresslist)
            {
                if (ipList.Contains(theaddress.ToString()))
                {
                    for(int j = 0; j < ipList.Count; j++)
                    {
                        if (ipList[j].Equals(theaddress.ToString()))
                        {
                            if (j == 0)
                            {
                                ipList[0] = ToAddr(ToInt(ipList[0])-1).ToString();
                                break;
                            }

                            if (j == ipList.Count-1)
                            {
                                ipList[j] = ToAddr(ToInt(ipList[j]) + 1).ToString();
                                break;
                            }
                            var it = ToInt(ipList[j]);
                            var prev= ToInt(ipList[j-1]);
                            var next = ToInt(ipList[j + 1]);
                            if (it - prev < next - it)
                            {
                                ipList[j] = ToAddr(ToInt(ipList[j]) + 1).ToString();
                            }
                            else
                            {
                                ipList[j] = ToAddr(ToInt(ipList[j]) - 1).ToString();
                            }
                        }
                    }
                    continue;
                }
                //var a = theaddress.Address;
                if (!theaddress.IsIPv6SiteLocal)
                {
                    byte[] ip = theaddress.GetAddressBytes();
                    ip[3] = (byte)(ip[3] + 1);
                    if (ip[3] == 0)
                    {
                        ip[2] = (byte)(ip[2] + 1);
                        if (ip[2] == 0)
                        {
                            ip[1] = (byte)(ip[1] + 1);
                            if (ip[1] == 0)
                            {
                                ip[0] = (byte)(ip[0] + 1);
                            }
                        }
                    }

                    byte[] ipMinus = theaddress.GetAddressBytes();
                    if (ipMinus[3] == 0)
                    {
                        ipMinus[3] = 255;
                        if (ipMinus[2] == 0)
                        {
                            ipMinus[2] = 255;

                            if (ipMinus[1] == 0)
                            {
                                ipMinus[1] = 255;
                                ipMinus[0] = (byte)(ipMinus[0] - 1);
                            }
                            else
                            {
                                ipMinus[1] = (byte)(ipMinus[1] - 1);
                            }

                        }
                        else
                        {
                            ipMinus[2] = (byte)(ipMinus[2] - 1);
                        }
                    }
                    else
                    {
                        ipMinus[3] = (byte)(ipMinus[3] - 1);
                    }
                    ipList.Add(new IPAddress(ipMinus).ToString());
                    ipList.Add(new IPAddress(ip).ToString());
                    ipList = ipList
                .Select(Version.Parse)
                .OrderBy(arg => arg)
                .Select(arg => arg.ToString())
                .ToList();
                    RemoteAddresses += theaddress.ToString() + ",";
                }
               

            }
            var sortedIps = ipList
            .Select(Version.Parse)
            .OrderBy(arg => arg)
            .Select(arg => arg.ToString())
            .ToList();

            string ra = "0.0.0.0-";
            ra += sortedIps.First();
            ra += ",";
            int i = 1;
            while (i < sortedIps.Count)
            {
                ra += sortedIps[i];
                ra += "-";
                i++;
                if (i < sortedIps.Count)
                {
                    ra += sortedIps[i];
                    ra += ",";
                    i++;
                }
                else
                {
                    ra += "255.255.255.255";
                    break;
                }
            }
            firewallRule.RemoteAddresses = ra;
            firewallRule.Action = NET_FW_ACTION_.NET_FW_ACTION_BLOCK;
            firewallRule.Direction = NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_OUT;
            firewallRule.Enabled = true;
            firewallRule.InterfaceTypes = "All";
            firewallRule.Name = "Ban internet allow sites";

            INetFwPolicy2 firewallPolicy = (INetFwPolicy2)Activator.CreateInstance(
                Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
            firewallPolicy.Rules.Add(firewallRule);
        }

        static long ToInt(string addr)
        {
            // careful of sign extension: convert to uint first;
            // unsigned NetworkToHostOrder ought to be provided.
            return (long)(uint)IPAddress.NetworkToHostOrder(
                 (int)IPAddress.Parse(addr).Address);
        }

        static string ToAddr(long address)
        {
            return IPAddress.Parse(address.ToString()).ToString();
            // This also works:
            // return new IPAddress((uint) IPAddress.HostToNetworkOrder(
            //    (int) address)).ToString();
        }

        void AllowInternet()
        {
            INetFwPolicy2 firewallPolicy = (INetFwPolicy2)Activator.CreateInstance(
            Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
            firewallPolicy.Rules.Remove("Block Internet");
            firewallPolicy.Rules.Remove("Ban internet allow sites");

        }
    }
}
