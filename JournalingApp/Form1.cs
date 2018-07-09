using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JournalingApp
{
    public partial class Form1 : Form
    {

        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        public extern static IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", EntryPoint = "GetForegroundWindow", CharSet = System.Runtime.InteropServices.CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetF(); //获得本窗体的句柄

        [DllImport("user32.dll", EntryPoint = "SetForegroundWindow")]
        public static extern bool SetF(IntPtr hWnd); //设置此窗体为活动窗体

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int y, int Width, int Height, int flags);

        [DllImport("user32")]
        private static extern bool AnimateWindow(IntPtr hwnd, int dwTime, int dwFlags);
        //下面是可用的常量，根据不同的动画效果声明自己需要的
        private const int AW_HOR_POSITIVE = 0x0001;//自左向右显示窗口，该标志可以在滚动动画和滑动动画中使用。使用AW_CENTER标志时忽略该标志
        private const int AW_HOR_NEGATIVE = 0x0002;//自右向左显示窗口，该标志可以在滚动动画和滑动动画中使用。使用AW_CENTER标志时忽略该标志
        private const int AW_VER_POSITIVE = 0x0004;//自顶向下显示窗口，该标志可以在滚动动画和滑动动画中使用。使用AW_CENTER标志时忽略该标志
        private const int AW_VER_NEGATIVE = 0x0008;//自下向上显示窗口，该标志可以在滚动动画和滑动动画中使用。使用AW_CENTER标志时忽略该标志该标志
        private const int AW_CENTER = 0x0010;//若使用了AW_HIDE标志，则使窗口向内重叠；否则向外扩展
        private const int AW_HIDE = 0x10000;//隐藏窗口
        private const int AW_ACTIVE = 0x20000;//激活窗口，在使用了AW_HIDE标志后不要使用这个标志
        private const int AW_SLIDE = 0x40000;//使用滑动类型动画效果，默认为滚动动画类型，当使用AW_CENTER标志时，这个标志就被忽略
        private const int AW_BLEND = 0x80000;//使用淡入淡出效果


        public Form1()
        {
            InitializeComponent();
            Dictionary<string, string> dict = this.ReadLineFile();
            this.textBox1.Text = dict["uid"];
            this.textBox2.Text = dict["pwd"];
            this.checkBox1.Checked = "true".Equals(dict["qkjcyl"]) ? true : false;
            this.checkBox2.Checked = "true".Equals(dict["outwork"]) ? true : false;
            this.textBox3.Text = dict["worktime"];
            this.richTextBox1.Text = dict["worktext"];
            DataTable tblDatas = new DataTable("Datas");
            tblDatas.Columns.Add("id", Type.GetType("System.String"));
            tblDatas.Columns.Add("name", Type.GetType("System.String"));

            tblDatas.Rows.Add(new object[] { dict["projectid"], dict["projectname"] });
            
            this.comboBox1.DisplayMember = "name";
            this.comboBox1.ValueMember = "id";
            this.comboBox1.DataSource = tblDatas;

            this.xmbm = dict["xmbm"];
            this.xmlb = dict["xmlb"];
            this.zcdh = dict["zcdh"];

            

            this.timer = new System.Timers.Timer();
            this.timer.Interval = 5000D;
            this.timer.Elapsed += new System.Timers.ElapsedEventHandler(this.submitJournalingByTime);
            this.timer.Start();

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                this.notifyIcon1.Visible = true;
            }
        }
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void notifyIcon1_Click(object sender, EventArgs e)
        {
            this.Visible = true;
            this.WindowState = FormWindowState.Normal;
            this.notifyIcon1.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
            if (LoginSimulation())
            {
                if (!"".Equals(this.mycookie))
                {
                    if (//true)
                        journalingSubmit())
                    {
                        this.isSubmit = true;
                        this.label7.Text = "已于"+ DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 手动提交日志";
                        writeLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 手动提交日志成功");
                    }
                    else
                    {
                        writeLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 手动提交日志失败");
                    }
                }
            }
            else {
                writeLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 登陆失败");
            }
            
           
        }

        private void submitJournalingByTime(object sender, EventArgs e)
        {
           
            if (DateTime.Now.ToString("HH:mm").Equals("00:00")) {
                this.isSubmit = false;
                this.submitTime = "";
                this.time = "";
            }
            if (!this.isSubmit) {
                if (//true)
                !"Saturday".Equals(DateTime.Now.DayOfWeek.ToString()) && !"Sunday".Equals(DateTime.Now.DayOfWeek.ToString()))
                {
                    if (this.submitTime == "")
                    {
                        Random random = new Random((int)(DateTime.Now.Ticks));
                        int hour = random.Next(16,18);
                        int minute = 0;
                        if (hour == 16)
                        {
                            minute = random.Next(0, 60);
                        }
                        else if (hour == 17)
                        {
                            minute = random.Next(0, 30);
                        }
                        int second = 0;
                        string tempStr = string.Format("{0}:{1}:{2}", hour, minute, second);
                        DateTime rTime = Convert.ToDateTime(tempStr);
                        this.label7.Text = "将在" + rTime.ToString("yyyy-MM-dd HH:mm:ss") + "自动提交日志";
                        this.submitTime = rTime.ToString("yyyy-MM-dd HH:mm:ss");
                        this.time = rTime.ToString("HH:mm");
                    }
                    else
                    {

                        if (this.time.Equals(DateTime.Now.ToString("HH:mm")))
                        {
                            //提交日志
                            this.timer.Stop();
                            this.setFocus();
                            try
                            {
                                if (LoginSimulation())
                                {
                                    if (!"".Equals(this.mycookie))
                                    {
                                        if (//true)
                                            journalingSubmit())
                                        {
                                            this.isSubmit = true;
                                            this.label7.Text = "已于" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 自动提交日志";
                                            writeLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 自动提交日志成功");
                                        }
                                        else
                                        {
                                            writeLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 自动提交日志失败");
                                        }
                                    }
                                }
                                else
                                {
                                    writeLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 登陆失败");
                                }
                            }
                            catch (Exception ex)
                            {
                                writeLog(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " 自动提交日志失败："+ex.Message);
                            }
                            finally
                            {
                                this.timer.Start();
                            }
                        }
                    }

                }
            else
            {
                    this.label7.Text = "今天是周末哦，您可以手动提交日志";
                }
            }
            
        }

        private bool LoginSimulation()
        {

            string url = "http://pro.bsoft.com.cn/platform/logon/myRoles";
            string postData = "{\"pwd\":\""+this.textBox2.Text+ "\",\"uid\":\""+this.textBox1.Text+"\",\"url\":\"logon/myRoles\"}";

            ////1.获取登录Cookie
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "POST";// POST OR GET， 如果是GET, 则没有第二步传参，直接第三步，获取服务端返回的数据
            req.AllowAutoRedirect = false;//服务端重定向。一般设置false
            req.ContentType = "application/x-www-form-urlencoded";//数据一般设置这个值，除非是文件上传

            byte[] postBytes = Encoding.UTF8.GetBytes(postData);
            req.ContentLength = postBytes.Length;
            Stream postDataStream = req.GetRequestStream();
            postDataStream.Write(postBytes, 0, postBytes.Length);
            postDataStream.Close();
            
            HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
            string cookies = resp.Headers.Get("Set-Cookie");//获取登录后的cookie值。
            string cookie = cookies.Split(';')[0];
            this.label7.Text = "获取cookie:"+cookie;
            string html = new StreamReader(resp.GetResponseStream()).ReadToEnd();
            JObject o = (JObject)Newtonsoft.Json.JsonConvert.DeserializeObject(html);
            string c = o["body"]["tokens"][0]["id"].ToString();
            string contentUrl2 = "http://pro.bsoft.com.cn/platform/logon/myApps?urt="+ c + "&deep=3";
            HttpWebRequest reqContent2 = (HttpWebRequest)WebRequest.Create(contentUrl2);
            //reqContent2.Method = "GET";
            reqContent2.MediaType = "GET";
            reqContent2.AllowAutoRedirect = false;//服务端重定向。一般设置false
            reqContent2.ContentType = "application/json";//数据一般设置这个值，除非是文件上传
            reqContent2.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
            reqContent2.Host = "pro.bsoft.com.cn";
            reqContent2.KeepAlive = true;
            reqContent2.Headers.Add("Accept-Encoding", "gzip, deflate");
            reqContent2.Headers.Add("Accept-Language", "zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3");
            reqContent2.Headers.Add("encoding", "utf-8");
            reqContent2.Referer = "http://pro.bsoft.com.cn/platform/index.html";
            reqContent2.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64; rv:45.0) Gecko/20100101 Firefox/45.0";
            reqContent2.Credentials = CredentialCache.DefaultCredentials;
            reqContent2.CookieContainer = new CookieContainer();
            reqContent2.CookieContainer.SetCookies(reqContent2.RequestUri, cookie);//将登录的cookie值赋予此次的请求。

            HttpWebResponse respContent2 = (HttpWebResponse)reqContent2.GetResponse();
            string html2 = new StreamReader(respContent2.GetResponseStream()).ReadToEnd();
            if (html2 != null)
            {
                this.label7.Text = "登陆成功";
                this.mycookie = cookie;
                return true;
            }
            else {
                return false;
            }

           
          
        }

        private bool journalingSubmit()
        {
            string contentUrl = "http://pro.bsoft.com.cn/platform/*.jsonRequest";
            HttpWebRequest reqContent = (HttpWebRequest)WebRequest.Create(contentUrl);
            reqContent.Method = "POST";
            reqContent.AllowAutoRedirect = false;//服务端重定向。一般设置false
            reqContent.ContentType = "application/json";//数据一般设置这个值，除非是文件上传
            string postData1 = "{\"serviceId\":\"SupportWorkLogService\",\"method\":\"execute\",\"body\":{\"gsxm\":\"" + this.comboBox1.SelectedValue + "\",\"gzrz\":\"" + this.richTextBox1.Text + "\",\"rzqk\":\"1\",\"zfid\":\"\",\"id\":\"\",\"kqid\":\"\",\"xmlb\":\"" + this.xmlb + "\",\"xmbm\":\"" + this.xmbm + "\",\"ccbz\":0,\"blbz\":0,\"xmmc\":\"" + this.comboBox1.SelectedText + "\",\"projectid\":\"" + this.comboBox1.SelectedValue + "\",\"zcgs\":" + this.textBox3.Text + ",\"zcgsmx\":\"" + this.textBox3.Text + ",\"},\"TaskLog\":{\"zcdh\":\"" + this.zcdh + "\",\"rzid\":\"\",\"zcry\":\"" + this.textBox1.Text + "\",\"cpmk\":\"" + (this.checkBox1.Checked ? this.cpmk : "") + "\",\"mkmc\":\"" + (this.checkBox1.Checked ? this.mkmc : "") + "\",\"nrid\":\"2139\",\"blbz\":0,\"zcgs\":" + this.textBox3.Text + ",\"zcgsmx\":\"" + this.textBox3.Text + ",\"}}";
            byte[] postBytes1 = Encoding.UTF8.GetBytes(postData1);
            reqContent.CookieContainer = new CookieContainer();
            reqContent.CookieContainer.SetCookies(reqContent.RequestUri, this.mycookie);//将登录的cookie值赋予此次的请求。
            reqContent.ContentLength = postBytes1.Length;
            Stream postDataStream1 = reqContent.GetRequestStream();
            postDataStream1.Write(postBytes1, 0, postBytes1.Length); postDataStream1.Close();
            HttpWebResponse resp1 = (HttpWebResponse)reqContent.GetResponse();

            string cookies1 = resp1.Headers.Get("Set-Cookie");
            JObject o = (JObject)Newtonsoft.Json.JsonConvert.DeserializeObject(new StreamReader(resp1.GetResponseStream()).ReadToEnd());
            if (o["code"].ToString().Equals("200"))
            {
                return true;
            }
            else {
                return false;
            }
           
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

       

        private void Form1_FormClosing(object sender, EventArgs e)
        {
            AnimateWindow(this.Handle, 1000, AW_BLEND | AW_HIDE);
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            int x = Screen.PrimaryScreen.WorkingArea.Right - this.Width;
            int y = Screen.PrimaryScreen.WorkingArea.Bottom - this.Height;
            this.Location = new Point(x, y);//设置窗体在屏幕右下角显示
            AnimateWindow(this.Handle, 1000, AW_SLIDE | AW_ACTIVE | AW_VER_NEGATIVE);
            this.richTextBox1.Focus();
            // 选中文本文本框中的关键字
            this.richTextBox1.Select(0, this.richTextBox1.Text.Length);
        }

        public  Dictionary<string, string> ReadLineFile()
        {
            string filePath = System.Windows.Forms.Application.StartupPath + "\\config.txt";

            Dictionary<string, string> contentDictionary = new Dictionary<string, string>();

            if (!File.Exists(filePath))
            {
                return contentDictionary;
            }

            FileStream fileStream = null;

            StreamReader streamReader = null;

            try
            {
                fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);

                streamReader = new StreamReader(fileStream, Encoding.Default);

                fileStream.Seek(0, SeekOrigin.Begin);

                string content = streamReader.ReadLine();

                while (content != null)
                {
                    if (content.Contains("="))
                    {
                        string key = content.Substring(0, content.LastIndexOf("=")).Trim();

                        string value = content.Substring(content.LastIndexOf("=") + 1).Trim();

                        if (!contentDictionary.ContainsKey(key))
                        {
                            contentDictionary.Add(key, value);
                        }
                    }
                    content = streamReader.ReadLine();
                }
            }
            catch
            {
            }
            finally
            {
                if (fileStream != null)
                {
                    fileStream.Close();
                }
                if (streamReader != null)
                {
                    streamReader.Close();
                }
            }
            return contentDictionary;
        }

        private void writeLog(string str) {

            using (System.IO.StreamWriter file = new System.IO.StreamWriter(System.Windows.Forms.Application.StartupPath + "\\log.txt", true))
            {
                file.WriteLine("工号："+ this.textBox1.Text + ":" + str);// 直接追加文件末尾，换行
                file.Flush();
                file.Close();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void setFocus()
        {

            SetWindowPos(this.Handle, -1, 0, 0, 0, 0, 1 | 2);
            ShowWindow(this.Handle, 9);
            SetF(this.Handle);

        }


    }
}
