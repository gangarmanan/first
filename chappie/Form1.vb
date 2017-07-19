using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//add namespaces
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.IO;
using System.Xml;
using System.Globalization;
using System.Reflection;
using System.Threading;

Namespace JARVIS
{


    public partial class Form1 : Form
    {
        Random rnd = new Random();
        int count = 1;
        int i;
        double timer = 10;
        string Temperature;
        string Condition;
        string Humidity;
        string WindSpeed;
        string Town;
        string user;

        static CultureInfo ci = new CultureInfo("en-US");
        SpeechRecognitionEngine _recognizer = new SpeechRecognitionEngine();
        SpeechSynthesizer JARVIS = new SpeechSynthesizer();
        string QEvent, ProcWindow, BrowseDirectory;

        Public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            user = System.Windows.Forms.SystemInformation.UserName;

            string[] cmds = (File.ReadAllLines(@"Commands.txt"));

            Choices commands = new Choices();
            commands.Add(cmds);
            GrammarBuilder gBuilder = new GrammarBuilder();
            gBuilder.Append(commands);
            Grammar grammar = new Grammar(gBuilder);
            _recognizer.LoadGrammarAsync(grammar);
            _recognizer.SetInputToDefaultAudioDevice();
            _recognizer.RecognizeAsync(RecognizeMode.Multiple);
            _recognizer.SpeechRecognized += _recognizer_SpeechRecognized;
        }


        void _recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string speech = e.Result.Text;

            switch (speech)
            {
                //greetings
                case "hello":
                case "hello jarvis":
                    DateTime now = DateTime.Now;
                    if (now.Hour >= 5 && now.Hour < 12)
                    { 
                        JARVIS.SpeakAsync("Goodmorning " + user); 
                    }
                    if (now.Hour >= 12 && now.Hour < 18)
                    {
                        JARVIS.SpeakAsync("Good afternoon " + user);
                    }
                    if (now.Hour >= 18 && now.Hour < 20)
                    {
                        JARVIS.SpeakAsync("Good evening " + user); 
                    }
                    if (now.Hour >= 20 && now.Hour < 24)
                    {
                        JARVIS.SpeakAsync("Good night " + user);
                    }
                    if (now.Hour < 5)
                    { 
                        JARVIS.SpeakAsync("Hello " + user + ", it's getting late");
                    }
                    break;

                case "goodbye":
                case "goodbye jarvis":
                case "close":
                case "close jarvis":
                    JARVIS.Speak("OK Nishant Sir Until next time");
                    Close();
                    break;

                case "good":
                    JARVIS.Speak("Thank you sir");
                    break;

                case "jarvis":
                        QEvent = "";
                        JARVIS.Speak("Yes Sir");
                    break;

                case "what is my name":
                    JARVIS.SpeakAsync(user);
                    break;


                //web commands
                case "open facebook":
                    System.Diagnostics.Process.Start("http://www.facebook.com");
                    break;

                case "open google":
                    System.Diagnostics.Process.Start("http://www.google.com");
                    break;

                //shell commands
                case "open command":
                    System.Diagnostics.Process.Start("cmd");
                    break;

                case "open visual studio":
                    System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\IDE\devenv.exe");
                    JARVIS.Speak("Loading");
                    break;

                case "open sql":
                    System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Microsoft SQL Server\110\Tools\Binn\ManagementStudio\Ssms.exe");
                    JARVIS.Speak("Loading");
                    break;

                case "open notepad":
                    System.Diagnostics.Process.Start("notepad");
                    JARVIS.Speak("Loading");
                    break;

                case "open word pad":
                    System.Diagnostics.Process.Start("wordpad");
                    JARVIS.Speak("Loading");
                    break;

                case "open msword":
                    System.Diagnostics.Process.Start("winword");
                    JARVIS.Speak("Loading");
                    break;

                case "open mspowerpoint":
                    System.Diagnostics.Process.Start("powerpnt");
                    JARVIS.Speak("Loading");
                    break;


                case "open excel":
                    System.Diagnostics.Process.Start("excel");
                    JARVIS.Speak("Loading");
                    break;

                case "open access":
                    System.Diagnostics.Process.Start("msaccess");
                    JARVIS.Speak("Loading");
                    break;

                case "open outlook":
                    System.Diagnostics.Process.Start("outlook");
                    JARVIS.Speak("Loading");
                    break;

                case "open task manager":
                    System.Diagnostics.Process.Start("taskmgr");
                    JARVIS.Speak("Loading");
                    break;

                case "open vlc player":
                    System.Diagnostics.Process.Start(@"C:\Program Files (x86)\VideoLAN\VLC\vlc.exe");
                    JARVIS.Speak("Loading");
                    break;

                //close programs
                case "close outlook":
                    ProcWindow = "outlook";
                    StopWindow();
                    break;

                case "close excel":
                    ProcWindow = "excel";
                    StopWindow();
                    break;

                case "close access":
                    ProcWindow = "msaccess";
                    StopWindow();
                    break;

                case "close visual studio":
                    ProcWindow = "devenv";
                    StopWindow();
                    break;

                case "close mspowerpoint":
                    ProcWindow = "powerpnt";
                    StopWindow();
                    break;

                case "close msword":
                    ProcWindow = "winword";
                    StopWindow();
                    break;

                case "close notepad":
                    ProcWindow = "notepad";
                    StopWindow();
                    break;

                case "close word pad":
                    ProcWindow = "wordpad";
                    StopWindow();
                    break;

                case "close command":
                    ProcWindow = "cmd";
                    StopWindow();
                    break;

                case "close task manager":
                    ProcWindow = "taskmgr";
                    StopWindow();
                    break;

                case "close sql":
                    ProcWindow = "ssms";
                    StopWindow();
                    break;

                case "close browser":
                    ProcWindow = "chrome";
                    StopWindow();
                    break;

                case "close vlc":
                    ProcWindow = "vlc";
                    StopWindow();
                    break;

                //condition of day
                case "what time is it":
                    now = DateTime.Now;
                    string time = now.GetDateTimeFormats('t')[0];
                    JARVIS.Speak(time);
                    break;

                case "what day is it":
                    JARVIS.Speak(DateTime.Today.ToString("dddd"));
                    break;

                case "whats the date":
                case "whats todays date":
                    JARVIS.Speak(DateTime.Today.ToString("dd-mm-yyyy"));
                    break;

                case "what weather is it":
                    try
                    {
                        GetWeather();
                        float temp = (float.Parse(Temperature) - 32) * 5 / 9;

                        JARVIS.Speak("Temperature is " + temp.ToString() + "degree celcius " + " and Condition is " + Condition + " and humidity is " + Humidity
                            + " and windspeed is " + WindSpeed + " in location " + Town);
                    }
                    catch
                    {
                        JARVIS.Speak("Unable to update weather....!" + " Check internet connection sir...!");
                    }
                    break;

                //other commands
                case "go fullscreen":
                    FormBorderStyle = FormBorderStyle.None;
                    WindowState = FormWindowState.Maximized;
                    TopMost = true;
                    JARVIS.Speak("Expending");
                    break;

                case "exit fullscreen":
                    FormBorderStyle = FormBorderStyle.Sizable;
                    WindowState = FormWindowState.Normal;
                    TopMost = false;
                    break;

                case "switch window":
                    SendKeys.Send("%{TAB " + count + "}");
                    count += 1;
                    JARVIS.Speak("sure sir...!");
                    break;

                case "reset":
                    timer = 10;
                    lblTimer.Visible = false;
                    ShutdownTimer.Enabled = false;
                    lstCommands.Visible = false;
                    JARVIS.Speak("Ok Sir...");
                    break;

                case "out of the way":
                    if (WindowState == FormWindowState.Normal || WindowState == FormWindowState.Maximized) 
                    {
                        WindowState = FormWindowState.Minimized;
                        JARVIS.Speak("My apologies");
                    }
                    break;

                case "come back":
                    if (WindowState == FormWindowState.Minimized) 
                    {
                        JARVIS.Speak("Alright?");
                        WindowState = FormWindowState.Normal;
                    }
                    break;

                case "show commands":
                    string[] commands = (File.ReadAllLines(@"Commands.txt"));
                    JARVIS.Speak("very well");
                    lstCommands.Items.Clear();
                    lstCommands.SelectionMode = SelectionMode.None;
                    lstCommands.Visible = true;

                    foreach(string cmd in commands)
                    {
                        lstCommands.Items.Add(cmd);
                    }
                    break;

                case "hide commands":
                    lstCommands.Visible = false;
                    break;


                    //shutdown restart log off
                case "abort":
                    if (ShutdownTimer.Enabled == true)
                    {
                        QEvent = "abort";
                        JARVIS.Speak("Ok Sir...");
                    }
                    break;

                case "Shutdown":
                    if (ShutdownTimer.Enabled == false)
                    {
                        QEvent = "shutdown";
                        JARVIS.Speak("Will see you later");
                        lblTimer.Visible = true;
                        ShutdownTimer.Enabled = true;
                    }
                    break;

                case "restart":
                    if (ShutdownTimer.Enabled == false)
                    {
                        QEvent = "restart";
                        JARVIS.Speak("I will be back shortly");
                        lblTimer.Visible = true;
                        ShutdownTimer.Enabled = true;
                    }
                    break;

                case "speed up":
                    if (ShutdownTimer.Enabled == true)
                    {
                        ShutdownTimer.Interval = ShutdownTimer.Interval / 10;
                        JARVIS.Speak("process speed up....!");
                    }
                    break;

                case "slow down":
                    if (ShutdownTimer.Enabled == true)
                    {
                        ShutdownTimer.Interval = ShutdownTimer.Interval * 10;
                        JARVIS.Speak("process slow down.....!");
                    }
                    break;

                //Loading Directories
                case
                "load music":
                    QEvent = "Music";
                    BrowseDirectory = @"E:\Multimedia\Songs\";
                    LoadDirectory();
                    lstCommands.Visible = true;
                    JARVIS.Speak("Loading");
                    break;

                case "load videos":
                    QEvent = "Videos";
                    BrowseDirectory = @"E:\Multimedia\Videos\";
                    LoadDirectory();
                    lstCommands.Visible = true;
                    JARVIS.Speak("Loading");
                    break;

                case "load pics":
                    QEvent = "Pics";
                    BrowseDirectory = @"E:\Multimedia\Images\";
                    LoadDirectory();
                    lstCommands.Visible = true;
                    JARVIS.Speak("Loading");
                    break;

                case "load directory":
                    FolderBrowserDialog fbd = new FolderBrowserDialog(); 
                    if(fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        BrowseDirectory = fbd.SelectedPath;
                        QEvent = "Load Folder";
                        LoadDirectory();
                        lstCommands.Visible = true;
                        JARVIS.Speak("Loading");
                    }
                    break;
            }
        }

        private void ShutdownTimer_Tick(object sender, EventArgs e)
        {
            if(timer == 0)
            {
                lblTimer.Visible = false;
                ComputerTermindation();
                ShutdownTimer.Enabled = false;
            }
            else if(QEvent == "abort")
            {
                timer = 10;
                lblTimer.Visible = false;
                ShutdownTimer.Enabled = false;
            }
            else
            {
                timer = timer - .01;
                lblTimer.Text = timer.ToString();
            }
        }

        private void ComputerTermindation()
        {
            switch (QEvent)
            {
                case "shutdown":
                    System.Diagnostics.Process.Start("shutdown", "-s");
                    break;

                case "restart":
                    System.Diagnostics.Process.Start("shutdown", "-r");
                    break;

                case "log off":
                    System.Diagnostics.Process.Start("shutdown", "-l");
                    break;
            }
        }


        private void StopWindow()
        {
            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName(ProcWindow);

            foreach(System.Diagnostics.Process proc in procs)
            {
                proc.CloseMainWindow();
            }
        }

        private void LoadDirectory()
        {
            string[] Files;

            lstCommands.Items.Clear();
            switch (QEvent)
            {
                case "Music":
                    Files = Directory.GetFiles(BrowseDirectory, "*", SearchOption.AllDirectories);
                    foreach (string file in Files)
                    {
                        lstCommands.Items.Add(file.Replace(BrowseDirectory, ""));
                    }
                    break;

                case "Videos":
                    Files = Directory.GetFiles(BrowseDirectory, "*", SearchOption.AllDirectories);
                    foreach (string file in Files)
                    {
                        lstCommands.Items.Add(file.Replace(BrowseDirectory, ""));
                    }
                    break;

                case "Pics":
                    Files = Directory.GetFiles(BrowseDirectory, "*", SearchOption.AllDirectories);
                    foreach (string file in Files)
                    {
                        lstCommands.Items.Add(file.Replace(BrowseDirectory, ""));
                    }
                    break;

                case "Load Folder":
                    Files = Directory.GetFiles(BrowseDirectory, "*", SearchOption.AllDirectories);
                    foreach (string file in Files)
                    {
                        lstCommands.Items.Add(file.Replace(BrowseDirectory, ""));
                    }
                    break;
            }
        }
        private void lstCommands_SelectedIndexChanged(object sender, EventArgs e)
        {
            Object openfile = BrowseDirectory + lstCommands.SelectedItem;
            System.Diagnostics.Process.Start(openfile.ToString());
        }
        private void GetWeather()
        {
            string query = String.Format("http://weather.yahooapis.com/forecastrss?w=2295019");
            XmlDocument wData = new XmlDocument();
            wData.Load(query);

            XmlNamespaceManager manager = new XmlNamespaceManager(wData.NameTable);
            manager.AddNamespace("yweather", "http://xml.weather.yahoo.com/ns/rss/1.0");

            XmlNode channel = wData.SelectSingleNode("rss").SelectSingleNode("channel");
            XmlNodeList nodes = wData.SelectNodes("/rss/channel/item/yweather:forecast", manager);

            Temperature = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["temp"].Value;

            Condition = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["text"].Value;

            Humidity = channel.SelectSingleNode("yweather:atmosphere", manager).Attributes["humidity"].Value;

            WindSpeed = channel.SelectSingleNode("yweather:wind", manager).Attributes["speed"].Value;

            Town = channel.SelectSingleNode("yweather:location", manager).Attributes["city"].Value;
        }
    }       
}
