using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Management.Automation;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace Comet_Updating_System__彗星更新系统_
{
    /// <summary>
    /// Interaction logic for Updater.xaml
    /// </summary>
    public partial class Updater : Window
    {
        #region General Pre-Requisites

        [DllImport("Comet 3\\bin\\CometAuth.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl, EntryPoint = "Verify")]
        [return: MarshalAs(UnmanagedType.I1)]
        public static extern bool Verify([MarshalAs(UnmanagedType.LPStr)] string key);

        [DllImport("Comet 3\\bin\\CometAuth.dll", CallingConvention = CallingConvention.Cdecl, EntryPoint = "HWID")]
        [return: MarshalAs(UnmanagedType.BStr)]
        public static extern string HWID();

        WebClient webClient = new WebClient();
        RegistryKey RegistrySettings = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Supercomet3");
        dynamic CometJSON = JsonConvert.DeserializeObject(HttpGet("https://cometrbx.xyz/external-files/CometJSONAPI.json"));
        DispatcherTimer RGBTime;
        int RGBSpinSpeed = 4;
        int currentprogress = 0;
        public static string HttpGet(string url)
        {
            using (var webClient = new WebClient())
            {
                webClient.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                return webClient.DownloadString(url);
            }
        }

        #endregion

        #region Updater Settings

        string Brand = "Comet";
        string Program = "Comet 3";
        string RenamedProgram = "Comet";
        string Distributor = "wearedevs.net";
        string BootstrapperVersion = "3.0";
        string GithubLink = "https://github.com/MarsQQ/Comet-Updating-System-/tree/master";
        string EULALink = "https://cometrbx.xyz/external-files/EULA";

        bool ExcludeCometFolders = false;

        #endregion

        #region Animation Pre-Requisites

        Storyboard StoryBoard = new Storyboard();
        TimeSpan duration = TimeSpan.FromMilliseconds(500);
        TimeSpan duration2 = TimeSpan.FromMilliseconds(1000);

        private IEasingFunction Smooth
        {
            get;
            set;
        }
        = new QuarticEase
        {
            EasingMode = EasingMode.EaseInOut
        };

        public void Fade(DependencyObject Object)
        {
            DoubleAnimation FadeIn = new DoubleAnimation()
            {
                From = 0.0,
                To = 1.0,
                Duration = new Duration(duration),
            };
            Storyboard.SetTarget(FadeIn, Object);
            Storyboard.SetTargetProperty(FadeIn, new PropertyPath("Opacity", 1));
            StoryBoard.Children.Add(FadeIn);
            StoryBoard.Begin();
            StoryBoard.Children.Remove(FadeIn);
        }

        public void FadeOut(DependencyObject Object)
        {
            DoubleAnimation Fade = new DoubleAnimation()
            {
                From = 1.0,
                To = 0.0,
                Duration = new Duration(duration),
            };
            Storyboard.SetTarget(Fade, Object);
            Storyboard.SetTargetProperty(Fade, new PropertyPath("Opacity", 1));
            StoryBoard.Children.Add(Fade);
            StoryBoard.Begin();
            StoryBoard.Children.Remove(Fade);
        }

        public void ObjectShift(Duration speed, DependencyObject Object, Thickness Get, Thickness Set)
        {
            ThicknessAnimation Animation = new ThicknessAnimation()
            {
                From = Get,
                To = Set,
                Duration = speed,
                EasingFunction = Smooth,
            };
            Storyboard.SetTarget(Animation, Object);
            Storyboard.SetTargetProperty(Animation, new PropertyPath(MarginProperty));
            StoryBoard.Children.Add(Animation);
            StoryBoard.Begin();
            StoryBoard.Children.Remove(Animation);
        }

        public void FasterObjectShift(DependencyObject Object, Thickness Get, Thickness Set)
        {
            ThicknessAnimation Animation = new ThicknessAnimation()
            {
                From = Get,
                To = Set,
                Duration = TimeSpan.FromMilliseconds(750),
                EasingFunction = Smooth,
            };
            Storyboard.SetTarget(Animation, Object);
            Storyboard.SetTargetProperty(Animation, new PropertyPath(MarginProperty));
            StoryBoard.Children.Add(Animation);
            StoryBoard.Begin();
            StoryBoard.Children.Remove(Animation);
        }

        #endregion

        #region Loading Updater

        public Updater()
        {
            InitializeComponent();
            #region RGBSpinner

            RGBTime = new DispatcherTimer(TimeSpan.FromMilliseconds(10), DispatcherPriority.Normal, delegate
            {
                rgbRotation.Angle += RGBSpinSpeed;
            }, System.Windows.Application.Current.Dispatcher);
            RGBTime.Start();

            #endregion

            LogoGrid.Opacity = 0;
            I1.Margin = new Thickness(50, -50, -50, 50);
            I3.Margin = new Thickness(-50, 50, 50, -50);

            LogoGrid2.Opacity = 0;
            II1.Margin = new Thickness(50, -50, -50, 50);
            II3.Margin = new Thickness(-50, 50, 50, -50);

            string verifieddistributor = HttpGet("https://cometrbx.xyz/external-files/verifieddistributorlist.txt");
            string DistributorName = Distributor;

            if (!verifieddistributor.Contains(DistributorName))
                WebsiteLabel.Content = "wearedevs.net";
            else if (DistributorName == "cometrbx.xyz")
                WebsiteLabel.Content = "";
            else
                WebsiteLabel.Content = DistributorName;

            if (!CometJSON.InstallerVer.ToString().Contains(BootstrapperVersion))
            {
                MessageBox.Show("This version of the Updater is outdated, please re-install it from " + "https://" + DistributorName);
                Process.Start("https://" + DistributorName);
                System.Diagnostics.Process.GetCurrentProcess().Kill();
            }

            EULAParagraph.Text = HttpGet(EULALink);
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            await Task.Delay(500);
            Fade(LogoGrid);
            ObjectShift(TimeSpan.FromMilliseconds(1200), I1, I1.Margin, new Thickness(0, 0, 0, 0));
            ObjectShift(TimeSpan.FromMilliseconds(1200), I3, I3.Margin, new Thickness(0, 0, 0, 0));
        }

        #endregion

        #region Functions

        void ExtractFile(string file, string destination)
        {
            try
            {
                ZipFile.ExtractToDirectory(file, destination);
            }
            catch
            {
                MessageBox.Show("I'm sorry, we encountered an error, I'm sorry about this. If you see this message box, please show this to Support, I haven't seen this actually happen yet. [" + Brand + " Extract Process]", Brand + " Updater");
            }
        }

        async void DownloadFile(string file, string destination)
        {
            webClient.DownloadFileAsync(new Uri(file), destination);
            while (webClient.IsBusy)
                await Task.Delay(1000);
        }

        async void ChangeProgress(string Text, int Progress)
        {
            LoadingText.Content = Text;

            // Originally made by MCGamin1738 for Olympus
            for (int i = currentprogress; i <= Progress; i++)
            {
                LoadingBar.Value = i;
                currentprogress = Convert.ToInt32(currentprogress);
                await Task.Delay(5);
            }
        }

        #endregion
        #region Exclude Comet
        void ExcludeComet(string destination)
        {
            // Skidded from Fluxus (Showerhead let me)

            try
            {
                using (PowerShell powerShell = PowerShell.Create())
                {
                    //powerShell.AddScript("Add-MpPreference -ExclusionPath '" + Directory.GetCurrentDirectory() + "'");
                    //powerShell.Invoke();
                    powerShell.AddScript("Add-MpPreference -ExclusionPath '" + System.IO.Path.GetFullPath(destination) + "'");
                    powerShell.Invoke();
                    powerShell.Dispose();
                }
            }
            catch (Exception)
            {
            }
        }
        #endregion

        #region Start Updating
        async void StartUpdating()
        {
            ObjectShift(TimeSpan.FromMilliseconds(1000), ExcludeGrid, ExcludeGrid.Margin, new Thickness(-710, 40, 710, 0));
            ObjectShift(TimeSpan.FromMilliseconds(1000), InstallGrid, InstallGrid.Margin, new Thickness(0, 40, 0, 0));

            await Task.Delay(500);
            Fade(LogoGrid2);
            ObjectShift(TimeSpan.FromMilliseconds(1200), II1, II1.Margin, new Thickness(0, 0, 0, 0));
            ObjectShift(TimeSpan.FromMilliseconds(1200), II3, II3.Margin, new Thickness(0, 0, 0, 0));

            #region Directory Check

            // Written by ImmuneLion318

            ChangeProgress("Adding AutoExec & Workspace Folder.", 10);

            string[] Directories = { Program + "\\autoexec", Program + "\\workspace" };
            for (int i = 0; i < Directories.Length; ++i)
                Directory.CreateDirectory(Directories[i]);

            ChangeProgress("Excluding " + Brand + " Folder.", 20);
            if (ExcludeCometFolders == true)
                ExcludeComet(Directory.GetCurrentDirectory() + "\\" + Program);

            if (Directory.Exists(Program))
            {
                ChangeProgress("Adding Scripts into Scripts Folder.", 30);
                if (!Directory.Exists(Program + "\\scripts"))
                {
                    Directory.CreateDirectory(Program + "\\scripts");
                    webClient.DownloadFileAsync(new Uri("https://cometrbx.xyz/external-files/scripts.zip"), Program + "\\scripts\\scripts.zip");
                    while (webClient.IsBusy)
                        await Task.Delay(1000);
                    ExtractFile(Program + "\\scripts\\scripts.zip", Program + "\\scripts");
                    try { File.Delete(Program + "\\scripts\\scripts.zip"); } catch { }
                }

                ChangeProgress("Deleting old Bin Folder.", 35);

                if (Directory.Exists(Program + "\\bin"))
                    Directory.Delete(Program + "\\bin", true);

                ChangeProgress("Deleting older Comet Program.", 40);
                if (File.Exists(Program + "\\" + Program + ".exe"))
                    File.Delete(Program + "\\" + Program + ".exe");

                if (File.Exists(Program + "\\" + RenamedProgram + ".exe"))
                    File.Delete(Program + "\\" + RenamedProgram + ".exe");
            }

            #endregion
            #region Download Files

            ChangeProgress("Setting Distributor Website.", 50);
            RegistrySettings.SetValue("Distributor", Distributor);

            ChangeProgress("Installing "+Program+".", 60);
            webClient.DownloadFileAsync(new Uri(CometJSON.UIInstall.ToString()), Program + "\\" + Brand + ".zip");
            while (webClient.IsBusy)
                await Task.Delay(1000);

            ChangeProgress("Extracting "+Program+" zip file.", 70);
            ExtractFile(Program + "\\" + Brand + ".zip", Program);

            ChangeProgress("Deleting " + Program + " zip file.", 80);
            try { File.Delete(Program + "\\"+ Brand +".zip"); } catch { }

            #endregion

            ChangeProgress("Excluding " + Program + " DLL file.", 90);
            if (ExcludeCometFolders == true)
                ExcludeComet(@"C:\Users\" + Environment.UserName + @"\AppData\Local\" + HWID());

            ChangeProgress("Done installing " + Program + "!", 100);

            await Task.Delay(1000);

            ObjectShift(TimeSpan.FromMilliseconds(1000), InstallGrid, InstallGrid.Margin, new Thickness(-710, 40, 710, 0));
            ObjectShift(TimeSpan.FromMilliseconds(1000), WelcomeGrid, WelcomeGrid.Margin, new Thickness(0, 40, 0, 0));
        }
        #endregion

        #region Window Controls

        private void ExitB_Click(object sender, RoutedEventArgs e) => Application.Current.Shutdown();

        private void MinimizeB_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try { DragMove(); } catch { }
        }

        private void WebsiteLabel_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                string verifieddistributor = HttpGet("https://cometrbx.xyz/external-files/verifieddistributorlist.txt");
                string DistributorName = Distributor;

                if (!verifieddistributor.Contains(DistributorName))
                    Process.Start("https://wearedevs.net");
                else if (DistributorName == "cometrbx.xyz")
                { }
                else
                    Process.Start("https://" + DistributorName);
            }
            catch { }
        }

        #endregion

        #region Buttons

        private void StartCometProcessB_Click(object sender, RoutedEventArgs e)
        {
            ObjectShift(TimeSpan.FromMilliseconds(1000), IntroGrid, IntroGrid.Margin, new Thickness(-710, 40, 710, 0));
            ObjectShift(TimeSpan.FromMilliseconds(1000), EULAGrid, EULAGrid.Margin, new Thickness(0, 40, 0, 0));
        }

        private void ViewOnGithubB_Click(object sender, RoutedEventArgs e) => Process.Start(GithubLink);

        private void EULAAcceptB_Click(object sender, RoutedEventArgs e)
        {
            ObjectShift(TimeSpan.FromMilliseconds(1000), EULAGrid, EULAGrid.Margin, new Thickness(-710, 40, 710, 0));
            ObjectShift(TimeSpan.FromMilliseconds(1000), JoinDiscordGrid, JoinDiscordGrid.Margin, new Thickness(0, 40, 0, 0));
        }

        private void EULADenyB_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("I'm sorry, but the Terms set in the EULA are necessary to run " + Program + ". If you would like to run " + Program + ", please run the Updater and accept the EULA.", Brand + " Updater");
            Application.Current.Shutdown();
        }

        private void JoinDiscordB_Click(object sender, RoutedEventArgs e)
        { 
            Process.Start(HttpGet("https://cometrbx.xyz/external-files/discord.txt"));
            ObjectShift(TimeSpan.FromMilliseconds(1000), JoinDiscordGrid, JoinDiscordGrid.Margin, new Thickness(-710, 40, 710, 0));
            ObjectShift(TimeSpan.FromMilliseconds(1000), ExcludeGrid, ExcludeGrid.Margin, new Thickness(0, 40, 0, 0));
        }

        private void NotJoinDiscordB_Click(object sender, RoutedEventArgs e)
        {
            ObjectShift(TimeSpan.FromMilliseconds(1000), JoinDiscordGrid, JoinDiscordGrid.Margin, new Thickness(-710, 40, 710, 0));
            ObjectShift(TimeSpan.FromMilliseconds(1000), ExcludeGrid, ExcludeGrid.Margin, new Thickness(0, 40, 0, 0));
        }

        private void ExcludeB_Click(object sender, RoutedEventArgs e)
        {
            ExcludeCometFolders = true;
            StartUpdating();
        }

        private void NoExcludeB_Click(object sender, RoutedEventArgs e)
        {
            ExcludeCometFolders = false;
            StartUpdating();
        }

        private void EnterCometB_Click(object sender, RoutedEventArgs e)
        {

            if (Directory.Exists(Program) && System.IO.File.Exists(Program + "\\" + Program + ".exe"))
            {
                // Rename File:
                // https://www.c-sharpcorner.com/blogs/how-to-rename-a-file-in-c-sharp
                System.IO.FileInfo RenameFile = new System.IO.FileInfo(Program + "\\" + Program + ".exe");
                if (RenameFile.Exists)
                    RenameFile.MoveTo(Program + "\\" + RenamedProgram + ".exe");

                new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        WindowStyle = ProcessWindowStyle.Hidden,
                        FileName = "cmd.exe",
                        Arguments = "/C start " + RenamedProgram + ".exe",
                        WorkingDirectory = Directory.GetCurrentDirectory() + "\\" + Program
                    }
                }.Start();
            }
            else
                MessageBox.Show("Installation has failed! Please make sure your Anti virus is disabled or any other third party program.\r\nOnce disabled, restart this and try again!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

            System.Diagnostics.Process.GetCurrentProcess().Kill();
        }
        #endregion
    }
}
