using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Security.Permissions;
using Microsoft.Win32.SafeHandles;
using System.Runtime.ConstrainedExecution;
using System.Security;

namespace Dashboard
{
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    
    public partial class Form1 : Form
    {
       
        public Form1()
        {
            InitializeComponent();
        }

        private static readonly char[] SpecialChars = "%&*=".ToCharArray();

        //%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        //Fill these out to match your environment
        public static string ArcGISLicenseFileLocation = ""; //The server that has the license manager and this is typically the path
        public static string IISFilesDirectory = ""; //The web server that has IIS.  This sifts through IIS logs to report to you who is using Portal today or by date range
        public static string ArcGISLicenseFileDestination = "";//The directory we make a copy of files to so no locking issues
        public static string IISFilesDestination = ""; //The directory we make a copy of files to so no locking issues
        public static string Adminusername = ""; //I use impersonation in this process so that users without admin rights can at least see who is using ArcGIS Viewer, Advanced, or Portal
        public static string Adminpassword = ""; //I use impersonation in this process so that users without admin rights can at least see who is using ArcGIS Viewer, Advanced, or Portal
        public static string Domainname = ""; //I use impersonation in this process so that users without admin rights can at least see who is using ArcGIS Viewer, Advanced, or Portal
        //%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

        public static string StripHTML(string input)
        {
            return Regex.Replace(input, "<.*?>", String.Empty);
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            var dic = File.ReadAllLines(@"\\YourServer\GISData\GIS_Models\Logging\Parameters\Params.txt") //this is the config file location so you can just change parameters within it instead of the recompiling
              .Select(l => l.Split(new[] { '=' }))
              .ToDictionary(s => s[0].Trim(), s => s[1].Trim());

            ArcGISLicenseFileLocation = dic["ArcGISLicenseFileLocation"];
            IISFilesDirectory = dic["IISFilesDirectory"];
            ArcGISLicenseFileDestination = dic["ArcGISLicenseFileDestination"];
            IISFilesDestination = dic["IISFilesDestination"];
            Adminusername = dic["Adminusername"];
            Adminpassword = dic["Adminpassword"];
            Domainname = dic["Domainname"];

            //*********************************************************************************************************************************************
            Impersonations domain = new Impersonations(Domainname, Adminusername, Adminpassword); //this allows non-Admin users to see who is using Viewer, Advanced, licenses etc because they probably won't have privileges to see the IIS logs and License log
            //*********************************************************************************************************************************************
            

            //Prepare todays date as a parameter to search log file
            DateTime dt = DateTime.Today;
            string month = DateTime.Now.ToString("MMMM");
            int day = dt.Day;
            int year = dt.Year;

            //Overwrites copy of arcgis desktop license file so that we can parse it without locks
            string sourcedir = ArcGISLicenseFileLocation;
            string backupdir = ArcGISLicenseFileDestination;
            string fName = "lmgrd9.log";
            File.Copy(Path.Combine(sourcedir, fName), Path.Combine(backupdir, fName), true);

            //Grab IIS for the day too
            string IISsource = IISFilesDirectory;
            string IISbackup = IISFilesDestination;
            //We need to parse the date to find correct IIS file
            string last2yr = year.ToString().Remove(0, 2);
            int correctmonth = DateTime.Now.Month;
            string cormnth = correctmonth.ToString();
            if (correctmonth < 10)
            {
                cormnth = "0" + correctmonth.ToString();
            }

            int correctday = DateTime.Now.Day;
            string corday = correctday.ToString();
            if (correctday < 10)
            {
                corday = "0" + correctday.ToString();
            }

            string IISName = "u_ex" + last2yr + cormnth + corday;
            File.Copy(Path.Combine(IISsource, IISName + ".log"), Path.Combine(IISbackup, IISName + ".log"), true);

            //IIS Log parse to string
            string suberstr1 = year.ToString() + "-" + cormnth + "-" + day.ToString();
            var lines1 = File.ReadAllLines(Path.Combine(IISbackup, IISName + ".log"));//exception
            List<string> iislist = new List<string>();
            foreach (var line1 in lines1)
            {
                if (line1.Contains(Domainname))
                {

                    string toBeSearched = Domainname;
                    string code = line1.Substring(line1.IndexOf(toBeSearched) + toBeSearched.Length);
                    var firstSpaceIndex = code.IndexOf(" ");
                    var firstString = code.Substring(0, firstSpaceIndex); // INAGX4
                    string finalfirst = firstString.Replace(@"\", "");
 
                    if (finalfirst.Contains('/'))
                    {
                       //do nothing
                    }
                    else
                    {
                        iislist.Add(finalfirst);
                    }
                    
                }
            }

            List<string> distiislist = iislist.Select(o => o.ToString()).Distinct().ToList();
            foreach (string i in distiislist)
            {
                int indexOf = i.IndexOfAny(SpecialChars);
                if (indexOf == -1)
                {
                    // No special chars
                    if (i.ToString().Any(Char.IsWhiteSpace))
                    {
                        //do nothing
                    }
                    else
                    {
                        listBox6.Items.Add(i.Trim());
                    }

                }
            }

            int count = listBox6.Items.Count;
            for (int i = count - 1; i >= 0; i--)
            {
                if (listBox6.Items[i].ToString() == " ")
                {
                    listBox6.Items.RemoveAt(i);
                }
                if (listBox6.Items[i].ToString() == "")
                {
                    listBox6.Items.RemoveAt(i);
                }
            }


            //Locate substring using results above
            string mnt = month.Substring(0, 3);
            string actday = dt.DayOfWeek.ToString();
            actday = actday.Substring(0, 3);
            string suberstr = "Reread time: " + actday + " " + mnt + " " + corday + " " + year.ToString() + " 00:";
            var lines = File.ReadAllLines(Path.Combine(backupdir, fName)).Reverse();
            

            List<string> ViewerUsers = new List<string>();
            List<string> AdvUsers = new List<string>();
            List<string> ThreeDUsers = new List<string>();
            List<string> SkipUsers = new List<string>();
            List<string> SkipUsersInfo = new List<string>();
            List<string> SkipUsers3D = new List<string>();
            Dictionary<string, string> dictionary = new Dictionary<string, string>();
            int counting = 0;
            int debicounter = 0;
            int joshcounter = 0;

            foreach (string line in lines)
            {
                if (line.Contains(suberstr))
                {
                    break;
                }
                else
                {
                    if (line.Contains("Viewer") && line.Contains("(ARCGIS)"))
                    {

                        string str1 = line.Substring(line.LastIndexOf("Viewer") + 8);
                        string usr = "";
                        string user = "";
                        string afterstr = "";
                        int index = str1.IndexOf('@');
                        if (index > 0)
                        {
                            user = str1.Substring(0, index);
                            usr = Regex.Replace(user, @"[^a-zA-Z]", "");
                            afterstr = line.Substring(line.LastIndexOf("Viewer") - 5);
                            afterstr = afterstr.Substring(0, 2);

                            string afterstr1 = str1.Substring(index + 1, (str1.Length - (index + 1)));
                            
                        }

                        if(afterstr == "IN")
                        {
                            SkipUsers.Add(usr + ";" + "IN");
                            //dictionary.Add(usr, "IN");
                        }
                        if(afterstr == "UT")
                        {
                            SkipUsers.Add(usr + ";" + "OUT");
                            //dictionary.Add(usr, "OUT");
                        }

                    }



                    if (line.Contains("(ARCGIS)") && line.Contains("ARC/INFO"))
                    {
                        string str1 = line.Substring(line.LastIndexOf("ARC/INFO") + 8);
                        string usr = "";
                        string user = "";
                        string afterstr = "";
                        int index = str1.IndexOf('@');
                        if (index > 0)
                        {
                            user = str1.Substring(0, index);
                            usr = Regex.Replace(user, @"[^a-zA-Z]", "");
                            afterstr = line.Substring(line.LastIndexOf("ARC/INFO") - 5);
                            afterstr = afterstr.Substring(0, 2);

                            string afterstr1 = str1.Substring(index + 1, (str1.Length - (index + 1)));
                            
                        }

                        if (afterstr == "IN")
                        {
                            SkipUsersInfo.Add(usr + ";" + "IN");
                            //dictionary.Add(usr, "IN");
                        }
                        if (afterstr == "UT")
                        {
                            SkipUsersInfo.Add(usr + ";" + "OUT");
                            //dictionary.Add(usr, "OUT");
                        }

                        //}
                        
                    }
                    if (line.Contains("(ARCGIS)") && line.Contains("TIN"))
                    {
                        string str1 = line.Substring(line.LastIndexOf("TIN") + 5);
                        string usr = "";
                        string user = "";
                        string afterstr = "";
                        int index = str1.IndexOf('@');
                        if (index > 0)
                        {
                            user = str1.Substring(0, index);
                            usr = Regex.Replace(user, @"[^a-zA-Z]", "");
                            afterstr = line.Substring(line.LastIndexOf("TIN") - 5);
                            afterstr = afterstr.Substring(0, 2);

                            string afterstr1 = str1.Substring(index + 1, (str1.Length - (index + 1)));
                            
                        }

                        if (afterstr == "IN")
                        {
                            SkipUsers3D.Add(usr + ";" + "IN");
                            //dictionary.Add(usr, "IN");
                        }
                        if (afterstr == "UT")
                        {
                            SkipUsers3D.Add(usr + ";" + "OUT");
                            //dictionary.Add(usr, "OUT");
                        }
                    }
                    counting++;
                }
            }

            //&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            //build 2 list: one for in's, one for out's per user... Every user must have an IN
            List<string> userlist = new List<string>();
            List<string> Advuserlist = new List<string>();
            List<string> ThreeDeelist = new List<string>();
            List<string> elimlistViewer = new List<string>();
            List<string> elimlistAdv = new List<string>();
            List<string> elimlist3D = new List<string>();
            List<string> userlistIN = new List<string>();
            List<string> userlistOUT = new List<string>();
            int countIN = 0;
            int countOUT = 0;
            foreach(string str in SkipUsers)
            {
                var action = str.Split(';');
                if(action[1].Contains("IN"))
                {
                    userlistIN.Add(action[0]);
                }
                else
                {
                    userlistOUT.Add(action[0]);
                }
            }

            var gin = userlistIN.GroupBy(i => i);
            var gout = userlistOUT.GroupBy(i => i);

            foreach (string str in SkipUsers)
            {
                countIN = 0;
                countOUT = 0;
                
                foreach (var grp in gin)
                {
                    if (str.Contains(grp.Key))
                    {
                        countIN = grp.Count();
                        break;
                    }
                    
                }
                foreach (var grp in gout) 
                {
                    if (str.Contains(grp.Key))
                    {
                        countOUT = grp.Count();
                        break;
                    }

                }

                if(countOUT>countIN)
                {
                    string nw = str.Split(';')[0].Trim();
                    if(!ViewerUsers.Contains(nw))
                    {
                        ViewerUsers.Add(nw);

                    }
                    
                }


            }
            countIN = 0;
            countOUT = 0;
            userlistIN.Clear();
            userlistOUT.Clear();

            //---------------------------------------------------------------------------------------------------
            foreach (string str in SkipUsersInfo)
            {
                var action = str.Split(';');
                if (action[1].Contains("IN"))
                {
                    userlistIN.Add(action[0]);
                }
                else
                {
                    userlistOUT.Add(action[0]);
                }
            }

            var ginInfo = userlistIN.GroupBy(i => i);
            var goutInfo = userlistOUT.GroupBy(i => i);

            foreach (string str in SkipUsersInfo)
            {
                foreach (var grp in ginInfo)
                {
                    if (str.Contains(grp.Key))
                    {
                        countIN = grp.Count();
                        break;
                    }

                }
                foreach (var grp in goutInfo)
                {
                    if (str.Contains(grp.Key))
                    {
                        countOUT = grp.Count();
                        break;
                    }

                }

                if (countOUT > countIN)
                {
                    string nw = str.Split(';')[0].Trim();
                    if (!AdvUsers.Contains(nw))
                    {
                        AdvUsers.Add(nw);
                    }

                }
            }
            countIN = 0;
            countOUT = 0;
            userlistIN.Clear();
            userlistOUT.Clear();
            //------------------------------------------------------------------------------------------------------------------
            foreach (string str in SkipUsers3D)
            {
                var action = str.Split(';');
                if (action[1].Contains("IN"))
                {
                    userlistIN.Add(action[0]);
                }
                else
                {
                    userlistOUT.Add(action[0]);
                }
            }

            var gin3D = userlistIN.GroupBy(i => i);
            var gout3D = userlistOUT.GroupBy(i => i);

            foreach (string str in SkipUsers3D)
            {
                foreach (var grp in gin3D)
                {
                    if (str.Contains(grp.Key))
                    {
                        countIN = grp.Count();
                        break;
                    }

                }
                foreach (var grp in gout3D)
                {
                    if (str.Contains(grp.Key))
                    {
                        countOUT = grp.Count();
                        break;
                    }

                }

                if (countOUT > countIN)
                {
                    string nw = str.Split(';')[0].Trim();
                    if (!ThreeDUsers.Contains(nw))
                    {
                        ThreeDUsers.Add(nw);
                    }

                }
            }
            countIN = 0;
            countOUT = 0;
            userlistIN.Clear();
            userlistOUT.Clear();
            //&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&



            //%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

            int newcount = 0;
            List<string> userlistLatestA = new List<string>();
            foreach (string b in ViewerUsers)
            {
                userlistLatestA.Add(b);
                

            }
            //%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

            //Lets get only records that match that user for the day
            List<string> userlistLatestAdv = new List<string>();
            foreach (string item in AdvUsers)
            {
                userlistLatestAdv.Add(item);
                
            }

            
            //######################################################################

 
            List<string> userlistLatestAdv3D = new List<string>();
            foreach (string item in ThreeDUsers)
            {
                userlistLatestAdv3D.Add(item);
                
            }

            
            //######################################################################

            foreach (var item in userlistLatestA)//distlist
            {
                listBox4.Items.Add(item);
                

            }
            foreach (var item in userlistLatestAdv)//distlistAdv
            {
                int indexOf = item.IndexOfAny(SpecialChars);
                if (indexOf == -1)
                {
                    // No special chars
                    if (item.ToString().Any(Char.IsWhiteSpace))
                    {

                    }
                    else
                    {
                        listBox5.Items.Add(item.Trim());
                    }
                }

                

            }
            foreach (var item in userlistLatestAdv3D)//distlistAdv
            {
                int indexOf = item.IndexOfAny(SpecialChars);
                if (indexOf == -1)
                {
                    // No special chars
                    if (item.Length < 1)
                    {

                    }
                    else
                    {
                        listBox1.Items.Add(item.Trim());


                    }
                }

            }

            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            label7.Visible = true;
            label6.Visible = true;
            dateTimePicker1.Visible = true;
            dateTimePicker2.Visible = true;
            dateTimePicker1.ResetText();
            dateTimePicker2.ResetText();
            //button2.Visible = true;

            label5.Text = "";
            label5.Visible = false;

            listBox6.Items.Clear();

            progressBar1.Value = 0;
            progressBar1.Visible = true;

            
        }

        public int GetMonth()
        {
            int Result = 0;
            Result = dateTimePicker1.Value.Month;
            return Result;
        }
        public int GetDay()
        {
            int Result = 0;
            Result = dateTimePicker1.Value.Day;
            return Result;
        }
        public int GetYear()
        {
            int Result = 0;
            Result = dateTimePicker1.Value.Year;
            return Result;
        }

        public int GetMonth2()
        {
            int Result = 0;
            Result = dateTimePicker2.Value.Month;
            return Result;
        }
        public int GetDay2()
        {
            int Result = 0;
            Result = dateTimePicker2.Value.Day;
            return Result;
        }
        public int GetYear2()
        {
            int Result = 0;
            Result = dateTimePicker2.Value.Year;
            return Result;
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            

            string last2yr = "";
            int correctmonth = 0;//GetMonth(); //DateTime.Now.Month;
            string cormnth = "";
            //Decipher date from date pickers above and grab copy of all IIS logs matching that date range
            //Grab IIS for the day too
            List<string> IISLogsList = new List<string>();
            string IISsource = @"\\WebserverServer\C$\inetpub\logs\LogFiles\W3SVC1";
            string IISbackup = @"\\LicenseManagerServer\GISData\GIS_Models\Logging\Dates";

            List<string> alliislist = new List<string>();

            //Now add first IIS log info to list and add all inbetween through IIS log datetime2
            var dates = new List<DateTime>();

            for (var dts = dateTimePicker1.Value; dts <= dateTimePicker2.Value; dts = dts.AddDays(1))
            {
                dates.Add(dts);
            }

            int counts = 0;
            progressBar1.Maximum = dates.Count;
            progressBar1.Step = 1;
            progressBar1.Value = 0;

            foreach (DateTime dv in dates)
            {
                DateTime dt = DateTime.Today;
                string month = DateTime.Now.ToString("MMMM");
                int year = dv.Year;//GetYear();//dt.Year;

                //We need to parse the date to find correct IIS file
                last2yr = year.ToString().Remove(0, 2);
                correctmonth = dv.Month;//GetMonth(); //DateTime.Now.Month;
                cormnth = correctmonth.ToString();
                if (correctmonth < 10)
                {
                    cormnth = "0" + correctmonth.ToString();
                }

                int correctday = dv.Day;//GetDay();//DateTime.Now.Day;
                string corday = correctday.ToString();
                if (correctday < 10)
                {
                    corday = "0" + correctday.ToString();
                }

                //Add userid to listbox6
                //IIS Log parse to string
                string suberstr1 = year.ToString() + "-" + cormnth + "-" + corday;
                string IISName = "u_ex" + last2yr + cormnth + corday;
                if (File.Exists(Path.Combine(IISsource, IISName + ".log")))
                {
                    File.Copy(Path.Combine(IISsource, IISName + ".log"), Path.Combine(IISbackup, IISName + ".log"), true);
                }
                
                if (File.Exists(Path.Combine(IISsource, IISName + ".log")))
                {
                    var lines1 = File.ReadAllLines(Path.Combine(IISbackup, IISName + ".log"));//exception
                    List<string> iislist = new List<string>();
                    foreach (var line1 in lines1)
                    {
                        if (line1.Contains(Domainname))
                        {

                            string toBeSearched = Domainname;
                            string code = line1.Substring(line1.IndexOf(toBeSearched) + toBeSearched.Length);
                            var firstSpaceIndex = code.IndexOf(" ");
                            var firstString = code.Substring(0, firstSpaceIndex); // INAGX4
                            string finalfirst = firstString.Replace(@"\", "");

                            if (finalfirst.Contains('/'))
                            {

                            }
                            else
                            {
                                iislist.Add(finalfirst);
                                alliislist.Add(finalfirst);
                            }

                        }
                    }

                   
                    counts = counts + 1;
                    progressBar1.Value = counts;
                }
            }
            int counting = 0;
            List<string> distiislist1 = alliislist.Select(o => o.ToString()).Distinct().ToList();
            foreach (string i in distiislist1)
            {
                //listBox6.Items.Add(i);
                int indexOf = i.IndexOfAny(SpecialChars);
                if (indexOf == -1)
                {
                    // No special chars
                    if (i.Length < 1)
                    {
                    }
                    else
                    {
                        listBox6.Items.Add(i.Trim());
                    }

                }
            }


            label5.Text = listBox6.Items.Count.ToString() + " users";
            label5.Visible = true;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //I have not included the ability to export listbox6 contents to csv or xls yet
        }

        private void listBox6_KeyDown(object sender, KeyEventArgs e)
        {
            //no code needed yet
        }

        private void listBox6_KeyDown_1(object sender, KeyEventArgs e)
        {
            //no code needed yet
        }
    }

    public class Impersonations : IDisposable
    {
        private readonly SafeTokenHandle _handle;
        private readonly WindowsImpersonationContext _context;

        const int LOGON32_LOGON_NEW_CREDENTIALS = 9;

        public Impersonations(string domain, string username, string password)
        {
            var ok = LogonUser(username, domain, password,
                           LOGON32_LOGON_NEW_CREDENTIALS, 0, out this._handle);
            if (!ok)
            {
                var errorCode = Marshal.GetLastWin32Error();
                throw new ApplicationException(string.Format("Could not impersonate the elevated user.  LogonUser returned error code {0}.", errorCode));
            }

            this._context = WindowsIdentity.Impersonate(this._handle.DangerousGetHandle());
        }

        public void Dispose()
        {
            this._context.Dispose();
            this._handle.Dispose();
        }

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool LogonUser(String lpszUsername, String lpszDomain, String lpszPassword, int dwLogonType, int dwLogonProvider, out SafeTokenHandle phToken);

        public sealed class SafeTokenHandle : SafeHandleZeroOrMinusOneIsInvalid
        {
            private SafeTokenHandle()
                : base(true) { }

            [DllImport("kernel32.dll")]
            [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
            [SuppressUnmanagedCodeSecurity]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool CloseHandle(IntPtr handle);

            protected override bool ReleaseHandle()
            {
                return CloseHandle(handle);
            }
        }
    }
}
