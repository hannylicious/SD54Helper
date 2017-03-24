using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Linq;
using System.Runtime.InteropServices;
using IWshRuntimeLibrary;
using System.DirectoryServices.AccountManagement;
using System.Management;
using System.Security.Principal;
using System.Threading.Tasks;
using System.Threading;

namespace Helpdesk54
{
    public partial class MainForm : Form

    {
        int existingSetupRowNumber, existingBackupRowNumber;
        long canItFit;
        string userDisplayFirstName, userDisplayLastName, userCustomDisplayName;
        string backupDirectoryName;
        string backupName;
        string backupToRestore;
        string serverName;
        public BackgroundWorker essentialBgWorker = new BackgroundWorker();
        public BackgroundWorker additionalBgWorker = new BackgroundWorker();
        public BackgroundWorker restoreEssentialBgWorker = new BackgroundWorker();
        public BackgroundWorker restoreAdditionalBgWorker = new BackgroundWorker();
        string clickedButton;
        string itemsChanged;
        long selectedDriveAvailableSize;

        public MainForm()
        {
            Hide();
            bool done = false;
            ThreadPool.QueueUserWorkItem((x) =>
            {
                using (var splashForm = new Splash())
                {
                    splashForm.Show();
                    while (!done)
                        Application.DoEvents();
                    splashForm.Close();
                }
            });
            InitializeComponent();

            //setup background worker for progress bars
            //essentialBgWorker
            essentialBgWorker.DoWork += new DoWorkEventHandler(essentialBgWorker_DoWork);
            essentialBgWorker.ProgressChanged += new ProgressChangedEventHandler(essentialBgWorker_ProgressChanged);
            essentialBgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(essentialBgWorker_RunWorkerCompleted);
            essentialBgWorker.WorkerReportsProgress = true;
            //additionalBgWorker
            additionalBgWorker.DoWork += new DoWorkEventHandler(additionalBgWorker_DoWork);
            additionalBgWorker.ProgressChanged += new ProgressChangedEventHandler(additionalBgWorker_ProgressChanged);
            additionalBgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(additionalBgWorker_RunWorkerCompleted);
            additionalBgWorker.WorkerReportsProgress = true;
            //restoreEssentialsWorker
            restoreEssentialBgWorker.DoWork += new DoWorkEventHandler(restoreEssentialBgWorker_DoWork);
            restoreEssentialBgWorker.ProgressChanged += new ProgressChangedEventHandler(restoreEssentialBgWorker_ProgressChanged);
            restoreEssentialBgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(restoreEssentialBgWorker_RunWorkerCompleted);
            restoreEssentialBgWorker.WorkerReportsProgress = true;
            //restoreAdditionalWorker
            restoreAdditionalBgWorker.DoWork += new DoWorkEventHandler(restoreAdditionalBgWorker_DoWork);
            restoreAdditionalBgWorker.ProgressChanged += new ProgressChangedEventHandler(restoreAdditionalBgWorker_ProgressChanged);
            restoreAdditionalBgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(restoreAdditionalBgWorker_RunWorkerCompleted);
            restoreAdditionalBgWorker.WorkerReportsProgress = true;
            restoreAdditionalBgWorker.WorkerSupportsCancellation = true;

            //Get the userName to the currently logged in user
            //also gets userCustomDisplayName
            string userName = getUsername().ToString();
            string customUserName = getUserCustomDisplayName().ToString();
            usernameLabel.Text = userName;
            //setup the backup tab
            setBackupPageOptions();
            //set restoreDriveCombo to dropdown
            setRestoreDriveCombo();
            //check if the user gets access to quicken
            if (doesUserGetQuicken(userName))
            {
                quickenButton.Enabled = true;
                quickenCheckBox.Enabled = true;
                quickenBackupCheckBox.Enabled = true;
            } else {
                quickenButton.Enabled = false;
                quickenCheckBox.Enabled = false;
                quickenBackupCheckBox.Enabled = false;
            }
            //set the serverNameLink to link to the appropriate location
            setServerNameLink();
            //check installations
            checkApplicationInstalls();
            //check to see if user has been backed up or restored already
            if (userHasBeenSetup(userName))
            {
                updateUserSetupChecks();
                userSetupAnswerLabel.Text = "YES";
                userSetupAnswerLabel.ForeColor = System.Drawing.Color.ForestGreen;

            }
            if (userHasBeenBackedUp(userName))
            {
                updateUserBackupChecks();
                userBackedUpAnswerLabel.Text = "YES";
                userBackedUpAnswerLabel.ForeColor = System.Drawing.Color.ForestGreen;
            }
            //update labels
            labelDirectorySizes(userName);

            done = true;
            Show();

        }
        private void setBackupPageOptions()
        {
            string userName = getUsername().ToString();
            //set the backupDirectoryName
            backupDirectoryName = userName.ToString() + "-54Help-" + DateTime.Now.Year.ToString();
            //load existing users
            setUserSelectBackupCombo();
            //set backupDriveCombo dropdown
            setBackupDriveCombo();
            //Set the selected drive freespace label 
            setBackupDriveFreeSpaceLabel();
        }
        /// <summary>
        /// Sets the restore page options. Called when restore drive combo changes.
        /// </summary>
        private void setRestorePageOptions()
        {
            userRestoreSelectCombo.Items.Clear();
            string directoryToHouseBackups = "54HelperBackups";
            string selectedDrive = restoreDriveCombo.SelectedItem.ToString();
            string fullPathForBackups = selectedDrive + directoryToHouseBackups;
            //there are backups
            if (doBackupsExist(fullPathForBackups))
            {
                //do things because backups were found
                backupFoundLabel.Text = "(Backups Found!)";
                backupFoundLabel.ForeColor = System.Drawing.Color.ForestGreen;
                userRestoreSelectCombo.Visible = true;
                //set the dropdown to display the existing backups
                string[] backupDirectories = Directory.GetDirectories(fullPathForBackups);
                foreach (string backupDirectory in backupDirectories)
                {
                    string shortenedBackupDirectory = backupDirectory.Contains("\\") && backupDirectory.Split('\\').GetLength(0) > 2 ? string.Join("\\", backupDirectory.Split('\\').Skip(2).ToList()) : backupDirectory;
                    userRestoreSelectCombo.Items.Add(shortenedBackupDirectory);
                    //set the current users backup as default if there is one
                    if (backupDirectory.Contains(backupDirectoryName))
                    {
                        userRestoreSelectCombo.SelectedItem = shortenedBackupDirectory;
                    }
                }
            }
            else //there are no backups
            {
                //mark no backups found
                backupFoundLabel.Text = "(No Backup Found!)";
                backupFoundLabel.ForeColor = System.Drawing.Color.Crimson;
                userRestoreSelectCombo.Visible = false;
            }
        }

        private bool doBackupsExist(string fullPathForBackups)
        {            
            if (Directory.Exists(fullPathForBackups))
            {
                int directoryCount = Directory.GetDirectories(fullPathForBackups).Length;
                if ( directoryCount > 0 ) {
                    return true;
                } 
                else 
                {
                    return false;
                }
                
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Loads the existing windows users.
        /// </summary>
        private void setUserSelectBackupCombo()
        {
            string userName = getUsername().ToString();
            userBackupSelectCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            string[] folders = System.IO.Directory.GetDirectories(@"C:\Users\", "*", System.IO.SearchOption.TopDirectoryOnly);

            foreach (string folder in folders)
            {
                //get the username (strip off 'C:\Users\'
                string localUserFolderName = folder.Substring(9);
                if (localUserFolderName == "All Users" || localUserFolderName == "Public")
                {

                } else
                {
                    userBackupSelectCombo.Items.Add(localUserFolderName);
                }
                
            }
            userBackupSelectCombo.SelectedIndex = userBackupSelectCombo.FindString(userName);
        }

        /// <summary>
        /// Sets the backup drive combo.
        /// </summary>
        private void setBackupDriveCombo()
        {
            string homeDirectory = getHomeDirectory().ToString();
            Array theDrives = getAttachedDrives(); ;
            //set it
            backupDriveCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            foreach (DriveInfo currentDrive in theDrives)
            {
                if (currentDrive.IsReady == true)
                {
                    backupDriveCombo.Items.Add(currentDrive);
                }
            }
            //Set the selected drive to the H:/ Drive or the first in the list if unknown
            if (homeDirectory != "Unknown")
            {
                backupDriveCombo.SelectedIndex = backupDriveCombo.FindString(homeDirectory);
            }
            else
            {
                backupDriveCombo.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// Sets the restore drive combo.
        /// </summary>
        private void setRestoreDriveCombo()
        {
            string homeDirectory = getHomeDirectory().ToString();
            Array theDrives = getAttachedDrives();
            //Set it
            restoreDriveCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            foreach (DriveInfo currentDrive in theDrives)
            {
                if (currentDrive.IsReady == true)
                {
                    restoreDriveCombo.Items.Add(currentDrive);
                }
            }
            //Try to set the selected drive to the H:/ Drive
            if (homeDirectory != "Unknown")
            {
                restoreDriveCombo.SelectedIndex = restoreDriveCombo.FindString(homeDirectory);
            }
            else
            {
                restoreDriveCombo.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// Sets the backup drive free space label.
        /// </summary>
        private void setBackupDriveFreeSpaceLabel()
        {
            //Set the selected drive freespace label        
            DriveInfo selectedDrive = (DriveInfo)backupDriveCombo.SelectedItem;
            //If it's the homedrive (or a network drive) - we need to do further calculation to determine 'available free space'
            //Max Available : 3GB
            if (selectedDrive.DriveType == DriveType.Network)
            {
                long maxAvailableDriveSpace = 3221225472;
                string folder = selectedDrive.ToString();
                long networkDirectorySize = DirSize(new DirectoryInfo(folder));
                selectedDriveAvailableSize = maxAvailableDriveSpace - networkDirectorySize;
                string folderMB = FormatBytes(selectedDriveAvailableSize);
                //folderMB = used space on network drive
                backupDriveLabel.Text = folderMB + " Free";
            }
            //otherwise prepare as normal and show appropriate free space
            else if (selectedDrive.AvailableFreeSpace > 0)
            {
                long driveSpace = selectedDrive.AvailableFreeSpace;
                string driveFreeSpace = FormatBytes(driveSpace);
                backupDriveLabel.Text = driveFreeSpace + " Free";
            }
        }
        private string getHomeDirectory()
        {
            Array theDrives = getAttachedDrives();
            string homeDirectory = "";
            string userName = getUsername().ToString();
            foreach (DriveInfo currentDrive in theDrives)
            {
                if (currentDrive.DriveType == DriveType.Network)
                {
                    string currentDriveString = currentDrive.Name.ToString();
                    string path = GetUNCPath(currentDriveString);
                    if (path.ToLower().Contains(userName.ToLower()))
                    {
                        // This is the Home drive! // 
                        homeDirectory = currentDrive.Name;
                    }
                    if (path.ToLower().Contains(userCustomDisplayName.ToLower()))
                    {
                        // This is the Home drive! // 
                        homeDirectory = currentDrive.Name;
                    }
                }
                else
                {
                    homeDirectory = "Unknown";
                }
            }
            return homeDirectory;
        }

        /// <summary>
        /// Sets the server name link.
        /// </summary>
        private void setServerNameLink()
        {
            string serverName;
            Array theDrives = getAttachedDrives();
            string homeDirectory = getHomeDirectory().ToString();
            string userName = getUsername().ToString();
            foreach (DriveInfo currentDrive in theDrives)
            {
                if (currentDrive.DriveType == DriveType.Network)
                {
                    string currentDriveString = currentDrive.Name.ToString();
                    string path = GetUNCPath(currentDriveString);
                    if (path.ToLower().Contains(userName.ToLower()))
                    {
                        // This is the Home drive! // 
                        homeDirectory = currentDrive.Name;
                        Uri uri = new Uri(path);
                        serverName = uri.Host.ToString();
                        serverNameLinkLabel.Text = serverName;
                    }
                    if (path.ToLower().Contains(userCustomDisplayName.ToLower()))
                    {
                        // This is the Home drive! // 
                        homeDirectory = currentDrive.Name;
                        Uri uri = new Uri(path);
                        serverName = uri.Host.ToString();
                        serverNameLinkLabel.Text = serverName;
                    }
                }
                else
                {
                    homeDirectory = "Unknown";
                    serverNameLinkLabel.Text = "Unknown";
                }
            }
        }

        private string getWordPath()
        {
            string wordPath;
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe");
            if (key == null)
            {
                wordPath = "";
                return wordPath;
            }
            else
            {
                wordPath = key.GetValue("").ToString();
                return wordPath;
            }
        }

        private string getExcelPath()
        {
            string excelPath;
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe");
            if (key == null)
            {
                excelPath = "";
                return excelPath;
            }
            else
            {
                excelPath = key.GetValue("").ToString();
                return excelPath;
            }
        }

        private string getOutlookPath()
        {
            string outlookPath;
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.exe");
            if (key == null)
            {
                outlookPath = "";
                return outlookPath;
            }
            else
            {
                outlookPath = key.GetValue("").ToString();
                return outlookPath;
            }
        }

        /// <summary>
        /// Gets the attached drives.
        /// </summary>
        /// <returns>Array theDrives</returns>
        private Array getAttachedDrives()
        {
            DriveInfo[] theDrives;
            //get the drives and set them to an array
            theDrives = DriveInfo.GetDrives();
            return theDrives;

        }

        /// <summary>
        /// Gets the username.
        /// </summary>
        /// <returns>String userName</returns>
        private string getUsername()
        {
            //set username - Do this early!
            string userName = Environment.UserName;
            userDisplayFirstName = UserPrincipal.Current.GivenName;
            userDisplayLastName = UserPrincipal.Current.Surname;
            userCustomDisplayName = userDisplayFirstName + userDisplayLastName;
            return userName;
        }
        private string getUserToBeBackedUp()
        {
            string userToBeBackedUp = userBackupSelectCombo.SelectedItem.ToString();
            return userToBeBackedUp;
        }
        private string getSelectedDrive()
        {
            string selectedDrive = backupDriveCombo.SelectedItem.ToString();
            return selectedDrive;
        }
        /// <summary>
        /// Gets the custom display name of the user.
        /// </summary>
        /// <returns>String userCustomDisplayName</returns>
        private string getUserCustomDisplayName()
        {
            userDisplayFirstName = UserPrincipal.Current.GivenName;
            userDisplayLastName = UserPrincipal.Current.Surname;
            userCustomDisplayName = userDisplayFirstName + userDisplayLastName;
            return userCustomDisplayName;
        }
        private bool userHasCustomName(string userName, string userCustomDisplayName)
        {
            if (userName != userCustomDisplayName)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// Handles the LinkClicked event of the serverNameLinkLabel control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="LinkLabelLinkClickedEventArgs"/> instance containing the event data.</param>
        private void serverNameLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            serverNameLinkLabel.LinkVisited = true;
            System.Diagnostics.Process.Start("explorer", @"\\" + serverNameLinkLabel.Text.ToString());
        }

        /// <summary>
        /// Checks the application installs.
        /// </summary>
        private void checkApplicationInstalls()
        {
            string userToBeBackedUp = userBackupSelectCombo.SelectedItem.ToString();
            string[] applicationList =
            {
                "DYMO Label v.8",
                "Adobe Acrobat X Pro",
                "ScanSnap Manager",
                "Quicken 2011",
            };
            //check for outlook
            Type officeType = Type.GetTypeFromProgID("Outlook.Application");
            if (officeType == null)
            {
                // Outlook is not installed.   
                outlookButton.Text = "Install Outlook";
                outlookButton.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                // Outlook is installed.    
                outlookButton.Text = "Open Outlook";
                outlookButton.ForeColor = System.Drawing.Color.Green;
            }
            //go through the application list
            foreach (string applicationName in applicationList)
            {
                //if it's installed do this block depending on application
                if (IsApplicationInstalled(applicationName))
                {
                    switch (applicationName)
                    {
                        case "DYMO Label v.8":
                            dymoButton.Text = "Open Dymo";
                            dymoButton.ForeColor = System.Drawing.Color.Green;
                            break;
                        case "Adobe Acrobat X Pro":
                            acrobatButton.Text = "Open Acrobat Pro";
                            acrobatButton.ForeColor = System.Drawing.Color.Green;
                            break;
                        case "ScanSnap Manager":
                            scanSnapButton.Text = "Open ScanSnap";
                            scanSnapButton.ForeColor = System.Drawing.Color.Green;
                            break;
                        case "Quicken 2011":
                            quickenButton.Text = "Open Quicken";
                            quickenButton.ForeColor = System.Drawing.Color.Green;
                            break;
                    }
                    //not installed do this block
                }
                else
                {
                    switch (applicationName)
                    {
                        case "DYMO Label v.8":
                            dymoButton.Text = "Install Dymo";
                            dymoButton.ForeColor = System.Drawing.Color.Red;
                            break;
                        case "Adobe Acrobat X Pro":
                            acrobatButton.Text = "Install Acrobat Pro";
                            acrobatButton.ForeColor = System.Drawing.Color.Red;
                            break;
                        case "ScanSnap Manager":
                            scanSnapButton.Text = "Install ScanSnap";
                            scanSnapButton.ForeColor = System.Drawing.Color.Red;
                            break;
                        case "Quicken 2011":
                            quickenButton.Text = "Install Quicken";
                            quickenButton.ForeColor = System.Drawing.Color.Red;
                            break;
                    }
                }
            }
            //check to see if stickynotes has been ran - if file exists enable button - otherwise disable and set to 'none'
            string stickyNotesDirectory;
            if (WinMajorVersion == 10)
            //windows 10
            {
                stickyNotesDirectory = "C:\\Users\\" + userToBeBackedUp + "\\AppData\\Local\\Packages\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe";
            }
            else
            {
                //windows 7
                stickyNotesDirectory = "C:\\Users\\" + userToBeBackedUp + "\\AppData\\Roaming\\Microsoft\\Sticky Notes";
            }
            if (Directory.Exists(stickyNotesDirectory))
            {
                backupStickyNotesButton.Enabled = true;
            }
            else
            {
                backupStickyNotesButton.Enabled = false;
            }

        }

        /// <summary>
        /// Opens the software click.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        void openSoftwareClick(object sender, EventArgs e)
        {
            string path;
            Button sentButton = sender as Button;
            switch (sentButton.Text.ToString())
            {
                case "Open Outlook":
                    path = getOutlookPath().ToString();
                    if (System.IO.File.Exists(path))
                    {
                        System.Diagnostics.Process.Start(path);
                    }
                    else
                    {
                        MessageBox.Show("There is an unknown error with the location of the Outlook installation or it has not been installed.");
                    }
                    break;
                case "Open Dymo":
                    path = @"C:\Program Files (x86)\DYMO\DYMO Label Software\DLS.exe";
                    if (System.IO.File.Exists(path))
                    {
                        System.Diagnostics.Process.Start(path);
                    }
                    else
                    {
                        System.Diagnostics.Process.Start(@"C:\Program Files\DYMO\DYMO Label Software\DLS.exe");
                    }
                    break;
                case "Open Acrobat Pro":
                    path = @"C:\Program Files (x86)\Adobe\Acrobat 10.0\Acrobat\Acrobat.exe";
                    if (System.IO.File.Exists(path))
                    {
                        System.Diagnostics.Process.Start(path);
                    }
                    else
                    {
                        System.Diagnostics.Process.Start(@"C:\Program Files\Adobe\Acrobat 10.0\Acrobat\Acrobat.exe");
                    }
                    break;
                case "Open ScanSnap":
                    path = @"C:\Program Files (x86)\PFU\ScanSnap\Driver\PfuSsMon.exe";
                    if (System.IO.File.Exists(path))
                    {
                        System.Diagnostics.Process.Start(path);
                    }
                    else
                    {
                        System.Diagnostics.Process.Start(@"C:\Program Files\PFU\ScanSnap\Driver\PfuSsMon.exe");
                    }
                    break;
                case "Open Quicken":
                    path = @"C:\Program Files (x86)\Quicken\qw.exe";
                    if (System.IO.File.Exists(path))
                    {
                        System.Diagnostics.Process.Start(path);
                    }
                    else
                    {
                        System.Diagnostics.Process.Start(@"C:\Program Files\Quicken\qw.exe");
                    }
                    break;
            }
        }

        /// <summary>
        /// Installs the software click.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        void installSoftwareClick(object sender, EventArgs e)
        {
            string path;
            Button sentButton = sender as Button;
            switch (sentButton.Text.ToString())
            {
                case "Install Outlook":
                    System.Diagnostics.Process.Start("control", "appwiz.cpl");
                    break;
                case "Install Dymo":
                    path = @"\\dataserver02\PCUpdate\SecretaryInstalls\Dymo\Setup.exe";
                    if (System.IO.File.Exists(path))
                    {
                        System.Diagnostics.Process.Start(path);
                    }
                    else
                    {
                        MessageBox.Show("The software is not available for installation at this time.");
                    }
                    break;
                case "Install Acrobat Pro":
                    path = @"\\dataserver02\PCUpdate\SecretaryInstalls\AcrobatXPro\AcroPro.msi";
                    if (System.IO.File.Exists(path))
                    {
                        System.Diagnostics.Process.Start(path);
                    }
                    else
                    {
                        MessageBox.Show("The software is not available for installation at this time.");
                    }
                    break;
                case "Install ScanSnap":
                    path = @"\\dataserver02\PCUpdate\SecretaryInstalls\ScanSnap\setup.exe";
                    if (System.IO.File.Exists(path))
                    {
                        System.Diagnostics.Process.Start(path);
                    }
                    else
                    {
                        MessageBox.Show("The software is not available for installation at this time.");
                    }
                    break;
                case "Install Quicken":
                    path = @"\\dataserver02\PCUpdate\SecretaryInstalls\Quicken2011\DISK1\Setup.exe";
                    if (System.IO.File.Exists(path))
                    {
                        System.Diagnostics.Process.Start(path);
                    }
                    else
                    {
                        MessageBox.Show("The software is not available for installation at this time.");
                    }
                    break;
            }
        }

        /// <summary>
        /// Formats the bytes string into appropriately size (GB, MB, etc.).
        /// </summary>
        /// <param name="bytes">The bytes</param>
        /// <returns></returns>
        public string FormatBytes(long bytes)
        {
            const int scale = 1024;
            string[] orders = new string[] { "GB", "MB", "KB", "Bytes" };
            long max = (long)Math.Pow(scale, orders.Length - 1);

            foreach (string order in orders)
            {
                if (bytes > max)
                    return string.Format("{0:##.##} {1}", decimal.Divide(bytes, max), order);

                max /= scale;
            }
            return "0 Bytes";
        }

        /// <summary>
        /// URLs the shortcut to desktop.
        /// </summary>
        /// <param name="linkName">Name of the link.</param>
        /// <param name="linkUrl">The link URL.</param>
        private void urlShortcutToDesktop(string linkName, string linkUrl)
        {
            string deskDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            using (StreamWriter writer = new StreamWriter(deskDir + "\\" + linkName + ".url"))
            {
                writer.WriteLine("[InternetShortcut]");
                writer.WriteLine("URL=" + linkUrl);
                writer.WriteLine("IDList=");
                writer.WriteLine("IconFile=" + linkUrl + "/favicon.ico");
                writer.WriteLine("IconIndex=1");
                writer.Flush();
            }
        }

        /// <summary>
        /// Gets the size of the directory.
        /// </summary>
        /// <param name="d">The directory</param>
        /// <returns>Long size</returns>
        public static long DirSize(DirectoryInfo d)
        {
            long size = 0;
            // Add file sizes.
            FileInfo[] fis = d.GetFiles();
            foreach (FileInfo fi in fis)
            {
                size += fi.Length;
            }
            // Add subdirectory sizes.
            DirectoryInfo[] dis = d.GetDirectories();
            foreach (DirectoryInfo di in dis)
            {
                if ((di.Name == "My Music") || (di.Name == "My Pictures") || (di.Name == "My Videos"))
                {
                    //do nothing with it since it is just a system link we cannot access
                }
                else
                {
                    size += DirSize(di);
                };
            }
            return size;
        }

        /// <summary>
        /// Creates the shortcut.
        /// </summary>
        /// <param name="shortcutName">Name of the shortcut.</param>
        /// <param name="shortcutPath">The shortcut path.</param>
        /// <param name="targetFileLocation">The target file location.</param>
        public static void CreateShortcut(string shortcutName, string shortcutPath, string targetFileLocation)
        {
            string shortcutLocation = System.IO.Path.Combine(shortcutPath, shortcutName + ".lnk");
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutLocation);

            //shortcut.Description = "My shortcut description";   // The description of the shortcut
            //shortcut.IconLocation = @"c:\myicon.ico";           // The icon of the shortcut
            if (shortcutName == "Infinite Campus")
            {
                shortcut.TargetPath = targetFileLocation;                 // The path of the file that will launch when the shortcut is run
                shortcut.Arguments = @"https://campus.sd54.org/campus/schaumburg.jsp";
                shortcut.IconLocation = @"\\dataserver02\PCUpdate\54Helper\icons\campus.ico";           // The icon of the shortcut
                shortcut.Save();                                    // Save the shortcut
            }
            if (shortcutName == "AESOP")
            {
                shortcut.TargetPath = targetFileLocation;                 // The path of the file that will launch when the shortcut is run
                shortcut.Arguments = @"https://www.aesoponline.com/login2.asp";
                shortcut.IconLocation = @"\\dataserver02\PCUpdate\54Helper\icons\aesop.ico";           // The icon of the shortcut
                shortcut.Save();
            }
            if (shortcutName == "E-Finance")
            {
                shortcut.TargetPath = targetFileLocation;                 // The path of the file that will launch when the shortcut is run
                shortcut.Arguments = @"https://efinance.sd54.org/gas2.50/wa/r/plus/finplus51/";
                shortcut.IconLocation = @"\\dataserver02\PCUpdate\54Helper\icons\efinance.ico";           // The icon of the shortcut
                shortcut.Save();
            }
            else {
                shortcut.TargetPath = targetFileLocation;                 // The path of the file that will launch when the shortcut is run
                shortcut.Save();                                    // Save the shortcut
            }

        }

        /// <summary>
        /// Handles the Click event of the installShortcutsButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void installShortcutsButton_Click(object sender, EventArgs e)
        {
            string homeDirectory = getHomeDirectory().ToString();
            string[] linkNames = {
                                     "InfiniteCampus",
                                     "AESOP",
                                     "E-Finance",
                                     "H-Drive",
                                     "Word",
                                 };

            foreach (string linkName in linkNames) 
            {
                switch (linkName)
                {
                    case "InfiniteCampus":
                        CreateShortcut("Infinite Campus", Environment.GetFolderPath(Environment.SpecialFolder.Desktop), @"C:\Program Files (x86)\Internet Explorer\iexplore.exe");
                        icShortcutCheckBox.Checked = true;
                        break;
                    case "AESOP":
                        CreateShortcut("AESOP", Environment.GetFolderPath(Environment.SpecialFolder.Desktop), @"C:\Program Files (x86)\Internet Explorer\iexplore.exe");
                        aesopShortcutCheckBox.Checked = true;
                        break;
                    case "E-Finance":
                        CreateShortcut("E-Finance", Environment.GetFolderPath(Environment.SpecialFolder.Desktop), @"C:\Program Files (x86)\Internet Explorer\iexplore.exe");
                        efinanceShortcutCheckBox.Checked = true;
                        break;
                    case "H-Drive":
                        CreateShortcut("H Drive", Environment.GetFolderPath(Environment.SpecialFolder.Desktop), homeDirectory);
                        homeShortcutCheckBox.Checked = true;
                        break;
                    case "Word":
                        string wordPath = getWordPath().ToString();
                        if (wordPath != "" || wordPath != null)
                        {
                            CreateShortcut("Microsoft Word", Environment.GetFolderPath(Environment.SpecialFolder.Desktop), wordPath);
                            wordShortcutCheckBox.Checked = true;
                        } else
                        {
                            MessageBox.Show("There is an issue with the installation of Word or it is not installed.");
                            wordShortcutCheckBox.Checked = false;
                        }
                        break;
                }
            }
        }

        /// <summary>
        /// Handles the SelectedIndexChanged event of the backupDriveCombo control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void backupDriveCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            using (BeginWaitCursorBlock())
            {
                //Set the selected drive freespace label
                DriveInfo selectedDrive = (DriveInfo)backupDriveCombo.SelectedItem;
                //If it's the homedrive (or a network drive) - we need to do further calculation to determine 'available free space'
                //Max Available For Standard Users: 3GB            
                if (selectedDrive.DriveType == DriveType.Network)
                {
                    long maxAvailableDriveSpace = 3221225472;
                    string folder = selectedDrive.ToString();
                    long networkDirectorySize = DirSize(new DirectoryInfo(folder));
                    selectedDriveAvailableSize = maxAvailableDriveSpace - networkDirectorySize;
                    string folderMB = FormatBytes(selectedDriveAvailableSize);
                    //folderMB = used space on network drive
                    backupDriveLabel.Text = folderMB + " Free";
                }
                //otherwise prepare as normal and show appropriate free space
                else
                {
                    if (selectedDrive.AvailableFreeSpace > 0)
                    {
                        selectedDriveAvailableSize = selectedDrive.AvailableFreeSpace;
                        string driveFreeSpace = FormatBytes(selectedDriveAvailableSize);
                        backupDriveLabel.Text = driveFreeSpace + " Free";
                    }
                }
                string userToBeBackedUp = userBackupSelectCombo.SelectedItem.ToString();
                labelDirectorySizes(userToBeBackedUp);
            }
        }
        /// <summary>
        /// Handles the SelectedIndexChanged event of the restoreDriveCombo control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void restoreDriveCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            using (BeginWaitCursorBlock())
            {
                //Set the selected drive freespace label
                disableRestoreButtons();
                DriveInfo selectedDrive = (DriveInfo)restoreDriveCombo.SelectedItem;
                setRestorePageOptions();
            }
        }

        /* ***** *
        * ***** * ***** * ***** *
        * ***** * ***** * ***** *
        * BACKUP BUTTONS SECTION
        * ***** * ***** * ***** *
        * ***** * ***** * ***** *
        * ***** */

        /// <summary>
        /// Handles the Click event of the backupDesktopButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void backupDesktopButton_Click(object sender, EventArgs e)
        {
            desktopBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            string selectedDrive = getSelectedDrive();
            string userToBeBackedUp = getUserToBeBackedUp();
            string userName = getUsername().ToString();
            Tuple<string, string, string> tuple = new Tuple<string, string, string>(userName, userToBeBackedUp, selectedDrive);
            essentialBgWorker.RunWorkerAsync(tuple);
        }

        /// <summary>
        /// Handles the Click event of the backupDocumentsButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void backupDocumentsButton_Click(object sender, EventArgs e)
        {
            documentsBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            string selectedDrive = getSelectedDrive();
            string userToBeBackedUp = getUserToBeBackedUp();
            string userName = getUsername().ToString();
            Tuple<string, string, string> tuple = new Tuple<string, string, string>(userName, userToBeBackedUp, selectedDrive);
            essentialBgWorker.RunWorkerAsync(tuple);
        }

        /// <summary>
        /// Handles the Click event of the backupFavoritesButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void backupFavoritesButton_Click(object sender, EventArgs e)
        {
            favoritesBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            string selectedDrive = getSelectedDrive();
            string userToBeBackedUp = getUserToBeBackedUp();
            string userName = getUsername().ToString();
            Tuple<string, string, string> tuple = new Tuple<string, string, string>(userName, userToBeBackedUp, selectedDrive);
            essentialBgWorker.RunWorkerAsync(tuple);
        }

        /// <summary>
        /// Handles the Click event of the backupAllEssentialsButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void backupAllEssentialsButton_Click(object sender, EventArgs e)
        {
            desktopBackupCheckBox.Checked = true;
            documentsBackupCheckBox.Checked = true;
            favoritesBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = @"Desktop, Documents and Favorites";
            string selectedDrive = getSelectedDrive();
            string userToBeBackedUp = getUserToBeBackedUp();
            string userName = getUsername().ToString();
            Tuple<string, string, string> tuple = new Tuple<string, string, string>(userName, userToBeBackedUp, selectedDrive);
            essentialBgWorker.RunWorkerAsync(tuple);
        }

        /// <summary>
        /// Handles the DoWork event of the essentialBgWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="DoWorkEventArgs"/> instance containing the event data.</param>
        public void essentialBgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            Tuple<string,string,string> tuple = e.Argument as Tuple<string,string,string>;
            string userName = tuple.Item1;
            string userToBeBackedUp = tuple.Item2;
            string selectedDrive = tuple.Item3;
            UserDesktop userDesktop = new UserDesktop();
            UserDocuments userDocuments = new UserDocuments();
            UserFavorites userFavorites = new UserFavorites();
            string buttonSender = clickedButton; //backupDesktopButton, backupDocumentsButton, etc.
            switch (buttonSender)
            {
                case "backupDesktopButton":         
                    userDesktop.backupDesktop(this, userName, selectedDrive, userToBeBackedUp, backupDirectoryName);
                   break;
                case "backupDocumentsButton":
                    userDocuments.backupDocuments(this, userName, selectedDrive, userToBeBackedUp, backupDirectoryName);
                    break;
                case "backupFavoritesButton":
                    userFavorites.backupFavorites(this, userName, selectedDrive, userToBeBackedUp, backupDirectoryName);
                    break;
                case "backupAllEssentialsButton":
                    userFavorites.backupFavorites(this, userName, selectedDrive, userToBeBackedUp, backupDirectoryName);
                    userDesktop.backupDesktop(this, userName, selectedDrive, userToBeBackedUp, backupDirectoryName);
                    userDocuments.backupDocuments(this, userName, selectedDrive, userToBeBackedUp, backupDirectoryName);
                    break;
                default:
                    //essentialBgWorker.CancelAsync();
                    break;
            }
        }

        /// <summary>
        /// Handles the ProgressChanged event of the essentialBgWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="ProgressChangedEventArgs"/> instance containing the event data.</param>
        void essentialBgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            essentialItemsProgressBar.Visible = true;
            essentialItemsProgressBar.Value = e.ProgressPercentage;
            //disable the panel
            foreach (Control cont in dataBackupGroupBox.Controls)
            {
                if (cont.Name.ToString() == "essentialItemsProgressBar" || cont.Name.ToString() == "essentialItemsProgressLabel")
                {

                }
                else
                {
                    cont.Enabled = false;
                }
            }
            int percent = (int)(((double)(essentialItemsProgressBar.Value - essentialItemsProgressBar.Minimum) /
            (double)(essentialItemsProgressBar.Maximum - essentialItemsProgressBar.Minimum)) * 100);
            using (Graphics gr = essentialItemsProgressBar.CreateGraphics())
            {
                gr.DrawString(percent.ToString() + "%",
                    SystemFonts.DefaultFont,
                    Brushes.Black,
                    new PointF(essentialItemsProgressBar.Width / 2 - (gr.MeasureString(percent.ToString() + "%",
                        SystemFonts.DefaultFont).Width / 2.0F),
                    essentialItemsProgressBar.Height / 2 - (gr.MeasureString(percent.ToString() + "%",
                        SystemFonts.DefaultFont).Height / 2.0F)));
            }

            essentialItemsProgressLabel.Visible = true;
            //essentialItemsProgressLabel.Text = String.Format("Progress: {0} % - All {1} Files Transferred", e.ProgressPercentage, clickedButton);
            //essentialItemsProgressLabel.Text = String.Format("Total items transfered: {0}", e.UserState);
        }

        /// <summary>
        /// Handles the RunWorkerCompleted event of the essentialBgWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RunWorkerCompletedEventArgs"/> instance containing the event data.</param>
        void essentialBgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //do the code when bgv completes its work           
            essentialItemsProgressBar.Visible = false;
            essentialItemsProgressLabel.Visible = true;
            essentialItemsProgressLabel.Text = String.Format("Progress: 100% - All {0} Files Transferred", itemsChanged);
            MessageBox.Show("Process Complete. Always verify all data was backed up appropriately.");
            //enable the panel
            foreach (Control cont in dataBackupGroupBox.Controls)
            {
                if (cont.Name.ToString() == "essentialItemsProgressBar" || cont.Name.ToString() == "essentialItemsProgressLabel")
                {

                }
                else
                {
                    cont.Enabled = true;
                }
            }
            setRestorePageOptions();
        }

        /// <summary>
        /// Handles the Click event of the backupStickyNotesButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void backupStickyNotesButton_Click(object sender, EventArgs e)
        {
            stickyNotesBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            string selectedDrive = getSelectedDrive();
            string userToBeBackedUp = getUserToBeBackedUp();
            string userName = getUsername().ToString();
            Tuple<string, string, string> tuple = new Tuple<string, string, string>(userName, userToBeBackedUp, selectedDrive);
            additionalBgWorker.RunWorkerAsync(tuple);
        }

        /// <summary>
        /// Handles the Click event of the backupPicturesButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void backupPicturesButton_Click(object sender, EventArgs e)
        {
            picturesBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            string selectedDrive = getSelectedDrive();
            string userToBeBackedUp = getUserToBeBackedUp();
            string userName = getUsername().ToString();
            Tuple<string, string, string> tuple = new Tuple<string, string, string>(userName, userToBeBackedUp, selectedDrive);
            additionalBgWorker.RunWorkerAsync(tuple);
        }

        /// <summary>
        /// Handles the Click event of the backupVideosButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void backupVideosButton_Click(object sender, EventArgs e)
        {
            videosBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            string selectedDrive = getSelectedDrive();
            string userToBeBackedUp = getUserToBeBackedUp();
            string userName = getUsername().ToString();
            Tuple<string, string, string> tuple = new Tuple<string, string, string>(userName, userToBeBackedUp, selectedDrive);
            additionalBgWorker.RunWorkerAsync(tuple);
        }

        /// <summary>
        /// Handles the Click event of the backupMusicButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void backupMusicButton_Click(object sender, EventArgs e)
        {
            musicBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            string selectedDrive = getSelectedDrive();
            string userToBeBackedUp = getUserToBeBackedUp();
            string userName = getUsername().ToString();
            Tuple<string, string, string> tuple = new Tuple<string, string, string>(userName, userToBeBackedUp, selectedDrive);
            additionalBgWorker.RunWorkerAsync(tuple);
        }

        /// <summary>
        /// Handles the DoWork event of the additionalBgWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="DoWorkEventArgs"/> instance containing the event data.</param>
        void additionalBgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            Tuple<string, string, string> tuple = e.Argument as Tuple<string, string, string>;
            string userName = tuple.Item1;
            string userToBeBackedUp = tuple.Item2;
            string selectedDrive = tuple.Item3;
            UserPictures userPictures = new UserPictures();
            UserStickyNotes userStickyNotes = new UserStickyNotes();
            UserVideos userVideos = new UserVideos();
            UserMusic userMusic = new UserMusic();
            string buttonSender = clickedButton; //Desktop, Documents, etc.
            switch (buttonSender)
            {
                case "backupStickyNotesButton":
                    userStickyNotes.backupStickyNotes(this, userName, selectedDrive, userToBeBackedUp, backupDirectoryName);
                    break;
                case "backupPicturesButton":
                    userPictures.backupPictures(this, userName, selectedDrive, userToBeBackedUp, backupDirectoryName);
                    break;
                case "backupVideosButton":
                    userVideos.backupVideos(this, userName, selectedDrive, userToBeBackedUp, backupDirectoryName);
                    break;
                case "backupMusicButton":
                    userMusic.backupMusic(this, userName, selectedDrive, userToBeBackedUp, backupDirectoryName);
                    break;
                default:

                    break;
            }
        }

        /// <summary>
        /// Handles the ProgressChanged event of the additionalBgWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="ProgressChangedEventArgs"/> instance containing the event data.</param>
        void additionalBgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            additionalItemsProgressBar.Visible = true;
            additionalItemsProgressBar.Value = e.ProgressPercentage;
            //disable the panel
            foreach (Control cont in dataBackupGroupBox.Controls)
            {
                if (cont.Name.ToString() == "additionalItemsProgressBar" || cont.Name.ToString() == "additionalItemsProgressLabel")
                {

                }
                else
                {
                    cont.Enabled = false;
                }
            }
            int percent = (int)(((double)(additionalItemsProgressBar.Value - additionalItemsProgressBar.Minimum) /
            (double)(additionalItemsProgressBar.Maximum - additionalItemsProgressBar.Minimum)) * 100);
            using (Graphics gr = additionalItemsProgressBar.CreateGraphics())
            {
                gr.DrawString(percent.ToString() + "%",
                    SystemFonts.DefaultFont,
                    Brushes.Black,
                    new PointF(additionalItemsProgressBar.Width / 2 - (gr.MeasureString(percent.ToString() + "%",
                        SystemFonts.DefaultFont).Width / 2.0F),
                    additionalItemsProgressBar.Height / 2 - (gr.MeasureString(percent.ToString() + "%",
                        SystemFonts.DefaultFont).Height / 2.0F)));
            }

            additionalItemsProgressLabel.Visible = true;
            //additionalItemsProgressLabel.Text = String.Format("Progress: {0} % - All {1} Files Transferred", e.ProgressPercentage, clickedButton);
            //additionalItemsProgressLabel.Text = String.Format("Total items transfered: {0}", e.UserState);
        }

        /// <summary>
        /// Handles the RunWorkerCompleted event of the additionalBgWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RunWorkerCompletedEventArgs"/> instance containing the event data.</param>
        void additionalBgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //do the code when bgv completes its work
            additionalItemsProgressBar.Visible = false;
            additionalItemsProgressLabel.Visible = true;
            additionalItemsProgressLabel.Text = String.Format("Progress: 100% - All {0} Files Transferred", itemsChanged);
            MessageBox.Show("Process complete. Always verify all data was backed up appropriately.");
            //enable the panel
            foreach (Control cont in dataBackupGroupBox.Controls)
            {
                if (cont.Name.ToString() == "additionalItemsProgressBar" || cont.Name.ToString() == "additionalItemsProgressLabel")
                {

                }
                else
                {
                    cont.Enabled = true;
                }
            }
            setRestorePageOptions();
        }

        /// <summary>
        /// Labels the directory sizes.
        /// </summary>
        public void labelDirectorySizes(string userToBeBackedUp)
        {
            string userName = getUsername().ToString();
            string userDirectoryLocation;
            string[] directoryLocations =
                {
                    "DesktopDirectory",
                    "MyDocuments",
                    "Favorites",
                    "MyMusic",
                    "MyPictures",
                    "My Videos"
                };

            if (userToBeBackedUp != userName)
            {
                userDirectoryLocation = "C:\\Users\\" + userToBeBackedUp;
                //Set the size & label for the button that backs up Desktop, Documents & Favorites
                var desktopFolder = userDirectoryLocation + "\\Desktop"; 
                long desktopSize = DirSize(new DirectoryInfo(desktopFolder));
                var documentsFolder = userDirectoryLocation + "\\Documents"; 
                long documentsSize = DirSize(new DirectoryInfo(documentsFolder));
                var favoritesFolder = userDirectoryLocation + "\\Favorites";
                long favoritesSize = DirSize(new DirectoryInfo(favoritesFolder));
                long totalSize = desktopSize + documentsSize + favoritesSize;
                string totalMB = FormatBytes(totalSize);
                allEssentialsSizeLabel.Text = totalMB;
            }
            else
            {
                userDirectoryLocation = Environment.GetEnvironmentVariable("userprofile");
                //Set the size & label for the button that backs up Desktop, Documents & Favorites
                var desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                long desktopSize = DirSize(new DirectoryInfo(desktopFolder));
                var documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                long documentsSize = DirSize(new DirectoryInfo(documentsFolder));
                var favoritesFolder = Environment.GetFolderPath(Environment.SpecialFolder.Favorites);
                long favoritesSize = DirSize(new DirectoryInfo(favoritesFolder));
                long totalSize = desktopSize + documentsSize + favoritesSize;
                string totalMB = FormatBytes(totalSize);
                allEssentialsSizeLabel.Text = totalMB;
            }
            // Get the directory sizes for each directoryLocation & set the label
            foreach (string directoryLocation in directoryLocations)
            {
                string folder;
                //'My Videos' is not supported in older frameworks so set it seperately
                if (directoryLocation == "My Videos")
                {
                    string folderName = "Videos";
                    folder = userDirectoryLocation + "\\" + folderName;
                    long folderSize = DirSize(new DirectoryInfo(folder));
                    canItFit = selectedDriveAvailableSize - folderSize;
                    string folderMB = FormatBytes(folderSize);
                    videosSizeLabel.Text = folderMB;
                    if (canItFit < 0)
                    {
                        videosSizeLabel.ForeColor = System.Drawing.Color.Red;
                    }
                    else
                    {
                        videosSizeLabel.ForeColor = System.Drawing.Color.ForestGreen;
                        backupVideosButton.Enabled = true;
                    }
                }
                else //iterate through all the other directory Locations
                {
                    switch (directoryLocation)
                    {
                        case "DesktopDirectory":
                            long selectedDriveSize, folderSize;
                            string folderMB;
                            if (userToBeBackedUp != userName)
                            {
                                //set to the other users directory
                                var dir = userDirectoryLocation + "\\Desktop";
                                folderSize = DirSize(new DirectoryInfo(dir));
                                folderMB = FormatBytes(folderSize);
                                selectedDriveSize = selectedDriveAvailableSize;
                            }
                            else
                            {
                                var dir = (Environment.SpecialFolder)Enum.Parse(typeof(Environment.SpecialFolder), directoryLocation);
                                folder = Environment.GetFolderPath(dir);
                                folderSize = DirSize(new DirectoryInfo(folder));
                                folderMB = FormatBytes(folderSize);
                                selectedDriveSize = selectedDriveAvailableSize;
                            }
                            canItFit = selectedDriveSize - folderSize;
                            desktopSizeLabel.Text = folderMB;
                            if (canItFit < 0)
                            {
                                desktopSizeLabel.ForeColor = System.Drawing.Color.Red;
                                backupDesktopButton.Enabled = false;
                                backupAllEssentialsButton.Enabled = false;
                            }
                            else
                            {
                                desktopSizeLabel.ForeColor = System.Drawing.Color.ForestGreen;
                                backupDesktopButton.Enabled = true;
                            }
                            break;
                        case "MyDocuments":
                            if (userToBeBackedUp != userName)
                            {
                                //set to the other users directory
                                var dir = userDirectoryLocation + "\\Documents";
                                folderSize = DirSize(new DirectoryInfo(dir));
                                folderMB = FormatBytes(folderSize);
                                selectedDriveSize = selectedDriveAvailableSize;
                            }
                            else
                            {
                                var dir = (Environment.SpecialFolder)Enum.Parse(typeof(Environment.SpecialFolder), directoryLocation);
                                folder = Environment.GetFolderPath(dir);
                                folderSize = DirSize(new DirectoryInfo(folder));
                                folderMB = FormatBytes(folderSize);
                                selectedDriveSize = selectedDriveAvailableSize;
                            }
                            canItFit = selectedDriveSize - folderSize;
                            documentsSizeLabel.Text = folderMB;
                            if (canItFit < 0)
                            {
                                documentsSizeLabel.ForeColor = System.Drawing.Color.Red;
                                backupDocumentsButton.Enabled = false;
                                backupAllEssentialsButton.Enabled = false;
                            }
                            else
                            {
                                documentsSizeLabel.ForeColor = System.Drawing.Color.ForestGreen;
                                backupDocumentsButton.Enabled = true;
                            }
                            break;
                        case "Favorites":
                            if (userToBeBackedUp != userName)
                            {
                                //set to the other users directory
                                var dir = userDirectoryLocation + "\\Favorites";
                                folderSize = DirSize(new DirectoryInfo(dir));
                                folderMB = FormatBytes(folderSize);
                                selectedDriveSize = selectedDriveAvailableSize;
                            }
                            else
                            {
                                var dir = (Environment.SpecialFolder)Enum.Parse(typeof(Environment.SpecialFolder), directoryLocation);
                                folder = Environment.GetFolderPath(dir);
                                folderSize = DirSize(new DirectoryInfo(folder));
                                folderMB = FormatBytes(folderSize);
                                selectedDriveSize = selectedDriveAvailableSize;
                            }
                            canItFit = selectedDriveSize - folderSize;
                            favoritesSizeLabel.Text = folderMB;
                            if (canItFit < 0)
                            {
                                favoritesSizeLabel.ForeColor = System.Drawing.Color.Red;
                                backupFavoritesButton.Enabled = false;
                                backupAllEssentialsButton.Enabled = false;
                            }
                            else
                            {
                                favoritesSizeLabel.ForeColor = System.Drawing.Color.ForestGreen;
                                backupFavoritesButton.Enabled = true;
                            }
                            break;
                        case "MyMusic":
                            if (userToBeBackedUp != userName)
                            {
                                //set to the other users directory
                                var dir = userDirectoryLocation + "\\Music";
                                folderSize = DirSize(new DirectoryInfo(dir));
                                folderMB = FormatBytes(folderSize);
                                selectedDriveSize = selectedDriveAvailableSize;
                            }
                            else
                            {
                                var dir = (Environment.SpecialFolder)Enum.Parse(typeof(Environment.SpecialFolder), directoryLocation);
                                folder = Environment.GetFolderPath(dir);
                                folderSize = DirSize(new DirectoryInfo(folder));
                                folderMB = FormatBytes(folderSize);
                                selectedDriveSize = selectedDriveAvailableSize;
                            }
                            canItFit = selectedDriveSize - folderSize;
                            musicSizeLabel.Text = folderMB;
                            if (canItFit < 0)
                            {
                                musicSizeLabel.ForeColor = System.Drawing.Color.Red;
                                backupMusicButton.Enabled = false;
                            }
                            else
                            {
                                musicSizeLabel.ForeColor = System.Drawing.Color.ForestGreen;
                                backupMusicButton.Enabled = true;
                            }
                            break;
                        case "MyPictures":
                            if (userToBeBackedUp != userName)
                            {
                                //set to the other users directory
                                var dir = userDirectoryLocation + "\\Pictures";
                                folderSize = DirSize(new DirectoryInfo(dir));
                                folderMB = FormatBytes(folderSize);
                                selectedDriveSize = selectedDriveAvailableSize;
                            }
                            else
                            {
                                var dir = (Environment.SpecialFolder)Enum.Parse(typeof(Environment.SpecialFolder), directoryLocation);
                                folder = Environment.GetFolderPath(dir);
                                folderSize = DirSize(new DirectoryInfo(folder));
                                folderMB = FormatBytes(folderSize);
                                selectedDriveSize = selectedDriveAvailableSize;
                            }
                            canItFit = selectedDriveSize - folderSize;
                            picturesSizeLabel.Text = folderMB;
                            if (canItFit < 0)
                            {
                                picturesSizeLabel.ForeColor = System.Drawing.Color.Red;
                                backupPicturesButton.Enabled = false;
                            }
                            else
                            {
                                picturesSizeLabel.ForeColor = System.Drawing.Color.ForestGreen;
                                backupPicturesButton.Enabled = true;
                            }
                            break;
                    }
                    if (backupDesktopButton.Enabled && backupDocumentsButton.Enabled && backupFavoritesButton.Enabled)
                    {
                        allEssentialsSizeLabel.ForeColor = System.Drawing.Color.ForestGreen;
                        backupAllEssentialsButton.Enabled = true;
                    }
                }

            }
            //Set the size & label for Sticky Notes
            string stickyNotesFolder;
            if (WinMajorVersion == 10)
            //windows 10
            {
                stickyNotesFolder = "C:\\Users\\" + userToBeBackedUp + "\\AppData\\Local\\Packages\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe";
            }
            else 
            {
            //windows 7
                stickyNotesFolder = "C:\\Users\\" + userToBeBackedUp + "\\AppData\\Roaming\\Microsoft\\Sticky Notes";
            }
            if (Directory.Exists(stickyNotesFolder)) //If they have launched sticky notes
            {
                long stickyNotesSize = DirSize(new DirectoryInfo(stickyNotesFolder));
                canItFit = selectedDriveAvailableSize - stickyNotesSize;
                if (canItFit < 0)
                {
                    stickyNotesSizeLabel.ForeColor = System.Drawing.Color.Red;
                    backupStickyNotesButton.Enabled = false;
                }
                else
                {
                    stickyNotesSizeLabel.ForeColor = System.Drawing.Color.ForestGreen;
                    backupStickyNotesButton.Enabled = true;
                }
                string stickyNotesMB = FormatBytes(stickyNotesSize);
                stickyNotesSizeLabel.Text = stickyNotesMB;
            }
            else //haven't launched sticky notes
            {
                stickyNotesSizeLabel.Text = "N/A";
            }
        }

        /// <summary>
        /// Copies the files recursively.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="target">The target.</param>
        public void CopyFilesRecursively(DirectoryInfo source, DirectoryInfo target)
                {
                    foreach (DirectoryInfo dir in source.GetDirectories())
                    {
                        if ((dir.Name == "My Music") || (dir.Name == "My Pictures") || (dir.Name == "My Videos"))
                        {
                            //do nothing with it since it is just a system link we cannot access
                        }
                        else
                        {
                            CopyFilesRecursively(dir, target.CreateSubdirectory(dir.Name));
                        }
                    }
                    foreach (FileInfo file in source.GetFiles())
                        file.CopyTo(Path.Combine(target.FullName, file.Name), true); //overwrites the existing files with newer
                }

        /* ***** *
        * ***** * ***** * ***** *
        * ***** * ***** * ***** *
        * RESTORE BUTTONS SECTION
        * ***** * ***** * ***** *
        * ***** * ***** * ***** *
        * ***** */

        /// <summary>
        /// Handles the Click event of the restoreDesktopButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void restoreDesktopButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();                        
            restoreEssentialBgWorker.RunWorkerAsync();
        }

        /// <summary>
        /// Handles the Click event of the restoreDocumentsButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void restoreDocumentsButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            restoreEssentialBgWorker.RunWorkerAsync();
        }

        /// <summary>
        /// Handles the Click event of the restoreFavoritesButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void restoreFavoritesButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            restoreEssentialBgWorker.RunWorkerAsync();
        }

        /// <summary>
        /// Handles the Click event of the restoreAllEssentialsButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void restoreAllEssentialsButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            string selectedDrive = backupDriveCombo.SelectedItem.ToString();
            restoreEssentialBgWorker.RunWorkerAsync();
            string[] directoryNameArray = { "Documents", "Favorites", "Desktop" };
            string destinationLocation = "";
            string directoryToRestore = "";

            for (int i = 0; i < directoryNameArray.Length; i++)
            {
                switch (directoryNameArray[i])
                {
                    case "Documents":
                        if (isRecoveryForLoggedInUser())
                        {
                            destinationLocation = selectedDrive + backupDirectoryName + "\\" + directoryNameArray[i] + "\\";
                            directoryToRestore = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                            break;
                        }
                        else
                        {
                            string selectedBackup = userRestoreSelectCombo.SelectedItem.ToString();
                            string selectedBackupUsername = selectedBackup.Split('-')[0];
                            destinationLocation = selectedDrive + backupDirectoryName + "\\" + directoryNameArray[i] + "\\";
                            directoryToRestore = "C:\\Users\\"+selectedBackupUsername+"\\Documents";
                            break;
                        }
                    case "Favorites":
                        if (isRecoveryForLoggedInUser())
                        {
                            destinationLocation = selectedDrive + backupDirectoryName + "\\" + directoryNameArray[i] + "\\";
                            directoryToRestore = Environment.GetFolderPath(Environment.SpecialFolder.Favorites);
                            break;
                        }
                        else
                        {
                            string selectedBackup = userRestoreSelectCombo.SelectedItem.ToString();
                            string selectedBackupUsername = selectedBackup.Split('-')[0];
                            destinationLocation = selectedDrive + backupDirectoryName + "\\" + directoryNameArray[i] + "\\";
                            directoryToRestore = "C:\\Users\\" + selectedBackupUsername + "\\Favorites";
                            break;
                        }
                
                    case "Desktop":
                        if (isRecoveryForLoggedInUser())
                        {
                            destinationLocation = selectedDrive + backupDirectoryName + "\\" + directoryNameArray[i] + "\\";
                            directoryToRestore = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                            break;
                        }
                        else
                        {
                            string selectedBackup = userRestoreSelectCombo.SelectedItem.ToString();
                            string selectedBackupUsername = selectedBackup.Split('-')[0];
                            destinationLocation = selectedDrive + backupDirectoryName + "\\" + directoryNameArray[i] + "\\";
                            directoryToRestore = "C:\\Users\\" + selectedBackupUsername + "\\Desktop";
                            break;
                        }
                    default:
                        break;
                }
                if (!Directory.Exists(destinationLocation))
                {
                    Directory.CreateDirectory(destinationLocation);
                }
                DirectoryInfo target = new DirectoryInfo(destinationLocation);
                DirectoryInfo source = new DirectoryInfo(directoryToRestore);
                CopyFilesRecursively(source, target);
            }
        }

        /// <summary>
        /// Handles the DoWork event of the restoreEssentialBgWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="DoWorkEventArgs"/> instance containing the event data.</param>
        void restoreEssentialBgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            UserDocuments userDocuments = new UserDocuments();
            UserDesktop userDesktop = new UserDesktop();
            UserFavorites userFavorites = new UserFavorites();
            string selectedDrive = restoreDriveCombo.SelectedItem.ToString();
            string buttonSender = clickedButton; //Desktop, Documents, etc.
            switch (buttonSender)
            {
                case "restoreDesktopButton":
                    userDesktop.restoreDesktop(this, selectedDrive, backupToRestore);
                    break;
                case "restoreDocumentsButton":
                    userDocuments.restoreDocuments(this, selectedDrive, backupToRestore);
                    break;
                case "restoreFavoritesButton":
                    userFavorites.restoreFavorites(this, selectedDrive, backupToRestore);
                    break;
                case "restoreAllEssentialsButton":
                    userFavorites.restoreFavorites(this, selectedDrive, backupToRestore);
                    userDesktop.restoreDesktop(this, selectedDrive, backupToRestore);
                    userDocuments.restoreDocuments(this, selectedDrive, backupToRestore);
                    break;
                default:

                    break;
            }

        }

        /// <summary>
        /// Handles the ProgressChanged event of the restoreEssentialBgWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="ProgressChangedEventArgs"/> instance containing the event data.</param>
        void restoreEssentialBgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            restoreEssentialsProgressBar.Visible = true;
            restoreEssentialsProgressBar.Value = e.ProgressPercentage;
            //disable the panel
            foreach (Control cont in restoreGroupBox.Controls)
            {
                if (cont.Name.ToString() == "restoreEssentialsProgressBar" || cont.Name.ToString() == "restoreEssentialsBarLabel")
                {

                }
                else
                {
                    cont.Enabled = false;
                }
            }
            int percent = (int)(((double)(restoreEssentialsProgressBar.Value - restoreEssentialsProgressBar.Minimum) /
            (double)(restoreEssentialsProgressBar.Maximum - restoreEssentialsProgressBar.Minimum)) * 100);
            using (Graphics gr = restoreEssentialsProgressBar.CreateGraphics())
            {
                gr.DrawString(percent.ToString() + "%",
                    SystemFonts.DefaultFont,
                    Brushes.Black,
                    new PointF(restoreEssentialsProgressBar.Width / 2 - (gr.MeasureString(percent.ToString() + "%",
                        SystemFonts.DefaultFont).Width / 2.0F),
                    restoreEssentialsProgressBar.Height / 2 - (gr.MeasureString(percent.ToString() + "%",
                        SystemFonts.DefaultFont).Height / 2.0F)));
            }

            restoreEssentialsBarLabel.Visible = true;
            //essentialItemsProgressLabel.Text = String.Format("Progress: {0} % - All {1} Files Transferred", e.ProgressPercentage, clickedButton);
            //essentialItemsProgressLabel.Text = String.Format("Total items transfered: {0}", e.UserState);
        }

        /// <summary>
        /// Handles the RunWorkerCompleted event of the restoreEssentialBgWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RunWorkerCompletedEventArgs"/> instance containing the event data.</param>
        void restoreEssentialBgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //do the code when bgv completes its work            
            restoreEssentialsProgressBar.Visible = false;
            restoreEssentialsBarLabel.Visible = true;
            restoreEssentialsBarLabel.Text = String.Format("Progress: 100% - All {0} Files Restored", itemsChanged);
            MessageBox.Show("Process Complete. Always verify all data was restored appropriately.");
            //enable the panel
            foreach (Control cont in restoreGroupBox.Controls)
            {
                if (cont.Name.ToString() == "restoreEssentialsProgressBar" || cont.Name.ToString() == "restoreEssentialsProgressLabel")
                {

                }
                else
                {
                    cont.Enabled = true;
                }
            }
            
        }

        /// <summary>
        /// Handles the Click event of the restoreStickyNotesButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void restoreStickyNotesButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            restoreAdditionalBgWorker.RunWorkerAsync();
            restoreAdditionalBarLabel.Visible = false;
        }

        /// <summary>
        /// Handles the Click event of the restorePicturesButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void restorePicturesButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            restoreAdditionalBgWorker.RunWorkerAsync();
            restoreAdditionalBarLabel.Visible = false;
        }

        /// <summary>
        /// Handles the Click event of the restoreVideosButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void restoreVideosButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            restoreAdditionalBgWorker.RunWorkerAsync();
            restoreAdditionalBarLabel.Visible = false;
        }

        /// <summary>
        /// Handles the Click event of the restoreMusicButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void restoreMusicButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            restoreAdditionalBgWorker.RunWorkerAsync();
            restoreAdditionalBarLabel.Visible = false;
        }

        /// <summary>
        /// Handles the DoWork event of the restoreAdditionalBgWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="DoWorkEventArgs"/> instance containing the event data.</param>
        void restoreAdditionalBgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            UserPictures userPictures = new UserPictures();
            UserStickyNotes userStickyNotes = new UserStickyNotes();
            UserVideos userVideos = new UserVideos();
            UserMusic userMusic = new UserMusic();
            string selectedDrive = restoreDriveCombo.SelectedItem.ToString();
            string buttonSender = clickedButton; //Desktop, Documents, etc.
            switch (buttonSender)
            {
                case "restoreStickyNotesButton":
                    userStickyNotes.restoreStickyNotes(this, selectedDrive, backupToRestore);
                    break;
                case "restorePicturesButton":
                    userPictures.restorePictures(this, selectedDrive, backupToRestore);
                    break;
                case "restoreVideosButton":
                    userVideos.restoreVideos(this, selectedDrive, backupToRestore);
                    break;
                case "restoreMusicButton":
                    userMusic.restoreMusic(this, selectedDrive, backupToRestore);
                    break;
                default:

                    break;
            }
            //set the cancel flag so we know the job was cancelled!
            if (restoreAdditionalBgWorker.CancellationPending)
            {
                e.Cancel = true;
                return;
            }
        }

        /// <summary>
        /// Handles the ProgressChanged event of the restoreAdditionalBgWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="ProgressChangedEventArgs"/> instance containing the event data.</param>
        void restoreAdditionalBgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            restoreAdditionalProgressBar.Visible = true;
            restoreAdditionalProgressBar.Value = e.ProgressPercentage;
            //disable the panel while restoring
            foreach (Control cont in restoreGroupBox.Controls)
            {
                if (cont.Name.ToString() == "restoreAdditionalProgressBar" || cont.Name.ToString() == "restoreAdditionalBarLabel")
                {

                }
                else
                {
                    cont.Enabled = false;
                }
            }
            int percent = (int)(((double)(restoreAdditionalProgressBar.Value - restoreAdditionalProgressBar.Minimum) /
            (double)(restoreAdditionalProgressBar.Maximum - restoreAdditionalProgressBar.Minimum)) * 100);
            using (Graphics gr = restoreAdditionalProgressBar.CreateGraphics())
            {
                gr.DrawString(percent.ToString() + "%",
                    SystemFonts.DefaultFont,
                    Brushes.Black,
                    new PointF(restoreAdditionalProgressBar.Width / 2 - (gr.MeasureString(percent.ToString() + "%",
                        SystemFonts.DefaultFont).Width / 2.0F),
                    restoreAdditionalProgressBar.Height / 2 - (gr.MeasureString(percent.ToString() + "%",
                        SystemFonts.DefaultFont).Height / 2.0F)));
            }

            //restoreAdditionalBarLabel.Visible = true;
            //essentialItemsProgressLabel.Text = String.Format("Progress: {0} % - All {1} Files Transferred", e.ProgressPercentage, clickedButton);
            //essentialItemsProgressLabel.Text = String.Format("Total items transfered: {0}", e.UserState);
        }

        /// <summary>
        /// Handles the RunWorkerCompleted event of the restoreAdditionalBgWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RunWorkerCompletedEventArgs"/> instance containing the event data.</param>
        void restoreAdditionalBgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // First, handle the case where an exception was thrown.
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            else if (e.Cancelled)
            {
                // Next, handle the case where the user canceled 
                // the operation.
                // Note that due to a race condition in 
                // the DoWork event handler, the Cancelled
                // flag may not have been set, even though
                // CancelAsync was called.
                restoreAdditionalBarLabel.Visible = false;
                MessageBox.Show("Restore Process Cancelled");
            }
            else
            {
                //do the code when bgv completes its work                
                restoreAdditionalProgressBar.Visible = false;
                restoreAdditionalBarLabel.ForeColor = System.Drawing.Color.ForestGreen;
                restoreAdditionalBarLabel.Text = String.Format("Progress: 100% - All {0} Files Restored", itemsChanged);
                restoreAdditionalBarLabel.Visible = true;
                MessageBox.Show("Process Complete. Always verify all data was restored appropriately.");
                //enable the panel
                foreach (Control cont in restoreGroupBox.Controls)
                {
                    if (cont.Name.ToString() == "restoreAdditionalProgressBar" || cont.Name.ToString() == "restoreAdditionalProgressLabel")
                    {

                    }
                    else
                    {
                        cont.Enabled = true;
                    }
                }
            }


        }

        /// <summary>
        /// Restores the files recursively.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="target">The target.</param>
        public void RestoreFilesRecursively(DirectoryInfo source, DirectoryInfo target)
        {
            foreach (DirectoryInfo dir in source.GetDirectories())
            {
                if ((dir.Name == "My Music") || (dir.Name == "My Pictures") || (dir.Name == "My Videos"))
                {
                    //do nothing with it since it is just a system link we cannot access
                }
                else
                {
                    RestoreFilesRecursively(dir, target.CreateSubdirectory(dir.Name));
                }
            }
            foreach (FileInfo file in source.GetFiles())
                file.CopyTo(Path.Combine(target.FullName, file.Name), true); //overwrites the existing files with newer
        }

        /// <summary>
        /// Determines whether the application is installed via the registry.
        /// </summary>
        /// <param name="p_name">Name of the program.</param>
        /// <returns>
        ///   Bool
        /// </returns>
        public static bool IsApplicationInstalled(string p_name)
        {
            string keyName;

            // search in: CurrentUser
            keyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            if (ExistsInSubKey(Registry.CurrentUser, keyName, "DisplayName", p_name) == true)
            {
                return true;
            }

            // search in: LocalMachine_32
            keyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
            if (ExistsInSubKey(Registry.LocalMachine, keyName, "DisplayName", p_name) == true)
            {
                return true;
            }

            // search in: LocalMachine_64
            keyName = @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall";
            if (ExistsInSubKey(Registry.LocalMachine, keyName, "DisplayName", p_name) == true)
            {
                return true;
            }

            // search in: LocalMachine_64
            keyName = @"SOFTWARE\Classes\Installer\Products";
            if (ExistsInSubKey(Registry.LocalMachine, keyName, "ProductName", p_name) == true)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Check registry to see if program exists via MSI Installer.
        /// </summary>
        /// <param name="p_root">RegistryKey Root key</param>
        /// <param name="p_subKeyName">String Name of the program sub key.</param>
        /// <param name="p_attributeName">String Name of the program attribute.</param>
        /// <param name="p_name">String Name of the program.</param>
        /// <returns></returns>
        private static bool ExistsInSubKey(RegistryKey p_root, string p_subKeyName, string p_attributeName, string p_name)
        {
            RegistryKey subkey;
            string displayName;

            using (RegistryKey key = p_root.OpenSubKey(p_subKeyName))
            {
                if (key != null)
                {
                    foreach (string kn in key.GetSubKeyNames())
                    {
                        using (subkey = key.OpenSubKey(kn))
                        {
                            displayName = subkey.GetValue(p_attributeName) as string;
                            if (p_name.Equals(displayName, StringComparison.OrdinalIgnoreCase) == true)
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Handles the Click event of the outlookButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void outlookButton_Click(object sender, EventArgs e)
        {
            Button sentButton = sender as Button;
            if (sentButton.Text.ToString() == "Open Outlook")
            {
                openSoftwareClick(outlookButton, null);
            }
            else
            {
                installSoftwareClick(outlookButton, null);
            }
            outlookCheckBox.Checked = true;
        }

        /// <summary>
        /// Handles the Click event of the dymoButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void dymoButton_Click(object sender, EventArgs e)
        {
            Button sentButton = sender as Button;
            if (sentButton.Text.ToString() == "Open Dymo")
            {
                openSoftwareClick(dymoButton, null);
            }
            else
            {
                installSoftwareClick(dymoButton, null);
            }
        }

        /// <summary>
        /// Handles the Click event of the acrobatButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void acrobatButton_Click(object sender, EventArgs e)
        {
            Button sentButton = sender as Button;
            if (sentButton.Text.ToString() == "Open Acrobat Pro")
            {
                openSoftwareClick(acrobatButton, null);
            }
            else
            {
                installSoftwareClick(acrobatButton, null);
            }
            adobeProCheckBox.Checked = true;
        }

        /// <summary>
        /// Handles the Click event of the scanSnapButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void scanSnapButton_Click(object sender, EventArgs e)
        {
            Button sentButton = sender as Button;
            if (sentButton.Text.ToString() == "Open ScanSnap")
            {
                openSoftwareClick(scanSnapButton, null);
            }
            else
            {
                installSoftwareClick(scanSnapButton, null);
            }
        }

        /// <summary>
        /// Handles the Click event of the quickenButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void quickenButton_Click(object sender, EventArgs e)
        {
            Button sentButton = sender as Button;
            if (sentButton.Text.ToString() == "Open Quicken")
            {
                openSoftwareClick(quickenButton, null);
            }
            else
            {
                installSoftwareClick(quickenButton, null);
            }
            quickenCheckBox.Checked = true;
        }

        /// <summary>
        /// Handles the Click event of the installPrintersButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void installPrintersButton_Click(object sender, EventArgs e)
        {
            string homeDirectory = getHomeDirectory().ToString();
            DriveInfo[] theDrives = DriveInfo.GetDrives();
            foreach (DriveInfo currentDrive in theDrives)
            {
                if (currentDrive.Name == homeDirectory)
                {
                    string currentDriveString = currentDrive.Name.ToString();
                    string path = GetUNCPath(currentDriveString);
                    Uri uri = new Uri(path);
                    string serverName = uri.Host.ToString();
                    System.Diagnostics.Process.Start("explorer", @"\\" + serverName);
                }
            }
            installPrintersCheckBox.Checked = true;
            imageRunnerCheckBox.Checked = true;
        }

        /*
        * ***** *
        * ***** GET UNC PATH ***** *
        * ***** *
        */
        [DllImport("mpr.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int WNetGetConnection(
            [MarshalAs(UnmanagedType.LPTStr)] string localName,
            [MarshalAs(UnmanagedType.LPTStr)] StringBuilder remoteName,
            ref int length);

        /// <summary>
        /// Gets the unc path.
        /// </summary>
        /// <param name="originalPath">The original path.</param>
        /// <returns>String newPath</returns>
        public static string GetUNCPath(string originalPath)
        {
            StringBuilder sb = new StringBuilder(512);
            int size = sb.Capacity;

            // look for the {LETTER}: combination ...
            if (originalPath.Length > 2 && originalPath[1] == ':')
            {
                // don't use char.IsLetter here - as that can be misleading
                // the only valid drive letters are a-z && A-Z.
                char c = originalPath[0];
                if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'))
                {
                    int error = WNetGetConnection(originalPath.Substring(0, 2),
                        sb, ref size);
                    if (error == 0)
                    {
                        DirectoryInfo dir = new DirectoryInfo(originalPath);

                        string path = Path.GetFullPath(originalPath)
                            .Substring(Path.GetPathRoot(originalPath).Length);
                        string newPath = Path.Combine(sb.ToString().TrimEnd(), path);
                        return newPath;
                    }
                }
            }

            return originalPath;
        }

        /// <summary>
        /// Searches the directory for backup.
        /// </summary>
        /// <param name="targetDirectory">The target directory.</param>
        public void searchDirectoryForBackup(string targetDirectory)
        {            
            string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
            backupName = targetDirectory + backupDirectoryName;
            MessageBox.Show(backupName);
            foreach (string subdirectory in subdirectoryEntries)
            {
                if (Directory.Exists(backupName))
                {
                    backupFoundLabel.Text = "(Backup Found!)";
                    backupFoundLabel.ForeColor = System.Drawing.Color.ForestGreen;
                    userRestoreSelectCombo.Visible = true;
                    userRestoreSelectCombo.Items.Add(backupDirectoryName);
                    userRestoreSelectCombo.SelectedItem = userRestoreSelectCombo.FindString(backupDirectoryName);
                }
                else
                {
                    backupFoundLabel.Text = "(No Backup Found!)";
                    backupFoundLabel.ForeColor = System.Drawing.Color.Crimson;
                    userRestoreSelectCombo.Visible = false;
                }
            }
        }
        /// <summary>
        /// Enables buttons for restoring from backup
        /// </summary>
        /// <param name="backupDirectoryName">Name of the directory.</param>
        /// <param name="userName">Name of the user.</param>
        public void enableRestoreButtons(string backupDirectoryName)
        {
            string drive = restoreDriveCombo.SelectedItem.ToString();
            string backupDirectory = drive + "54HelperBackups\\" + backupDirectoryName;
            string backupSubDirectory;
            string[] directoryList =
            {
                "Documents",
                "Favorites",
                "Desktop",
                "Sticky Notes",
                "Pictures",
                "Videos",
                "Music"
            };
            foreach (string subDirectory in directoryList)
            {
                backupSubDirectory = backupDirectory + "\\" + subDirectory;
                if (Directory.Exists(backupSubDirectory))
                {
                    switch (subDirectory)
                    {
                        case "Documents":
                            recoverDocumentsLabel.Enabled = true;
                            recoverDocumentsLabel.Text = "(Found!)";
                            recoverDocumentsLabel.ForeColor = System.Drawing.Color.ForestGreen;
                            restoreDocumentsButton.Enabled = true;
                            break;
                        case "Favorites":
                            recoverFavoritesLabel.Enabled = true;
                            recoverFavoritesLabel.Text = "(Found!)";
                            recoverFavoritesLabel.ForeColor = System.Drawing.Color.ForestGreen;
                            restoreFavoritesButton.Enabled = true;
                            break;
                        case "Desktop":
                            recoverDesktopLabel.Enabled = true;
                            recoverDesktopLabel.Text = "(Found!)";
                            recoverDesktopLabel.ForeColor = System.Drawing.Color.ForestGreen;
                            restoreDesktopButton.Enabled = true;
                            break;
                        case "Sticky Notes":
                            recoverStickyNotesLabel.Enabled = true;
                            recoverStickyNotesLabel.Text = "(Found!)";
                            recoverStickyNotesLabel.ForeColor = System.Drawing.Color.ForestGreen;
                            restoreStickyNotesButton.Enabled = true;
                            break;
                        case "Pictures":
                            recoverPicturesLabel.Enabled = true;
                            recoverPicturesLabel.Text = "(Found!)";
                            recoverPicturesLabel.ForeColor = System.Drawing.Color.ForestGreen;
                            restorePicturesButton.Enabled = true;
                            break;
                        case "Videos":
                            recoverVideosLabel.Enabled = true;
                            recoverVideosLabel.Text = "(Found!)";
                            recoverVideosLabel.ForeColor = System.Drawing.Color.ForestGreen;
                            restoreVideosButton.Enabled = true;
                            break;
                        case "Music":
                            recoverMusicLabel.Enabled = true;
                            recoverMusicLabel.Text = "(Found!)";
                            recoverMusicLabel.ForeColor = System.Drawing.Color.ForestGreen;
                            restoreMusicButton.Enabled = true;
                            break;
                        default:
                            break;
                    }
                    if (restoreDesktopButton.Enabled && restoreDocumentsButton.Enabled && restoreFavoritesButton.Enabled)
                    {
                        recoverAllEssentialsLabel.ForeColor = System.Drawing.Color.ForestGreen;
                        restoreAllEssentialsButton.Enabled = true;
                    }
                }
            }
        }

        /// <summary>
        /// Handles the Click event of the userSetupCompleteButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void userSetupCompleteButton_Click(object sender, EventArgs e)
        {
            string userName = getUsername().ToString();
            using (BeginWaitCursorBlock())
            {
                //1YWi6rMUie3BD_AXH85Ew0TGJpgP_AVNN4Ury2bpIaQk - setup spreadsheet ID
                SheetsService service = authenticateServiceAccount();
                ValueRange valueRange = new ValueRange();
                var oblist = new List<object>() { };
                //set year
                string yearDate = DateTime.Today.ToString("yyyy");
                oblist.Add(yearDate);
                //set date
                string monthDay = DateTime.Today.ToString("MM/dd");
                oblist.Add(monthDay);
                //set location
                string serverOutput = Regex.Replace(serverName, @"[\d-]", string.Empty);
                oblist.Add(serverOutput);
                //set username
                oblist.Add(userName);

                //create array of checkboxes in appropriate order for spreadsheet (must match the order of columns in the spreadsheet!)
                CheckBox[] checkBoxNames =
                {
                outlookCheckBox,
                quickenCheckBox,
                adobeProCheckBox,
                icPrintingCheckBox,
                dymoPrintingCheckBox,
                scanSnapCheckBox,
                installPrintersCheckBox,
                imageRunnerCheckBox,
                restoreFavoritesCheckBox,
                homeShortcutCheckBox,
                efinanceShortcutCheckBox,
                wordShortcutCheckBox,
                icShortcutCheckBox,
                aesopShortcutCheckBox
            };
                //set X's for checked and O for unchecked
                //iterate through checkBoxNames Array
                foreach (Control c in checkBoxNames)
                {
                    if ((c is CheckBox) && ((CheckBox)c).Checked)
                    {
                        oblist.Add("X");
                    }
                    if ((c is CheckBox) && ((CheckBox)c).Checked == false)
                    {
                        oblist.Add("O");
                    }
                }
                valueRange.Values = new List<IList<object>> { oblist };
                String spreadsheetId = "1YWi6rMUie3BD_AXH85Ew0TGJpgP_AVNN4Ury2bpIaQk";
                //If the user has been Setup once before - run an update command on the appropriate row - otherwise append it to end of spreadsheet
                if (userHasBeenSetup(userName))
                {
                    int x = existingSetupRowNumber;
                    String customRange = "A" + (x + 1);
                    SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, customRange);
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                    UpdateValuesResponse result = update.Execute();
                    updateUserSetupChecks();
                    userSetupAnswerLabel.Text = "YES";
                    userSetupAnswerLabel.ForeColor = System.Drawing.Color.ForestGreen;
                    MessageBox.Show("The users setup information has been updated.");
                }
                else
                {
                    // Define request parameters.
                    String range = "A1";
                    SpreadsheetsResource.ValuesResource.AppendRequest update = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                    AppendValuesResponse result = update.Execute();
                    updateUserSetupChecks();
                    userSetupAnswerLabel.Text = "YES";
                    userSetupAnswerLabel.ForeColor = System.Drawing.Color.ForestGreen;
                    MessageBox.Show("The users setup information has been recorded.");
                }
            }
        }

        /// <summary>
        /// Handles the Click event of the userBackupCompleteButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void userBackupCompleteButton_Click(object sender, EventArgs e)
        {
            string userName = getUsername().ToString();
            using (BeginWaitCursorBlock())
            {
                //1XFsFJ2nqrXsxO9ShRJOwQ2MMlDpxf1wmU7RgSKUvmKM - spreadsheet ID
                SheetsService service = authenticateServiceAccount();
                ValueRange valueRange = new ValueRange();
                var oblist = new List<object>() { };
                //set year
                string yearDate = DateTime.Today.ToString("yyyy");
                oblist.Add(yearDate);
                //set date
                string monthDay = DateTime.Today.ToString("MM/dd");
                oblist.Add(monthDay);
                //set location
                string serverOutput = Regex.Replace(serverName, @"[\d-]", string.Empty);
                oblist.Add(serverOutput);
                //set username
                oblist.Add(userName);
                //create array of checkboxes in appropriate order for spreadsheet (must match the order of columns in the spreadsheet!)
                CheckBox[] checkBoxNames =
                {
                desktopBackupCheckBox,
                documentsBackupCheckBox,
                favoritesBackupCheckBox,
                quickenBackupCheckBox,
                stickyNotesBackupCheckBox,
                picturesBackupCheckBox,
                videosBackupCheckBox,
                musicBackupCheckBox,
            };
                //set X's for checked and O for unchecked
                //iterate through checkBoxNames Array
                foreach (Control c in checkBoxNames)
                {
                    if ((c is CheckBox) && ((CheckBox)c).Checked)
                    {
                        oblist.Add("X");
                    }
                    if ((c is CheckBox) && ((CheckBox)c).Checked == false)
                    {
                        oblist.Add("O");
                    }
                }
                valueRange.Values = new List<IList<object>> { oblist };
                // Define request parameters.
                String spreadsheetId = "1XFsFJ2nqrXsxO9ShRJOwQ2MMlDpxf1wmU7RgSKUvmKM";
                //If the user has been Setup once before - run an update command on the appropriate row - otherwise append it to end of spreadsheet
                if (userHasBeenBackedUp(userName))
                {
                    int x = existingSetupRowNumber;
                    String customRange = "A" + (x + 1);
                    SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, customRange);
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                    UpdateValuesResponse result = update.Execute();
                    updateUserBackupChecks();
                    userBackedUpAnswerLabel.Text = "YES";
                    userBackedUpAnswerLabel.ForeColor = System.Drawing.Color.ForestGreen;
                    MessageBox.Show("The users backup information has been updated.");
                }
                else
                {
                    // Define request parameters.
                    String range = "A1";
                    SpreadsheetsResource.ValuesResource.AppendRequest update = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                    AppendValuesResponse result = update.Execute();
                    updateUserBackupChecks();
                    userBackedUpAnswerLabel.Text = "YES";
                    userBackedUpAnswerLabel.ForeColor = System.Drawing.Color.ForestGreen;
                    MessageBox.Show("The users backup information has been recorded.");
                }
            }
        }

        /// <summary>
        /// Checks the Google Spreadsheet to see if the user exists in the current year as having been setup.
        /// If the user DOES exist - it updates the UserChecklist tab accordingly in the Setup section
        /// </summary>
        /// <returns>Bool</returns>
        public bool userHasBeenSetup(string userName)
        {
            SheetsService service = authenticateServiceAccount();

            // Define request parameters.
            String spreadsheetId = "1YWi6rMUie3BD_AXH85Ew0TGJpgP_AVNN4Ury2bpIaQk";
            String range = "A1:R2000";
            
            SpreadsheetsResource.ValuesResource.GetRequest getData = service.Spreadsheets.Values.Get(spreadsheetId, range);
            getData.MajorDimension = SpreadsheetsResource.ValuesResource.GetRequest.MajorDimensionEnum.ROWS;
            ValueRange sheetData = getData.Execute();

            int i = 0;
            foreach (var row in sheetData.Values)
            {
                if (row[0].ToString() == DateTime.Today.ToString("yyyy") && row[3].ToString() == userName)
                {
                    existingSetupRowNumber = i;
                    return true;
                }
                i++;
            }
            return false;
        }

        /// <summary>
        /// Updates the user setup checks.
        /// </summary>
        public void updateUserSetupChecks()
        {
            SheetsService service = authenticateServiceAccount();

            // Define request parameters.
            String spreadsheetId = "1YWi6rMUie3BD_AXH85Ew0TGJpgP_AVNN4Ury2bpIaQk";
            String range = "A1:R2000";

            SpreadsheetsResource.ValuesResource.GetRequest getData = service.Spreadsheets.Values.Get(spreadsheetId, range);
            getData.MajorDimension = SpreadsheetsResource.ValuesResource.GetRequest.MajorDimensionEnum.ROWS;
            ValueRange sheetData = getData.Execute();

            //get the appropriate row
            var row = sheetData.Values[existingSetupRowNumber];

            //create array of checkboxes in appropriate order for spreadsheet (must match the order of columns in the spreadsheet!)
            CheckBox[] checkBoxNames =
            {
                outlookCheckBox,
                quickenCheckBox,
                adobeProCheckBox,
                icPrintingCheckBox,
                dymoPrintingCheckBox,
                scanSnapCheckBox,
                installPrintersCheckBox,
                imageRunnerCheckBox,
                restoreFavoritesCheckBox,
                homeShortcutCheckBox,
                efinanceShortcutCheckBox,
                wordShortcutCheckBox,
                icShortcutCheckBox,
                aesopShortcutCheckBox
            };
            int i = 4;
            foreach (CheckBox c in checkBoxNames)
            {
                string cellContent = row[i].ToString();
                if (cellContent == "X")
                {
                    c.Checked = true;
                    i++;
                }
                if (cellContent == "O")
                {
                    c.Checked = false;
                    i++;
                }
            }
        }
        /// <summary>
        /// Check to see if the user has been backed up.
        /// </summary>
        /// <returns>Bool</returns>
        public bool userHasBeenBackedUp(string userName)
        {
            SheetsService service = authenticateServiceAccount();

            // Define request parameters.
            String spreadsheetId = "1XFsFJ2nqrXsxO9ShRJOwQ2MMlDpxf1wmU7RgSKUvmKM";
            String range = "A1:L2000";

            SpreadsheetsResource.ValuesResource.GetRequest getData = service.Spreadsheets.Values.Get(spreadsheetId, range);
            getData.MajorDimension = SpreadsheetsResource.ValuesResource.GetRequest.MajorDimensionEnum.ROWS;
            ValueRange sheetData = getData.Execute();

            int i = 0;
            foreach (var row in sheetData.Values)
            {
                if (row[0].ToString() == DateTime.Today.ToString("yyyy") && row[3].ToString() == userName)
                {
                    existingBackupRowNumber = i;
                    return true;
                }
                i++;
            }
            return false;
        }

        /// <summary>
        /// Updates the user backup checkboxes.
        /// </summary>
        public void updateUserBackupChecks()
        {
            SheetsService service = authenticateServiceAccount();

            // Define request parameters.
            String spreadsheetId = "1XFsFJ2nqrXsxO9ShRJOwQ2MMlDpxf1wmU7RgSKUvmKM";
            String range = "A1:L2000";

            SpreadsheetsResource.ValuesResource.GetRequest getData = service.Spreadsheets.Values.Get(spreadsheetId, range);
            getData.MajorDimension = SpreadsheetsResource.ValuesResource.GetRequest.MajorDimensionEnum.ROWS;
            ValueRange sheetData = getData.Execute();

            //get the appropriate row
            var row = sheetData.Values[existingBackupRowNumber];

            //create array of checkboxes in appropriate order for spreadsheet (must match the order of columns in the spreadsheet!)
            CheckBox[] checkBoxNames =
            {
                desktopBackupCheckBox,
                documentsBackupCheckBox,
                favoritesBackupCheckBox,
                quickenBackupCheckBox,
                stickyNotesBackupCheckBox,
                picturesBackupCheckBox,
                videosBackupCheckBox,
                musicBackupCheckBox,
            };
            int i = 4;
            foreach (CheckBox c in checkBoxNames)
            {
                string cellContent = row[i].ToString();
                if (cellContent == "X")
                {
                    c.Checked = true;
                    i++;
                }
                if (cellContent == "O")
                {
                    c.Checked = false;
                    i++;
                }
            }
        }

        /// <summary>
        /// Handles the SelectedIndexChanged event of the userBackupSelectCombo control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void userBackupSelectCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            //set the backupDirectoryName
            backupDirectoryName = userBackupSelectCombo.SelectedItem.ToString() + "-54Help-" + DateTime.Now.Year.ToString();
            labelDirectorySizes(userBackupSelectCombo.SelectedItem.ToString());
        }

        /// <summary>
        /// Checks to see if the user has access to get quicken.
        /// </summary>
        /// <returns>Bool</returns>
        public bool doesUserGetQuicken(string userName)
        {
            SheetsService service = authenticateServiceAccount();

            // Define request parameters.
            String spreadsheetId = "1-wg63_jtlZfT-De6zi2zG705_oIF2TgfkLDijmcZgIc";
            String range = "A2:A100";

            SpreadsheetsResource.ValuesResource.GetRequest getData = service.Spreadsheets.Values.Get(spreadsheetId, range);
            getData.MajorDimension = SpreadsheetsResource.ValuesResource.GetRequest.MajorDimensionEnum.ROWS;
            ValueRange sheetData = getData.Execute();

            var i = 0;
            var secretaryCount = sheetData.Values.Count();

            while (i < secretaryCount)
            {
                foreach (var cell in sheetData.Values[i])
                {
                    if (userName == cell.ToString())
                    {
                        return true;
                    }
                }
                i++;
            }
            return false;
        }

        /// <summary>
        /// Handles the SelectedIndexChanged event of the userRestoreSelectCombo control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void userRestoreSelectCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            using (BeginWaitCursorBlock()) {
                backupToRestore = userRestoreSelectCombo.SelectedItem.ToString();
                string selectedBackup = userRestoreSelectCombo.SelectedItem.ToString();
                string selectedBackupUsername = selectedBackup.Split('-')[0];
                //Check that the user backup selected has a user folder existing on the computer
                if (Directory.Exists("C:\\Users\\" + selectedBackupUsername))
                {
                    //user account exists on PC - enable buttons to restore data
                    enableRestoreButtons(userRestoreSelectCombo.SelectedItem.ToString());
                }
                else
                {
                    disableRestoreButtons();
                    MessageBox.Show("The backup you've selected is for a user that has not logged on to this PC yet. Please have " + selectedBackupUsername + " logon to this PC before restoring from this backup.");
                }
            }
        }
        public bool isRecoveryForLoggedInUser()
        {
            string userName = getUsername().ToString();
            string selectedBackup = userRestoreSelectCombo.SelectedItem.ToString();
            string selectedBackupUsername = selectedBackup.Split('-')[0];
            if (userName == selectedBackupUsername)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// Authenticates the service account.
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception">Create ServiceAccount Failed</exception>
        public static SheetsService authenticateServiceAccount()
        {            
            try
            {
                GoogleCredential credential;
                string applicationName = "SD54Helper";
                using (var stream = new FileStream(Environment.CurrentDirectory + @"\json\SD54HelperServiceAccount.json", FileMode.Open, FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream)
                          .CreateScoped(SheetsService.Scope.Spreadsheets, SheetsService.Scope.Drive);
                }
                // Create Google Sheets API service.
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = applicationName,
                });
                return service;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Create service account SD54HelperServiceAccount failed" + ex.Message);
                throw new Exception("Create ServiceAccount Failed", ex);
            }
            
        }

        /// <summary>
        /// Begins the wait cursor block.
        /// </summary>
        /// <returns></returns>
        public static IDisposable BeginWaitCursorBlock()
        {
            return ((!_waitCursorIsActive) ? (IDisposable)new waitCursor() : null);
        }
        private static bool _waitCursorIsActive;
        private class waitCursor : IDisposable
        {
            private Cursor oldCur;
            public waitCursor()
            {
                _waitCursorIsActive = true;
                oldCur = Cursor.Current;
                Cursor.Current = Cursors.WaitCursor;
            }
            public void Dispose()
            {
                Cursor.Current = oldCur;
                _waitCursorIsActive = false;
            }
        }

        /// <summary>
        /// Disables the restore buttons.
        /// </summary>
        private void disableRestoreButtons()
        {
            //disable buttons in restore Panel
            foreach (Control cont in restoreEssentialGroupBox.Controls)
            {
                cont.Enabled = false;
                if (cont is Label)
                {
                    cont.Text = "(Not Found)";
                }
            }
            foreach (Control cont in restoreAdditionalGroupBox.Controls)
            {
                cont.Enabled = false;
                if (cont is Label)
                {
                    cont.Text = "(Not Found)";
                }
            }

        }
        public static uint WinMajorVersion
        {
            get
            {
                dynamic major;
                // The 'CurrentMajorVersionNumber' string value in the CurrentVersion key is new for Windows 10, 
                // and will most likely (hopefully) be there for some time before MS decides to change this - again...
                if (TryGetRegistryKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentMajorVersionNumber", out major))
                {
                    return (uint)major;
                }

                // When the 'CurrentMajorVersionNumber' value is not present we fallback to reading the previous key used for this: 'CurrentVersion'
                dynamic version;
                if (!TryGetRegistryKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentVersion", out version))
                    return 0;

                var versionParts = ((string)version).Split('.');
                if (versionParts.Length != 2) return 0;
                uint majorAsUInt;
                return uint.TryParse(versionParts[0], out majorAsUInt) ? majorAsUInt : 0;
            }
        }
        private static bool TryGetRegistryKey(string path, string key, out dynamic value)
        {
            value = null;
            try
            {
                var rk = Registry.LocalMachine.OpenSubKey(path);
                if (rk == null) return false;
                value = rk.GetValue(key);
                return value != null;
            }
            catch
            {
                return false;
            }
        }
    }
}
