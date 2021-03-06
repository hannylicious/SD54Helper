﻿using Google.Apis.Auth.OAuth2;
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


namespace Helpdesk54
{
    public partial class UserControl1 : Form

    {
        int existingSetupRowNumber, existingBackupRowNumber;
        long canItFit;
        string userDisplayFirstName, userDisplayLastName, userCustomDisplayName;
        string backupDirectoryName;
        string backupName;
        string serverName;
        string homeDirectory;
        string desktopFolder, documentsFolder, favoritesFolder;
        string userName;
        string wordPath, excelPath, outlookPath;
        BackgroundWorker essentialBgWorker = new BackgroundWorker();
        BackgroundWorker additionalBgWorker = new BackgroundWorker();
        BackgroundWorker restoreEssentialBgWorker = new BackgroundWorker();
        BackgroundWorker restoreAdditionalBgWorker = new BackgroundWorker();
        string selectedDrive;
        string clickedButton;
        string itemsChanged;
        long selectedDriveAvailableSize;
        int totalFileCount;
        string destinationLocation;
        DirectoryInfo source;
        int fileCount;
        DriveInfo[] theDrives;

        public UserControl1()
        {
            InitializeComponent();
            //set the paths to wordPath, excelPath and outlookPath           
            setOfficePaths();
            //Get the userName to the currently logged in user
            getUsername();         
            usernameLabel.Text = userName;
            //Get the attached drives            
            getAttachedDrives();
            //set the backupDirectoryName
            backupDirectoryName = userName.ToString() + "-Backups-" + DateTime.Now.Year.ToString();
            //check if the user gets access to quicken
            if (doesUserGetQuicken())
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
            if (userHasBeenSetup())
            {
                updateUserSetupChecks();
                userSetupAnswerLabel.Text = "YES";
                userSetupAnswerLabel.ForeColor = System.Drawing.Color.ForestGreen;

            }
            if (userHasBeenBackedUp())
            {
                updateUserBackupChecks();
                userBackedUpAnswerLabel.Text = "YES";
                userBackedUpAnswerLabel.ForeColor = System.Drawing.Color.ForestGreen;
            }
            //set backupDriveCombo dropdown
            setBackupDriveCombo();
            //set restoreDriveCombo to dropdown
            setRestoreDriveCombo();
            //Set the selected drive freespace label 
            setBackupDriveFreeSpaceLabel();
            //Check the H:\ drive for a backup
            checkForExistingBackup();
            //update labels
            labelDirectorySizes();
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
        }

        /// <summary>
        /// Sets the backup drive combo.
        /// </summary>
        private void setBackupDriveCombo()
        {
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

        /// <summary>
        /// Sets the server name link.
        /// </summary>
        private void setServerNameLink()
        {
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
                    Uri uri = new Uri(path);
                    serverName = uri.Host.ToString();
                    serverNameLinkLabel.Text = serverName;
                }
                else
                {
                    homeDirectory = "Unknown";
                    serverNameLinkLabel.Text = "Unknown";
                }
            }
        }

        /// <summary>
        /// Sets the office paths.
        /// </summary>
        private void setOfficePaths()
        {
            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe");
            if (key == null)
            {
                wordPath = "";
            }
            else
            {
                wordPath = key.GetValue("").ToString();
            }

            key = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe");
            if (key == null)
            {
                excelPath = "";
            }
            else
            {
                excelPath = key.GetValue("").ToString();
            }

            key = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.exe");
            if (key == null)
            {
                outlookPath = "";
            }
            else
            {
                outlookPath = key.GetValue("").ToString();
            }
        }

        /// <summary>
        /// Gets the attached drives.
        /// </summary>
        /// <returns>Array theDrives</returns>
        private Array getAttachedDrives()
        {
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
            userName = Environment.UserName;
            userDisplayFirstName = UserPrincipal.Current.GivenName;
            userDisplayLastName = UserPrincipal.Current.Surname;
            userCustomDisplayName = userDisplayFirstName + userDisplayLastName;
            return userName;
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
        /// Checks for existing backup on H: drive.
        /// </summary>
        private void checkForExistingBackup() {
                    try
                    {
                        //Check the selected restoreComboDrive for a backup matching: username-Backup-year
                        DriveInfo restoreSelectedDrive = (DriveInfo)restoreDriveCombo.SelectedItem;
                        searchDirectoryForBackup(restoreSelectedDrive.ToString());
                        if (backupFoundLabel.Text.ToString() == "(Backup Found!)")
                        {
                            checkBackupDirectories(backupName);
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Exception source from checkForExistingBackup() : {0}", e.Source);

                        throw;
                    }
        }

        /// <summary>
        /// Checks the application installs.
        /// </summary>
        private void checkApplicationInstalls()
        {
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
            string stickyNotesDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Microsoft\Sticky Notes\";
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
                    path = outlookPath;
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
                    labelDirectorySizes();
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
                        labelDirectorySizes();
                        backupDriveLabel.Text = driveFreeSpace + " Free";
                    }
                }
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
                searchDirectoryForBackup(selectedDrive.ToString());
                if (backupFoundLabel.Text.ToString() == "(Backup Found!)")
                {
                    checkBackupDirectories(backupName);
                }
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
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            essentialBgWorker.RunWorkerAsync();
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
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            essentialBgWorker.RunWorkerAsync();
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
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            essentialBgWorker.RunWorkerAsync();
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
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            essentialBgWorker.RunWorkerAsync();
        }

        /// <summary>
        /// Handles the DoWork event of the essentialBgWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="DoWorkEventArgs"/> instance containing the event data.</param>
        void essentialBgWorker_DoWork(object sender, DoWorkEventArgs e)
        {            
            string buttonSender = clickedButton; //backupDesktopButton, backupDocumentsButton, etc.
            switch (buttonSender)
            {
                case "backupDesktopButton":
                    backupDesktop();
                    break;
                case "backupDocumentsButton":
                    backupDocuments();
                    break;
                case "backupFavoritesButton":
                    backupFavorites();
                    break;
                case "backupAllEssentialsButton":
                    backupFavorites();
                    backupDesktop();
                    backupDocuments();
                    break;
                default:
                    essentialBgWorker.CancelAsync();
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
            checkForExistingBackup();
            checkBackupDirectories(backupName);
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
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            additionalBgWorker.RunWorkerAsync();
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
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            additionalBgWorker.RunWorkerAsync();
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
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            additionalBgWorker.RunWorkerAsync();
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
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            additionalBgWorker.RunWorkerAsync();
        }

        /// <summary>
        /// Handles the DoWork event of the additionalBgWorker control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="DoWorkEventArgs"/> instance containing the event data.</param>
        void additionalBgWorker_DoWork(object sender, DoWorkEventArgs e)
        {

            string stickyNotesFolder, picturesFolder, videosFolder, musicFolder;
            string buttonSender = clickedButton; //Desktop, Documents, etc.
            switch (buttonSender)
            {
                case "backupStickyNotesButton":
                    destinationLocation = selectedDrive + backupDirectoryName + "\\Sticky Notes\\";
                    stickyNotesFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Microsoft\Sticky Notes\";
                    source = new DirectoryInfo(stickyNotesFolder);
                    break;
                case "backupPicturesButton":
                    destinationLocation = selectedDrive + backupDirectoryName + "\\Pictures\\";
                    picturesFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
                    source = new DirectoryInfo(picturesFolder);
                    break;
                case "backupVideosButton":
                    string pathWithEvn = @"%USERPROFILE%\Videos";
                    destinationLocation = selectedDrive + backupDirectoryName + "\\Videos\\";
                    videosFolder = Environment.ExpandEnvironmentVariables(pathWithEvn);
                    source = new DirectoryInfo(videosFolder);
                    break;
                case "backupMusicButton":
                    destinationLocation = selectedDrive + backupDirectoryName + "\\Music\\";
                    musicFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyMusic);
                    source = new DirectoryInfo(musicFolder);
                    break;
                default:

                    break;
            }
            DirectoryInfo target = new DirectoryInfo(destinationLocation);
            fileCount = source.GetFiles("*", SearchOption.AllDirectories).Length;
            totalFileCount = fileCount;
            int total = totalFileCount; //total things being transferred
            for (int i = 0; i <= total; i++) //report those numbers
            {
                System.Threading.Thread.Sleep(100);
                int percents = (i * 100) / total;
                additionalBgWorker.ReportProgress(percents, i);
                //2 arguments:
                //1. procenteges (from 0 t0 100) - i do a calcumation 
                //2. some current value!
            }
            if (!Directory.Exists(destinationLocation))
            {
                Directory.CreateDirectory(destinationLocation);
            }
            CopyFilesRecursively(source, target);
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
            checkForExistingBackup();
            checkBackupDirectories(backupName);
        }

        /// <summary>
        /// Labels the directory sizes.
        /// </summary>
        public void labelDirectorySizes()
        {
            string[] directoryLocations =
            {
            "DesktopDirectory",
            "MyDocuments",
            "Favorites",
            "MyMusic",
            "MyPictures",
            "My Videos"
            };
            // Get the directory sizes for each directoryLocation & set the label
            foreach (string directoryLocation in directoryLocations)
            {
                string userDirectoryLocation = Environment.GetEnvironmentVariable("userprofile");
                //'My Videos' is not supported in older frameworks so set it seperately
                if (directoryLocation == "My Videos")
                {
                    string folderName = "Videos";
                    string folder = userDirectoryLocation + "\\" + folderName;
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
                    var dir = (Environment.SpecialFolder)Enum.Parse(typeof(Environment.SpecialFolder), directoryLocation);
                    string folder = Environment.GetFolderPath(dir);
                    long folderSize = DirSize(new DirectoryInfo(folder));
                    string folderMB = FormatBytes(folderSize);
                    var selectedDriveSize = selectedDriveAvailableSize;
                    
                    switch (directoryLocation)
                    {
                        case "DesktopDirectory":
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
                        backupAllEssentialsButton.Enabled = true;
                    }
                }

            }
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
            //Set the size & label for Sticky Notes
            string stickyNotesFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            stickyNotesFolder = stickyNotesFolder + "\\Microsoft\\Sticky Notes";
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
        /// Backups the desktop.
        /// </summary>
        private void backupDesktop()
        {
            destinationLocation = selectedDrive + backupDirectoryName + "\\Desktop\\";
            desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            source = new DirectoryInfo(desktopFolder);
            DirectoryInfo desktopSource = new DirectoryInfo(desktopFolder);
            DirectoryInfo target = new DirectoryInfo(destinationLocation);
            try
            {
                fileCount = source.GetFiles("*", SearchOption.AllDirectories).Length;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            totalFileCount = fileCount;
            int total = totalFileCount; //total things being transferred
            for (int i = 0; i <= total; i++) //report those numbers
            {
                System.Threading.Thread.Sleep(100);
                int percents = (i * 100) / total;
                essentialBgWorker.ReportProgress(percents, i);
                //2 arguments:
                //1. procenteges (from 0 t0 100) - i do a calcumation 
                //2. some current value!
            }
            if (!Directory.Exists(destinationLocation))
            {
                Directory.CreateDirectory(destinationLocation);
            }
            CopyFilesRecursively(desktopSource, target);
        }

        /// <summary>
        /// Backups the favorites.
        /// </summary>
        private void backupFavorites()
        {
            /*FAVORITES*/
            destinationLocation = selectedDrive + backupDirectoryName + "\\Favorites\\";
            favoritesFolder = Environment.GetFolderPath(Environment.SpecialFolder.Favorites);
            DirectoryInfo source = new DirectoryInfo(favoritesFolder);
            DirectoryInfo target = new DirectoryInfo(destinationLocation);
            DirectoryInfo favoritesSource = new DirectoryInfo(favoritesFolder);
            try
            {
                fileCount = source.GetFiles("*", SearchOption.AllDirectories).Length;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            totalFileCount = fileCount;
            int total = totalFileCount; //total things being transferred
            for (int i = 0; i <= total; i++) //report those numbers
            {
                System.Threading.Thread.Sleep(100);
                int percents = (i * 100) / total;
                essentialBgWorker.ReportProgress(percents, i);
                //2 arguments:
                //1. procenteges (from 0 t0 100) - i do a calcumation 
                //2. some current value!
            }
            if (!Directory.Exists(destinationLocation))
            {
                Directory.CreateDirectory(destinationLocation);
            }
            CopyFilesRecursively(favoritesSource, target);
        }

        /// <summary>
        /// Backups the documents.
        /// </summary>
        private void backupDocuments()
        {
            /*DOCUMENTS*/
            destinationLocation = selectedDrive + backupDirectoryName + "\\Documents\\";
            documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            source = new DirectoryInfo(documentsFolder);
            DirectoryInfo target = new DirectoryInfo(destinationLocation);
            DirectoryInfo documentsSource = new DirectoryInfo(documentsFolder);
            int fileCount = 0;
            //retrieve files
            fileCount += Directory.GetFiles(documentsFolder, "*.*", SearchOption.TopDirectoryOnly).Length;
            //retrieve folders
            foreach (var subFolder in Directory.GetDirectories(documentsFolder, "*.*", SearchOption.TopDirectoryOnly))
            {
                /*set the unauthorized directories to skip*/       
                string[] folderArray = {
                    @"C:\Users\" + usernameLabel.Text.ToString() + @"\Documents\My Pictures",
                    @"C:\Users\" + usernameLabel.Text.ToString() + @"\Documents\My Videos",
                    @"C:\Users\" + usernameLabel.Text.ToString() + @"\Documents\My Music",
                    };
                if (folderArray.Contains(subFolder))
                {
                    continue;
                }
                else
                {
                    try
                    {
                        Directory.GetFiles(subFolder, "*.*", SearchOption.AllDirectories).ToList().ForEach(f => fileCount++);
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.ToString());
                    }
                }
            }
            /*
            try
            {
                fileCount = source.GetFiles("*", SearchOption.AllDirectories).Length;    
            } catch (Exception e) {
                MessageBox.Show(e.ToString());
            }
            */
            totalFileCount = fileCount;
            int total = totalFileCount; //total things being transferred
            for (int i = 0; i <= total; i++) //report those numbers
            {
                System.Threading.Thread.Sleep(100);
                int percents = (i * 100) / total;
                essentialBgWorker.ReportProgress(percents, i);
                //2 arguments:
                //1. procenteges (from 0 t0 100) - i do a calcumation 
                //2. some current value!
            }
            if (!Directory.Exists(destinationLocation))
            {
                Directory.CreateDirectory(destinationLocation);
            }
            CopyFilesRecursively(documentsSource, target);
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
            selectedDrive = restoreDriveCombo.SelectedItem.ToString();
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
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
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
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
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
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            restoreEssentialBgWorker.RunWorkerAsync();
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            string[] directoryNameArray = { "Documents", "Favorites", "Desktop" };
            string destinationLocation = "";
            string directoryToBackup = "";

            for (int i = 0; i < directoryNameArray.Length; i++)
            {
                switch (directoryNameArray[i])
                {
                    case "Documents":
                        destinationLocation = selectedDrive + backupDirectoryName + "\\" + directoryNameArray[i] + "\\";
                        directoryToBackup = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                        break;
                    case "Favorites":
                        destinationLocation = selectedDrive + backupDirectoryName + "\\" + directoryNameArray[i] + "\\";
                        directoryToBackup = Environment.GetFolderPath(Environment.SpecialFolder.Favorites);
                        break;
                    case "Desktop":
                        destinationLocation = selectedDrive + backupDirectoryName + "\\" + directoryNameArray[i] + "\\";
                        directoryToBackup = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        break;
                    default:
                        break;
                }
                if (!Directory.Exists(destinationLocation))
                {
                    Directory.CreateDirectory(destinationLocation);
                }
                DirectoryInfo target = new DirectoryInfo(destinationLocation);
                DirectoryInfo source = new DirectoryInfo(directoryToBackup);
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
            
            string buttonSender = clickedButton; //Desktop, Documents, etc.
            switch (buttonSender)
            {
                case "restoreDesktopButton":
                    restoreDesktop();
                    break;
                case "restoreDocumentsButton":
                    restoreDocuments();
                    break;
                case "restoreFavoritesButton":
                    restoreFavorites();
                    break;
                case "restoreAllEssentialsButton":
                    restoreFavorites();
                    restoreDesktop();
                    restoreDocuments();
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
        /// Restores the desktop.
        /// </summary>
        private void restoreDesktop()
        {
            string userDirectory = @"C:\Users\" + usernameLabel.Text.ToString();
            destinationLocation = userDirectory + "\\Desktop\\";
            string desktopBackupFolder = selectedDrive + backupDirectoryName + "\\Desktop\\";
            source = new DirectoryInfo(desktopBackupFolder);
            DirectoryInfo target = new DirectoryInfo(destinationLocation);
            fileCount = source.GetFiles("*", SearchOption.AllDirectories).Length;
            totalFileCount = fileCount;
            int total = totalFileCount; //total things being transferred
            for (int i = 0; i <= total; i++) //report those numbers
            {
                System.Threading.Thread.Sleep(100);
                int percents = (i * 100) / total;
                restoreEssentialBgWorker.ReportProgress(percents, i);
                //2 arguments:
                //1. procenteges (from 0 t0 100) - i do a calcumation 
                //2. some current value!
            }
            if (!Directory.Exists(destinationLocation))
            {
                MessageBox.Show("This location does not exist for the logged in user.");
            }
            RestoreFilesRecursively(source, target);
        }

        /// <summary>
        /// Restores the favorites.
        /// </summary>
        private void restoreFavorites()
        {
            string userDirectory = @"C:\Users\" + usernameLabel.Text.ToString();
            destinationLocation = userDirectory + "\\Favorites\\";
            string favoritesBackupFolder = selectedDrive + backupDirectoryName + "\\Favorites\\";
            source = new DirectoryInfo(favoritesBackupFolder);
            DirectoryInfo target = new DirectoryInfo(destinationLocation);
            fileCount = source.GetFiles("*", SearchOption.AllDirectories).Length;
            totalFileCount = fileCount;
            int total = totalFileCount; //total things being transferred
            for (int i = 0; i <= total; i++) //report those numbers
            {
                System.Threading.Thread.Sleep(100);
                int percents = (i * 100) / total;
                restoreEssentialBgWorker.ReportProgress(percents, i);
                //2 arguments:
                //1. procenteges (from 0 t0 100) - i do a calcumation 
                //2. some current value!
            }
            if (!Directory.Exists(destinationLocation))
            {
                MessageBox.Show("This location does not exist for the logged in user.");
            }
            RestoreFilesRecursively(source, target);
        }

        /// <summary>
        /// Restores the documents.
        /// </summary>
        private void restoreDocuments()
        {
            string userDirectory = @"C:\Users\" + usernameLabel.Text.ToString();
            destinationLocation = userDirectory + "\\Documents\\";
            string documentsBackupFolder = selectedDrive + backupDirectoryName + "\\Documents\\";
            source = new DirectoryInfo(documentsBackupFolder);
            DirectoryInfo target = new DirectoryInfo(destinationLocation);
            fileCount = source.GetFiles("*", SearchOption.AllDirectories).Length;
            totalFileCount = fileCount;
            int total = totalFileCount; //total things being transferred
            for (int i = 0; i <= total; i++) //report those numbers
            {
                System.Threading.Thread.Sleep(100);
                int percents = (i * 100) / total;
                restoreEssentialBgWorker.ReportProgress(percents, i);
                //2 arguments:
                //1. procenteges (from 0 t0 100) - i do a calcumation 
                //2. some current value!
            }
            if (!Directory.Exists(destinationLocation))
            {
                MessageBox.Show("This location does not exist for the logged in user.");
            }
            RestoreFilesRecursively(source, target);
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
            selectedDrive = restoreDriveCombo.SelectedItem.ToString();
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
            selectedDrive = restoreDriveCombo.SelectedItem.ToString();
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
            selectedDrive = restoreDriveCombo.SelectedItem.ToString();
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
            selectedDrive = restoreDriveCombo.SelectedItem.ToString();
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
            string userDirectory = @"C:\Users\" + usernameLabel.Text.ToString();
            string buttonSender = clickedButton; //Desktop, Documents, etc.
            switch (buttonSender)
            {
                case "Sticky Notes":
                    destinationLocation = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Microsoft\Sticky Notes\";
                    string stickyNotesBackupFolder = selectedDrive + backupDirectoryName + "\\Sticky Notes\\";
                    source = new DirectoryInfo(stickyNotesBackupFolder);
                    break;
                case "Pictures":
                    destinationLocation = userDirectory + "\\Pictures\\";
                    string picturesBackupFolder = selectedDrive + backupDirectoryName + "\\Pictures\\";
                    source = new DirectoryInfo(picturesBackupFolder);
                    break;
                case "Videos":
                    destinationLocation = userDirectory + "\\Videos\\";
                    string videosFolder = selectedDrive + backupDirectoryName + "\\Videos\\";
                    source = new DirectoryInfo(videosFolder);
                    break;
                case "Music":
                    destinationLocation = userDirectory + "\\Music\\";
                    string musicFolder = selectedDrive + backupDirectoryName + "\\Music\\";
                    source = new DirectoryInfo(musicFolder);
                    break;
                default:

                    break;
            }
            //By default - Sticky Notes is really the only file that might fail here - the rest are system folders which will always exist
            if (!Directory.Exists(destinationLocation))
            {
                if (buttonSender == "Sticky Notes")
                {
                    DialogResult result = MessageBox.Show(String.Format("The directory {0} does not exist. Would you like to create this directory now and restore the file?", destinationLocation), "Confirmation", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {
                        //create the Sticky Notes directory (C:\users\*username*\AppData\Roaming\Microsoft\Sticky Notes\)
                        Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\Microsoft\Sticky Notes\");
                        //restore the Sticky Notes file
                        DirectoryInfo target = new DirectoryInfo(destinationLocation);
                        fileCount = source.GetFiles("*", SearchOption.AllDirectories).Length;
                        totalFileCount = fileCount;
                        int total = totalFileCount; //total things being transferred
                        for (int i = 0; i <= total; i++) //report those numbers
                        {
                            System.Threading.Thread.Sleep(100);
                            int percents = (i * 100) / total;
                            restoreAdditionalBgWorker.ReportProgress(percents, i);
                            //2 arguments:
                            //1. procenteges (from 0 t0 100) - i do a calcumation 
                            //2. some current value!
                        }
                        RestoreFilesRecursively(source, target);
                    }
                    else if (result == DialogResult.No)
                    {
                        restoreAdditionalBgWorker.CancelAsync();
                    }
                }
                else
                {

                }

            }
            else
            {
                DirectoryInfo target = new DirectoryInfo(destinationLocation);
                fileCount = source.GetFiles("*", SearchOption.AllDirectories).Length;
                totalFileCount = fileCount;
                int total = totalFileCount; //total things being transferred
                for (int i = 0; i <= total; i++) //report those numbers
                {
                    System.Threading.Thread.Sleep(100);
                    int percents = (i * 100) / total;
                    restoreAdditionalBgWorker.ReportProgress(percents, i);
                    //2 arguments:
                    //1. procenteges (from 0 t0 100) - i do a calcumation 
                    //2. some current value!
                }

                RestoreFilesRecursively(source, target);
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
            foreach (string subdirectory in subdirectoryEntries)
            {
                if (Directory.Exists(backupName))
                {
                    backupFoundLabel.Text = "(Backup Found!)";
                    backupFoundLabel.ForeColor = System.Drawing.Color.ForestGreen;

                }
                else
                {
                    backupFoundLabel.Text = "(No Backup Found!)";
                    backupFoundLabel.ForeColor = System.Drawing.Color.Crimson;
                }
            }
        }
        /// <summary>
        /// Restores from backup.
        /// </summary>
        /// <param name="directoryName">Name of the directory.</param>
        /// <param name="userName">Name of the user.</param>
        public void restoreFromBackup(string directoryName, string userName)
        {
            string userDirectory;
            string[] directoryList =
            {
                "Documents",
                "Favorites",
                "Desktop",
                "StickyNotes",
                "Pictures",
                "Videos",
                "Music"
            };
            foreach (string subDirectory in directoryList)
            {
                userDirectory = backupName + "\\" + subDirectory;
                if (Directory.Exists(userDirectory))
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
                        case "StickyNotes":
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
                if (userHasBeenSetup())
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
                if (userHasBeenBackedUp())
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
        public bool userHasBeenSetup()
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
        public bool userHasBeenBackedUp()
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
        /// Checks to see if the user has access to get quicken.
        /// </summary>
        /// <returns>Bool</returns>
        public bool doesUserGetQuicken()
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
        /// Checks the backup directories and updates the labels.
        /// </summary>
        /// <param name="backupName">Name of the backup.</param>
        public void checkBackupDirectories(string backupName)
        {
            string subDirectoryURL;
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
                subDirectoryURL = backupName + "\\" + subDirectory;
                if (Directory.Exists(subDirectoryURL))
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
                }

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

    }
}
