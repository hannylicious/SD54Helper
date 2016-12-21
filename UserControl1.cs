using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Management;
using System.Linq;
using System.Runtime.InteropServices;
using IWshRuntimeLibrary;

namespace Helpdesk54
{
    public partial class UserControl1 : Form

    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "SD54Helper";
        int existingSetupRowNumber, existingBackupRowNumber; 

        string backupDirectoryName;
        string backupName;
        string serverName;
        string homeDirectory;
        string desktopFolder, documentsFolder, favoritesFolder;
        string userName;
        BackgroundWorker essentialBgWorker = new BackgroundWorker();
        BackgroundWorker additionalBgWorker = new BackgroundWorker();
        BackgroundWorker restoreEssentialBgWorker = new BackgroundWorker();
        BackgroundWorker restoreAdditionalBgWorker = new BackgroundWorker();
        string selectedDrive;
        string clickedButton;
        string itemsChanged;
        int totalFileCount;
        string destinationLocation;
        DirectoryInfo source;
        int fileCount;

        public UserControl1()
        {
            InitializeComponent();
            //set username - Do this early!
            userName = Environment.UserName;
            usernameLabel.Text = userName;
            backupDirectoryName = usernameLabel.Text.ToString() + "-Backups-" + DateTime.Now.Year.ToString();
            //check if secretary needs quicken
            if (doesSecretaryGetQuicken())
            {
                quickenButton.Enabled = true;
                quickenCheckBox.Enabled = true;
                quickenBackupCheckBox.Enabled = true;
            } else {
                quickenButton.Enabled = false;
                quickenCheckBox.Enabled = false;
                quickenBackupCheckBox.Enabled = false;
            }
            //set the servernamelink
            DriveInfo[] theDrives = DriveInfo.GetDrives();
            foreach (DriveInfo currentDrive in theDrives)
            {
                if (currentDrive.DriveType == DriveType.Network)
                {
                    string currentDriveString = currentDrive.Name.ToString();
                    string path = GetUNCPath(currentDriveString);
                    if (path.ToLower().Contains(userName))
                    {
                        // This is the Home drive! // 
                        homeDirectory = currentDrive.Name;
                        
                    }
                    Uri uri = new Uri(path);
                    serverName = uri.Host.ToString();
                    serverNameLinkLabel.Text = serverName;
                }
            }
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
            //*****
            //set backupDriveCombo to dropdown
            //*****
            backupDriveCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            foreach (DriveInfo currentDrive in theDrives)
            {
                if (currentDrive.IsReady == true)
                {
                    backupDriveCombo.Items.Add(currentDrive);
                }
            }

            //Try to set the selected drive to the H:/ Drive
            try
            {
                backupDriveCombo.SelectedIndex = backupDriveCombo.FindString(homeDirectory);
            }
            catch (IOException e)
            {
                Console.WriteLine("IOException source: {0}", e.Source);

                throw;
            }
            //Set the selected drive freespace label
            try
            {
                DriveInfo selectedDrive = (DriveInfo)backupDriveCombo.SelectedItem;
                long driveSpace = selectedDrive.AvailableFreeSpace;
                string driveFreeSpace = FormatBytes(driveSpace);
                backupDriveLabel.Text = driveFreeSpace + " Free";
            }
            catch (IOException e)
            {
                Console.WriteLine("IOException source: {0}", e.Source);

                throw;
            }
            //*****
            //set restoreDriveCombo to dropdown
            //*****
            restoreDriveCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            foreach (DriveInfo currentDrive in theDrives)
            {
                if (currentDrive.IsReady == true)
                {
                    restoreDriveCombo.Items.Add(currentDrive);
                }
            }

            //Try to set the selected drive to the H:/ Drive
            try
            {
                restoreDriveCombo.SelectedIndex = restoreDriveCombo.FindString(homeDirectory);
            }
            catch (IOException e)
            {
                Console.WriteLine("IOException source: {0}", e.Source);

                throw;
            }

            //Check the H:\ drive for a backup
            checkForExistingBackup();
            getFilesToListView();

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
                    //System.Windows.Forms.MessageBox.Show(folder);
                    long folderSize = DirSize(new DirectoryInfo(folder));
                    string folderMB = FormatBytes(folderSize);
                    videosSizeLabel.Text = folderMB;
                }
                else //iterate through all the other directory Locations
                {
                    var dir = (Environment.SpecialFolder)Enum.Parse(typeof(Environment.SpecialFolder), directoryLocation);
                    string folder = Environment.GetFolderPath(dir);
                    long folderSize = DirSize(new DirectoryInfo(folder));
                    string folderMB = FormatBytes(folderSize);
                    switch (directoryLocation)
                    {
                        case "DesktopDirectory":
                            desktopSizeLabel.Text = folderMB;
                            break;
                        case "MyDocuments":
                            documentsSizeLabel.Text = folderMB;
                            break;
                        case "Favorites":
                            favoritesSizeLabel.Text = folderMB;
                            break;
                        case "MyMusic":
                            musicSizeLabel.Text = folderMB;
                            break;
                        case "MyPictures":
                            picturesSizeLabel.Text = folderMB;
                            break;
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
                string stickyNotesMB = FormatBytes(stickyNotesSize);
                stickyNotesSizeLabel.Text = stickyNotesMB;
            }
            else //haven't launched sticky notes
            {
                stickyNotesSizeLabel.Text = "N/A";
            }

        }
        //set all the files in the server location to the serverListView tab
        public void getFilesToListView()
        {
/*
            string[] files = System.IO.Directory.GetFiles(serverName);

            for (int x = 0; x < files.Length; x++)
            {
                serverListView.Items.Add(files[x]);
            }
*/
        }
        //handle clicking of serverNameLinkLabel
        private void serverNameLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            serverNameLinkLabel.LinkVisited = true;
            System.Diagnostics.Process.Start("explorer", @"\\" + serverNameLinkLabel.Text.ToString());
        }
        /*Check for existing backups on the H: drive*/
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
                    catch (IOException e)
                    {
                        Console.WriteLine("IOException source: {0}", e.Source);

                        throw;
                    }
        }
        /*
         * ***** CHECK TO SEE IF APPLICATIONS ARE INSTALLED ***** *
         */
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
        /*
        * ***** *
        * Open software for configuration
        * ***** *
        */
        void openSoftwareClick(object sender, EventArgs e)
        {
            string path;
            Button sentButton = sender as Button;
            switch (sentButton.Text.ToString())
            {
                case "Open Outlook":
                    path = @"C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE";
                    if (System.IO.File.Exists(path))
                    {
                        System.Diagnostics.Process.Start(path);
                    }
                    else
                    {
                        System.Diagnostics.Process.Start(@"C:\Program Files\Microsoft Office\Office14\OUTLOOK.EXE");
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
        /*
        * ***** *
        * Go to \\dataserver02 and begin the installers for the software
        * ***** *
        */
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
        /*
        * ***** *
        * Format the bytes string into appropriately chosen string
        * ***** *
        */
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
        /*
        * ***** *
        * Create Shortcut URL's on users Desktop *
        * ***** *
        */
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
        /*
        * ***** *
        * Get Directory Size *
        * ***** *
        */
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
                shortcut.IconLocation = @"\\dataserver02\PCUpdate\54Helper\icons\ico\campus.ico";           // The icon of the shortcut
                shortcut.Save();                                    // Save the shortcut
            }
            if (shortcutName == "AESOP")
            {
                shortcut.TargetPath = targetFileLocation;                 // The path of the file that will launch when the shortcut is run
                shortcut.Arguments = @"https://www.aesoponline.com/login2.asp";
                shortcut.IconLocation = @"\\dataserver02\PCUpdate\54Helper\icons\ico\aesop.ico";           // The icon of the shortcut
                shortcut.Save();
            }
            if (shortcutName == "E-Finance")
            {
                shortcut.TargetPath = targetFileLocation;                 // The path of the file that will launch when the shortcut is run
                shortcut.Arguments = @"https://efinance.sd54.org/gas2.50/wa/r/plus/finplus51";
                shortcut.IconLocation = @"\\dataserver02\PCUpdate\54Helper\icons\ico\efinance.ico";           // The icon of the shortcut
                shortcut.Save();
            }
            else {
                shortcut.TargetPath = targetFileLocation;                 // The path of the file that will launch when the shortcut is run
                shortcut.Save();                                    // Save the shortcut
            }

        }
        /* ***** *
         * Install Shortcuts Button
         * ***** *
         */
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
                        CreateShortcut("Microsoft Word", Environment.GetFolderPath(Environment.SpecialFolder.Desktop), @"C:\Program Files (x86)\Microsoft Office\Office14\WINWORD.exe");
                        wordShortcutCheckBox.Checked = true;
                        break;
                }
            }
        }
        /* ***** *
        * Drive selection dropdown changes *
        * ***** *
        */
        private void backupDriveCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Set the selected drive freespace label
            DriveInfo selectedDrive = (DriveInfo)backupDriveCombo.SelectedItem;
            long driveSpace = selectedDrive.AvailableFreeSpace;
            string driveFreeSpace = FormatBytes(driveSpace);
            backupDriveLabel.Text = driveFreeSpace + " Free";
        }
        private void restoreDriveCombo_SelectedIndexChanged(object sender, EventArgs e)
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
/* ***** *
* ***** * ***** * ***** *
* ***** * ***** * ***** *
* BACKUP BUTTONS SECTION
* ***** * ***** * ***** *
* ***** * ***** * ***** *
* ***** */
/*
         * ***** DESKTOP BACKUP ***** *
        */
        private void backupDesktopButton_Click(object sender, EventArgs e)
        {
            desktopBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            essentialBgWorker.RunWorkerAsync();
        }
        /*
         * ***** DOCUMENTS BACKUP ***** *
        */
        private void backupDocumentsButton_Click(object sender, EventArgs e)
        {
            documentsBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            essentialBgWorker.RunWorkerAsync();
        }
        /*
         * ***** FAVORITES BACKUP ***** *
        */
        private void backupFavoritesButton_Click(object sender, EventArgs e)
        {
            favoritesBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            essentialBgWorker.RunWorkerAsync();
        }
        /*
         * ***** DOCUMENTS, FAVORITES & DESKTOP BACKUP ***** *
        */
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
        /* 
         * *****
         * essentialBgWorker Do Work Method *
         * *****
        */
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
        /* 
         * *****
         * essentialBgWorker Process Changed *
         * *****
        */
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
        /* 
         * *****
         * essentialBgWorker Completed *
         * *****
        */
        void essentialBgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //do the code when bgv completes its work
            essentialItemsProgressBar.Visible = false;
            essentialItemsProgressLabel.Visible = true;
            essentialItemsProgressLabel.Text = String.Format("Progress: 100% - All {0} Files Transferred", itemsChanged);
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
        /*
         * ***** STICKY NOTES BACKUP ***** *
        */
        private void backupStickyNotesButton_Click(object sender, EventArgs e)
        {
            stickyNotesBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            additionalBgWorker.RunWorkerAsync();
        }
        /*
         * ***** PICTURES BACKUP ***** *
        */
        private void backupPicturesButton_Click(object sender, EventArgs e)
        {
            picturesBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            additionalBgWorker.RunWorkerAsync();
        }
        /*
         * ***** VIDEOS BACKUP ***** *
        */
        private void backupVideosButton_Click(object sender, EventArgs e)
        {
            videosBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            additionalBgWorker.RunWorkerAsync();
        }
        /*
         * ***** MUSIC BACKUP ***** *
        */
        private void backupMusicButton_Click(object sender, EventArgs e)
        {
            musicBackupCheckBox.Checked = true;
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            additionalBgWorker.RunWorkerAsync();
        }
        /* 
         * *****
         * additionalBgWorker Do Work Method *
         * *****
        */
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
        /* 
         * *****
         * additionalBgWorker Process Changed *
         * *****
        */
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
        /* 
         * *****
         * additionalBgWorker Completed *
         * *****
        */
        void additionalBgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //do the code when bgv completes its work
            additionalItemsProgressBar.Visible = false;
            additionalItemsProgressLabel.Visible = true;
            additionalItemsProgressLabel.Text = String.Format("Progress: 100% - All {0} Files Transferred", itemsChanged);
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
        /*
            * ***** *
            * Backup Desktop
            * ***** *
        */
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
        /*
            * ***** *
            * Backup Favorites
            * ***** *
        */
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
        /*
            * ***** *
            * Backup My Documents
            * ***** *
        */
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


        /*
         * ***** *
         * Copy all directories and files recursively
         * ***** *
        */
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
        /*
         * ***** DESKTOP RESTORE ***** *
        */
        private void restoreDesktopButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            selectedDrive = restoreDriveCombo.SelectedItem.ToString();
            restoreEssentialBgWorker.RunWorkerAsync();
        }
        /*
         * ***** DOCUMENTS RESTORE ***** *
        */
        private void restoreDocumentsButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            restoreEssentialBgWorker.RunWorkerAsync();
        }
        /*
         * ***** FAVORITES RESTORE ***** *
        */
        private void restoreFavoritesButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            selectedDrive = backupDriveCombo.SelectedItem.ToString();
            restoreEssentialBgWorker.RunWorkerAsync();
        }
        /*
         * ***** DOCUMENTS, FAVORITES & DESKTOP RESTORE ***** *
        */
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
        /* 
         * *****
         * restoreWorker Do Work Method *
         * *****
        */
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
        /* 
         * *****
         * restoreWorker Process Changed *
         * *****
        */
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
        /* 
         * *****
         * restoreWorker Completed *
         * *****
        */
        void restoreEssentialBgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //do the code when bgv completes its work
            restoreEssentialsProgressBar.Visible = false;
            restoreEssentialsBarLabel.Visible = true;
            restoreEssentialsBarLabel.Text = String.Format("Progress: 100% - All {0} Files Restored", itemsChanged);
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
        /*
         * ***** STICKY NOTES RESTORE ***** *
        */
        private void restoreStickyNotesButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            selectedDrive = restoreDriveCombo.SelectedItem.ToString();
            restoreAdditionalBgWorker.RunWorkerAsync();
            restoreAdditionalBarLabel.Visible = false;
        }
        /*
         * ***** PICTURES RESTORE ***** *
        */
        private void restorePicturesButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            selectedDrive = restoreDriveCombo.SelectedItem.ToString();
            restoreAdditionalBgWorker.RunWorkerAsync();
            restoreAdditionalBarLabel.Visible = false;
        }
        /*
         * ***** VIDEOS RESTORE ***** *
        */
        private void restoreVideosButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            selectedDrive = restoreDriveCombo.SelectedItem.ToString();
            restoreAdditionalBgWorker.RunWorkerAsync();
            restoreAdditionalBarLabel.Visible = false;
        }
        /*
         * ***** MUSIC RESTORE ***** *
        */
        private void restoreMusicButton_Click(object sender, EventArgs e)
        {
            clickedButton = ((Button)sender).Name.ToString();
            itemsChanged = ((Button)sender).Text.ToString();
            selectedDrive = restoreDriveCombo.SelectedItem.ToString();
            restoreAdditionalBgWorker.RunWorkerAsync();
            restoreAdditionalBarLabel.Visible = false;
        }
        /* 
         * *****
         * restoreAdditionalBgWorker Do Work Method *
         * *****
        */
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
        /* 
         * *****
         * restoreAdditionalBgWorker Process Changed *
         * *****
        */
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
        /* 
         * *****
         * restoreWorker Completed *
         * *****
        */
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
        /*
         * ***** *
         * Restore all directories and files recursively
         * ***** *
        */
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
        /*
        * ***** CHECK REGISTRY TO SEE IF PROGRAM IS INSTALLED ***** *
        */
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
        /*
        * ***** CHECK REGISTRY TO SEE IF PROGRAM IS INSTALLED VIA MSI ISNTALLER ***** *
        */
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
                        return Path.Combine(sb.ToString().TrimEnd(), path);
                    }
                }
            }

            return originalPath;
        }


        /*
        * ***** *
        * ***** END GET UNC PATH ***** *
        * ***** *
        */

        /*
         * ***** *
         * SEARCH FOR BACKUP DIRECTORY
         * ***** *
         */
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

        private void userSetupCompleteButton_Click(object sender, EventArgs e)
        {
            //1YWi6rMUie3BD_AXH85Ew0TGJpgP_AVNN4Ury2bpIaQk - spreadsheet ID
            UserCredential credential;
            using (var stream =
                          new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);


                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            ValueRange valueRange = new ValueRange();
            var oblist = new List<object>() {  };

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
                String customRange = "A" + (x+1);
                SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId, customRange);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                UpdateValuesResponse result = update.Execute();
                MessageBox.Show("The users setup information has been updated.");
            } else {
                // Define request parameters.
                
                String range = "A1";
                SpreadsheetsResource.ValuesResource.AppendRequest update = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                AppendValuesResponse result = update.Execute();
                MessageBox.Show("The users setup information has been recorded.");
            }

        }

        private void userBackupCompleteButton_Click(object sender, EventArgs e)
        {
            //1XFsFJ2nqrXsxO9ShRJOwQ2MMlDpxf1wmU7RgSKUvmKM - spreadsheet ID
            UserCredential credential;
            using (var stream =
                          new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);


                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

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
                MessageBox.Show("The users backup information has been updated.");
            }
            else
            {
                // Define request parameters.
                String range = "A1";
                SpreadsheetsResource.ValuesResource.AppendRequest update = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                AppendValuesResponse result = update.Execute();
                MessageBox.Show("The users backup information has been recorded.");
            }
        }
        //Checks the Google Spreadsheet to see if the user exists in the current year as having been setup
        //If the user DOES exist - it updates the UserChecklist tab accordingly in the Setup section
        //returns bool
        public bool userHasBeenSetup()
        {
            UserCredential credential;
            using (var stream =
                          new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);


                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

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
        //Updates the user CheckBoxes for their setup if the user exists
        public void updateUserSetupChecks()
        {
            UserCredential credential;
            using (var stream =
                          new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);


                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

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
        public bool userHasBeenBackedUp()
        {
            UserCredential credential;
            using (var stream =
                          new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);


                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

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
        public void updateUserBackupChecks()
        {
            UserCredential credential;
            using (var stream =
                          new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);


                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

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
        public bool doesSecretaryGetQuicken()
        {
            UserCredential credential;
            using (var stream =
                          new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);


                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

            }
            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

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
