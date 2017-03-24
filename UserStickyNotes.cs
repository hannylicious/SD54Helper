using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Helpdesk54
{
    class UserStickyNotes
    {
        public void backupStickyNotes(MainForm mainForm, string userName, string selectedDrive, string userToBeBackedUp, string backupDirectoryName)
        {
            string stickyNotesFolder = "";
            string destinationLocation = selectedDrive + "\\54HelperBackups\\" + backupDirectoryName + "\\Sticky Notes\\";
            if (MainForm.WinMajorVersion == 10)
            //windows 10
            {
                stickyNotesFolder = "C:\\Users\\" + userToBeBackedUp + "\\AppData\\Local\\Packages\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe";
            }
            else
            {
                //windows 7
                stickyNotesFolder = "C:\\Users\\" + userToBeBackedUp + "\\AppData\\Roaming\\Microsoft\\Sticky Notes";
            }
            DirectoryInfo source = new DirectoryInfo(stickyNotesFolder);
            DirectoryInfo target = new DirectoryInfo(destinationLocation);
            int fileCount = source.GetFiles("*", SearchOption.AllDirectories).Length;
            int totalFileCount = fileCount;
            int total = totalFileCount; //total things being transferred
            for (int i = 0; i <= total; i++) //report those numbers
            {
                System.Threading.Thread.Sleep(100);
                int percents = (i * 100) / total;
                mainForm.additionalBgWorker.ReportProgress(percents, i);
                //2 arguments:
                //1. procenteges (from 0 t0 100) - i do a calcumation 
                //2. some current value!
            }
            if (!Directory.Exists(destinationLocation))
            {
                Directory.CreateDirectory(destinationLocation);
            }
            mainForm.CopyFilesRecursively(source, target);

        }
        public void restoreStickyNotes(MainForm mainForm, string selectedDrive, string backupToRestore)
        {
            string selectedBackup = backupToRestore;
            string backupDrive = selectedDrive;
            string selectedBackupUsername = selectedBackup.Split('-')[0];
            string destinationLocation = "";
            string stickyNotesBackupFolder = "";
            //set the destination to stickynotes (C:\users\*username*\AppData\Roaming\Microsoft\Sticky Notes\)
            if (MainForm.WinMajorVersion == 10)
            //windows 10
            {
                destinationLocation = "C:\\Users\\" + selectedBackupUsername + "\\AppData\\Local\\Packages\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe";
            }
            else
            {
                //windows 7
                destinationLocation = "C:\\Users\\" + selectedBackupUsername + "\\AppData\\Roaming\\Microsoft\\Sticky Notes";
            }
            //By default - Sticky Notes is really the only file that might fail here - the rest are system folders which will always exist
            if (!Directory.Exists(destinationLocation))
            {
                DialogResult result = MessageBox.Show(String.Format("The directory {0} does not exist. Would you like to create this directory now and restore the file?", destinationLocation), "Confirmation", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    //create the Sticky Notes directory (C:\users\*username*\AppData\Roaming\Microsoft\Sticky Notes\)
                    if (MainForm.WinMajorVersion == 10)
                    //windows 10
                    {
                        Directory.CreateDirectory("C:\\Users\\" + selectedBackupUsername + "\\AppData\\Local\\Packages\\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe");
                    }
                    else
                    {
                        //windows 7
                        Directory.CreateDirectory("C:\\Users\\" + selectedBackupUsername + "\\AppData\\Roaming\\Microsoft\\Sticky Notes");
                    }

                    //restore the Sticky Notes file
                    DirectoryInfo source = new DirectoryInfo(stickyNotesBackupFolder);
                    DirectoryInfo target = new DirectoryInfo(destinationLocation);
                    int fileCount = source.GetFiles("*", SearchOption.AllDirectories).Length;
                    int totalFileCount = fileCount;
                    int total = totalFileCount; //total things being transferred
                    for (int i = 0; i <= total; i++) //report those numbers
                    {
                        System.Threading.Thread.Sleep(100);
                        int percents = (i * 100) / total;
                        mainForm.restoreAdditionalBgWorker.ReportProgress(percents, i);
                        //2 arguments:
                        //1. procenteges (from 0 t0 100) - i do a calcumation 
                        //2. some current value!
                    }
                    mainForm.RestoreFilesRecursively(source, target);
                }
                else if (result == DialogResult.No)
                {
                    mainForm.restoreAdditionalBgWorker.CancelAsync();
                }
            }
            else
            {
                //restore the Sticky Notes file
                stickyNotesBackupFolder = backupDrive + "54HelperBackups\\" + selectedBackup + "\\Sticky Notes\\";
                DirectoryInfo source = new DirectoryInfo(stickyNotesBackupFolder);
                DirectoryInfo target = new DirectoryInfo(destinationLocation);
                int fileCount = source.GetFiles("*", SearchOption.AllDirectories).Length;
                int totalFileCount = fileCount;
                int total = totalFileCount; //total things being transferred
                for (int i = 0; i <= total; i++) //report those numbers
                {
                    System.Threading.Thread.Sleep(100);
                    int percents = (i * 100) / total;
                    mainForm.restoreAdditionalBgWorker.ReportProgress(percents, i);
                    //2 arguments:
                    //1. procenteges (from 0 t0 100) - i do a calcumation 
                    //2. some current value!
                }
                mainForm.RestoreFilesRecursively(source, target);
            }
        }
    }
}
