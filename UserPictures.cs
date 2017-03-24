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
    class UserPictures
    {
        public void backupPictures(MainForm mainForm, string userName, string selectedDrive, string userToBeBackedUp, string backupDirectoryName)
        {
            string picturesFolder = "";
            string destinationLocation = selectedDrive + "\\54HelperBackups\\" + backupDirectoryName + "\\Pictures\\";  
            //Backup the logged in users pictures         
            if (userToBeBackedUp == userName)
            {
                picturesFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            }
            //Backup the selected users pictures
            if (userToBeBackedUp != userName)
            {
                picturesFolder = "C:\\Users\\" + userToBeBackedUp + "\\Pictures\\";
            }
            DirectoryInfo source = new DirectoryInfo(picturesFolder);
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
        public void restorePictures(MainForm mainForm, string selectedDrive, string backupToRestore)
        {
            string selectedBackup = backupToRestore;
            string backupDrive = selectedDrive;
            string selectedBackupUsername = selectedBackup.Split('-')[0];
            string destinationLocation = @"C:\Users\" + selectedBackupUsername + "\\Pictures\\";
            string picturesBackupFolder = backupDrive + "54HelperBackups\\" + selectedBackup + "\\Pictures\\";
            DirectoryInfo source = new DirectoryInfo(picturesBackupFolder);
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
            if (!Directory.Exists(destinationLocation))
            {
                MessageBox.Show("This location does not exist for the logged in user.");
            }
            mainForm.RestoreFilesRecursively(source, target);
        }
    }
}
