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
    class UserVideos
    {
        public void backupVideos(MainForm mainForm, string userName, string selectedDrive, string userToBeBackedUp, string backupDirectoryName)
        {
            string videosFolder = "";
            string pathWithEvn = @"%USERPROFILE%\Videos";
            string destinationLocation = selectedDrive + "\\54HelperBackups\\" + backupDirectoryName + "\\Videos\\";
            videosFolder = "";
            //Backup the logged in users Desktop
            if (userToBeBackedUp == userName)
            {
                videosFolder = Environment.ExpandEnvironmentVariables(pathWithEvn);
            }
            if (userToBeBackedUp != userName)
            //backup the selected users Desktop
            {
                videosFolder = "C:\\Users\\" + userToBeBackedUp + "\\Videos\\";
            }
            DirectoryInfo source = new DirectoryInfo(videosFolder);
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
        public void restoreVideos(MainForm mainForm, string selectedDrive, string backupToRestore)
        {
            string selectedBackup = backupToRestore;
            string backupDrive = selectedDrive;
            string selectedBackupUsername = selectedBackup.Split('-')[0];
            string destinationLocation = @"C:\Users\" + selectedBackupUsername + "\\Videos\\";
            string videosBackupFolder = backupDrive + "54HelperBackups\\" + selectedBackup + "\\Videos\\";
            DirectoryInfo source = new DirectoryInfo(videosBackupFolder);
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
