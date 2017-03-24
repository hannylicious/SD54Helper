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
    public class UserDesktop
    {
        public void backupDesktop(MainForm mainForm, string userName, string selectedDrive, string userToBeBackedUp, string backupDirectoryName)
        {
            string desktopFolder = "";
            string destinationLocation = selectedDrive + "\\54HelperBackups\\" + backupDirectoryName + "\\Desktop\\";
            //Backup the logged in users Desktop
            if (userToBeBackedUp == userName)
            {
                desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            if (userToBeBackedUp != userName)
            //backup the selected users Desktop
            {
                desktopFolder = "C:\\Users\\" + userToBeBackedUp + "\\Desktop\\";
            }
            DirectoryInfo source = new DirectoryInfo(desktopFolder);
            DirectoryInfo desktopSource = new DirectoryInfo(desktopFolder);
            DirectoryInfo target = new DirectoryInfo(destinationLocation);
            int fileCount = 0;
            try
            {
                fileCount = source.GetFiles("*", SearchOption.AllDirectories).Length;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            int totalFileCount = fileCount;
            int total = totalFileCount; //total things being transferred
            for (int i = 0; i <= total; i++) //report those numbers
            {
                System.Threading.Thread.Sleep(100);
                int percents = (i * 100) / total;
                mainForm.essentialBgWorker.ReportProgress(percents, i);
                //2 arguments:
                //1. procenteges (from 0 t0 100) - i do a calcumation 
                //2. some current value!
            }
            if (!Directory.Exists(destinationLocation))
            {
                Directory.CreateDirectory(destinationLocation);
            }
            mainForm.CopyFilesRecursively(desktopSource, target);            
        }
        public void restoreDesktop(MainForm mainForm, string selectedDrive, string backupToRestore)
        {
            string selectedBackup = backupToRestore;
            string backupDrive = selectedDrive;
            string selectedBackupUsername = selectedBackup.Split('-')[0];
            string destinationLocation = @"C:\Users\" + selectedBackupUsername + "\\Desktop\\";
            string desktopBackupFolder = backupDrive + "54HelperBackups\\" + selectedBackup + "\\Desktop\\";
            DirectoryInfo source = new DirectoryInfo(desktopBackupFolder);
            DirectoryInfo target = new DirectoryInfo(destinationLocation);
            int fileCount = source.GetFiles("*", SearchOption.AllDirectories).Length;
            int totalFileCount = fileCount;
            int total = totalFileCount; //total things being transferred
            for (int i = 0; i <= total; i++) //report those numbers
            {
                System.Threading.Thread.Sleep(100);
                int percents = (i * 100) / total;
                mainForm.restoreEssentialBgWorker.ReportProgress(percents, i);
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
