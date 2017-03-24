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
    class UserFavorites
    {
        public void backupFavorites(MainForm mainForm, string userName, string selectedDrive, string userToBeBackedUp, string backupDirectoryName)
        {
            string favoritesFolder = "";
            /*FAVORITES*/
            string destinationLocation = selectedDrive + "\\54HelperBackups\\" + backupDirectoryName + "\\Favorites\\";
            //Backup the logged in users Favorites
            if (userToBeBackedUp == userName)
            {
                favoritesFolder = Environment.GetFolderPath(Environment.SpecialFolder.Favorites);
            }
            if (userToBeBackedUp != userName)
            //backup the selected users Favorites
            {
                favoritesFolder = "C:\\Users\\" + userToBeBackedUp + "\\Favorites\\";
            }
            DirectoryInfo source = new DirectoryInfo(favoritesFolder);
            DirectoryInfo target = new DirectoryInfo(destinationLocation);
            DirectoryInfo favoritesSource = new DirectoryInfo(favoritesFolder);
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
            mainForm.CopyFilesRecursively(favoritesSource, target);
        }
        public void restoreFavorites(MainForm mainForm, string selectedDrive, string backupToRestore)
        {
            string selectedBackup = backupToRestore;
            string backupDrive = selectedDrive;
            string selectedBackupUsername = selectedBackup.Split('-')[0];
            string destinationLocation = @"C:\Users\" + selectedBackupUsername + "\\Favorites\\";
            string favoritesBackupFolder = backupDrive + "54HelperBackups\\" + selectedBackup + "\\Favorites\\";
            DirectoryInfo source = new DirectoryInfo(favoritesBackupFolder);
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
