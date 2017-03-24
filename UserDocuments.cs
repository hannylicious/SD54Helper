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
    class UserDocuments
    {
        public void backupDocuments(MainForm mainForm, string userName, string selectedDrive, string userToBeBackedUp, string backupDirectoryName)
        {
            string documentsFolder = "";
            /*DOCUMENTS*/
            string destinationLocation = selectedDrive + "\\54HelperBackups\\" + backupDirectoryName + "\\Documents\\";
            //Backup the logged in users Documents
            if (userToBeBackedUp == userName)
            {
                documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            if (userToBeBackedUp != userName)
            //backup the selected users Desktop
            {
                documentsFolder = "C:\\Users\\" + userToBeBackedUp + "\\Documents\\";
            }
            DirectoryInfo source = new DirectoryInfo(documentsFolder);
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
                    @"C:\Users\" + userToBeBackedUp + @"\Documents\My Pictures",
                    @"C:\Users\" + userToBeBackedUp + @"\Documents\My Videos",
                    @"C:\Users\" + userToBeBackedUp + @"\Documents\My Music",
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
            mainForm.CopyFilesRecursively(documentsSource, target);
        }
        public void restoreDocuments(MainForm mainForm, string selectedDrive, string backupToRestore)
        {
            string selectedBackup = backupToRestore;
            string backupDrive = selectedDrive;
            string selectedBackupUsername = selectedBackup.Split('-')[0];
            string destinationLocation = @"C:\Users\" + selectedBackupUsername + "\\Documents\\";
            string documentsBackupFolder = backupDrive + "54HelperBackups\\" + selectedBackup + "\\Documents\\";
            DirectoryInfo source = new DirectoryInfo(documentsBackupFolder);
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
