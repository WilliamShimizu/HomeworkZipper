using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ionic.Zip;

namespace HomeworkZipper
{
    public partial class Form1 : Form
    {
        private const string BASE_DIR = @"C:\Users\Will Conrardy\Documents\GitHub\CSS501";
        private string DESKTOP = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        private string TEST_DIR;
        private const string VS_2010_LOCATION = @"C:\Program Files (x86)\Microsoft Visual Studio 10.0\Common7\IDE\devenv.exe";

        public Form1()
        {
            TEST_DIR = DESKTOP + "\\TEST";
            InitializeComponent();
            populateTreeView();
        }

        
        /// <summary>
        /// Populates the TreeView with projects based on the base directory of projects.
        /// </summary>
        private void populateTreeView()
        {
            treeView1.Nodes.Clear();
            foreach (string dir in Directory.GetDirectories(BASE_DIR))
            {
                if (dir.Contains(".git")) continue;
                treeView1.Nodes.Add(getLastString(dir)).Name = dir;
                foreach (string zipFile in Directory.GetFiles(dir, "*.zip"))
                {
                    treeView1.Nodes[dir].Nodes.Add(getLastString(zipFile)).Name = zipFile;
                }
            }
        }

        /// <summary>
        /// Gets the file name or directory name at the end of an absolute directory or file path.
        /// </summary>
        /// <param name="dir">The absolute directory or absolute file path</param>
        /// <returns></returns>
        private string getLastString(string dir)
        {
            return dir.Remove(0, dir.LastIndexOf("\\") + 1);
        }

        /// <summary>
        /// Gets the package name based on an absolute file path and the parent project. Necessary for creating the zip file archive directory structure.
        /// </summary>
        /// <param name="projectName">Example: Homework1</param>
        /// <param name="fileName">Example: C:\dev\Homework1\Package1\main.cpp</param>
        /// <returns></returns>
        private string getPackageName(string projectName, string fileName)
        {
            int start = fileName.IndexOf(projectName) + projectName.Length + 1;
            return fileName.Substring(start, fileName.LastIndexOf("\\") - start);
        }

        /// <summary>
        /// Validates that the main Zip It method can proceed.
        /// </summary>
        /// <returns></returns>
        private bool validateZipItClick()
        {
            if (treeView1.SelectedNode.Text.Contains(".zip"))
            {
                MessageBox.Show("Click on a Node with a project name to execute this command.");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Validates that the main Test It method can proceed.
        /// </summary>
        /// <returns></returns>
        private bool validateTestItClick()
        {
            if (treeView1.SelectedNode.Text.Contains(".zip")) return true;
            MessageBox.Show("Click on a Node representing a zip file to execute this command.");
            return false;
        }

        /// <summary>
        /// Click event for "Zip It". Creates a zip in the base project directory containing all necessary files.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void zipIt_Click(object sender, EventArgs e)
        {
            if (!validateZipItClick()) return;
            string dir = treeView1.SelectedNode.Name.ToString();
            string projectName = treeView1.SelectedNode.Text;
            using (ZipFile zip = new ZipFile())
            {
                zip.AddFile(dir + "\\" + projectName + ".sln", projectName);
                foreach (string file in getListOfImportantFiles(dir))
                {
                    zip.AddFile(file, projectName + "\\" + getPackageName(projectName, file));
                }
                zip.Save(getSubmissionZipName(dir, projectName));
            }
            populateTreeView();
            treeView1.Nodes[dir].Expand();
        }

        /// <summary>
        /// Gets a name for a zip file in case a zip file already exists. Increments in the form projectName\projectName_submission_1.zip.
        /// </summary>
        /// <param name="dir">Directory of the main project</param>
        /// <param name="projectName">The name of the project</param>
        /// <returns></returns>
        private string getSubmissionZipName(string dir, string projectName)
        {
            int subNum = 1;
            string newDir = dir + "\\" + projectName + "_submission_" + subNum + ".zip";
            while (File.Exists(newDir))
            {
                subNum++;
                newDir = dir + "\\" + projectName + "_submission_" + subNum + ".zip";
            }
            return newDir;
        }

        /// <summary>
        /// Gets the list of important files in all packages (excluding .sln).
        /// Directories that are Release, Debug, or ipch will not be added.
        /// Files that don't end in .filter, or .user, or the ReadMe.txt file will not be included.
        /// </summary>
        /// <param name="dir">The Base directory of the project</param>
        /// <returns></returns>
        private List<string> getListOfImportantFiles(string dir)
        {
            List<string> files = new List<string>();
            foreach (string directory in Directory.EnumerateDirectories(dir))
            {
                if (directory.Contains("Debug") || directory.Contains("Release") || directory.Contains("ipch")) continue;
                foreach (string file in Directory.EnumerateFiles(directory).Where(s => !s.EndsWith(".filters") && !s.EndsWith(".user")))
                {
                    if (file.Contains("ReadMe.txt") || file.Contains("Thumbs.db")) continue;
                    files.Add(file);
                }
            }
            return files;
        }

        /// <summary>
        /// Extracts the zipped file represented by the selected .zip node to Desktop\TEST (deletes this directory if it exists).
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void testIt_Click(object sender, EventArgs e)
        {
            if (!validateTestItClick()) return;
            extractToTestDir();
            string projName = treeView1.SelectedNode.Text.Substring(0, treeView1.SelectedNode.Text.IndexOf("_"));
            runWithVS2010(projName);
        }

        /// <summary>
        /// Starts VS 2010 given the project name for the .sln file (assumes extractToTestDir() has already been called).
        /// VS 2010 will automatically exit when the program exits.
        /// </summary>
        /// <param name="projectName">Name of the project for the .sln file.</param>
        private void runWithVS2010(string projectName)
        {
            if (!Directory.Exists(TEST_DIR)) return;
            string sln = "\"" + TEST_DIR + "\\" + projectName + "\\" + projectName + ".sln\"";
            Process process = new Process();
            process.StartInfo.FileName = VS_2010_LOCATION;
            process.StartInfo.Arguments = "/runexit " + sln;
            process.Start();
            process.WaitForExit();
            Directory.Delete(TEST_DIR, true);
        }

        /// <summary>
        /// Extracts the zip file represented by the selected node to Desktop\TEST (deletes this directory if it already exists.
        /// </summary>
        private void extractToTestDir()
        {
            using (ZipFile zip = new ZipFile(treeView1.SelectedNode.Name.ToString()))
            {
                if (Directory.Exists(TEST_DIR))
                {
                    Directory.Delete(TEST_DIR, true);
                    Thread.Sleep(1000);
                }
                zip.ExtractAll(TEST_DIR + "\\");
            }
        }
    }
}
