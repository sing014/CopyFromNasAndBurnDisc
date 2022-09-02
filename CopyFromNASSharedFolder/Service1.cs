using System;
using System.Xml;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Threading;
using Microsoft.VisualBasic.FileIO;
using System.Drawing;
using System.DirectoryServices.AccountManagement;
using System.Security.Principal;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using IMAPI2.Interop;
using IMAPI2.MediaItem;


namespace CopyFromNASSharedFolder
{
    public partial class Service1 : ServiceBase
    {
        const string csProjectID = "CopyFromNASSharedFolder";
        const string msDirectory = @"C:\Temp\CopyFromNASSharedFolder";

        double startTime = 13.5;     //start time
        double overDays = 180;       //超過的天數
        double dayTimeStamp = 1.0;

        string configFile = "CopyFromNASSharedFolder.conf";
        string srcNasDirectory = @"E:\NATS_Archive_A";
        string isoFilePath = "";
        string isoName = "";
        string isoFileDir = @"E:\ISO\";
        string strSrcDirPath = "";
        string copyToTargetDir = @"E:\TempOfShared";
        string strTargetDirPath = "";

        System.Threading.Timer t = null;


        string currentFolder = "";

        int debugCount = 0;
        bool isSucceed = false;
        int nasFolderCount = 0;

        //Burn variables 22/05/20
        private const string ClientName = "BurnMedia";

        Int64 _totalDiscSize;

        private bool _isBurning;
        private bool _isFormatting;
        private IMAPI_BURN_VERIFICATION_LEVEL _verificationLevel =
            IMAPI_BURN_VERIFICATION_LEVEL.IMAPI_BURN_VERIFICATION_NONE;
        private bool _closeMedia;
        private bool _ejectMedia;

        //MsftDiscRecorder2 discRecorder2 ;
        DirectoryItem directoryItem;

        IMediaItem mediaItem;
        MsftDiscRecorder2 discRecorder;
        MsftDiscFormat2Data discFormatData;
        MsftFileSystemImage fileSystemImage;

        private BurnData _burnData = new BurnData();

        private string burnDiscParameters = "";

        public Service1()
        {
            InitializeComponent();

        }

        protected override void OnStart(string[] args)
        {
            try
            {
                BuildDirectoryTemp();
                string sFile = Path.Combine(msDirectory,
                    $"{csProjectID}-OnStart.txt");
                string sMsg = string.Format($"{csProjectID}-OnStart, {DateTime.Now.ToString("HH:mm:ss")}");
                EventLog.WriteEntry(sMsg);

                string cFile = Path.Combine(msDirectory, configFile);

                using (StreamReader sr = new StreamReader(cFile, Encoding.UTF8))
                {
                    string line;
                    int rowNumber = 0;
                    while ((line = sr.ReadLine()) != null)
                    {
                        try
                        {
                            switch (rowNumber)
                            {
                                case 0:
                                    startTime = Convert.ToDouble(line);
                                    rowNumber++;
                                    break;
                                case 1:
                                    srcNasDirectory = line;
                                    rowNumber++;
                                    break;
                                case 2:
                                    overDays = Convert.ToDouble(line);
                                    rowNumber++;
                                    break;
                                case 3:
                                    copyToTargetDir = line;
                                    rowNumber++;
                                    break;
                                case 4:
                                    dayTimeStamp = Convert.ToDouble(line);
                                    rowNumber++;
                                    break;

                            }
                        }
                        catch (Exception ex210630A)
                        {
                            WriteException(ex210630A);
                        }

                    }
                }

                using (StreamWriter sw = new StreamWriter(sFile, true, Encoding.UTF8))
                {
                    sw.WriteLine(sMsg);
                    sw.Close();
                }

                //開始處理   20/01/27
                setTaskAtFixedTime();

            }
            catch (Exception ex)
            {
                WriteException(ex);
            }




        }



        protected override void OnStop()
        {
            try
            {
                BuildDirectoryTemp();
                string sFile = Path.Combine(msDirectory,
                    $"{csProjectID}-OnStop.txt");
                string sMsg = string.Format($"{csProjectID}-OnStop, {DateTime.Now.ToString("HH:mm:ss")}");
                EventLog.WriteEntry(sMsg);
                using (StreamWriter sw = new StreamWriter(sFile, true, Encoding.UTF8))
                {
                    sw.WriteLine(sMsg);
                }
            }
            catch (Exception ex)
            {
                WriteException(ex);
            }
        }

        private void BuildDirectoryTemp()
        {
            if (!Directory.Exists(msDirectory))
            {
                Directory.CreateDirectory(msDirectory);
            }
        }

        private void BuildDirectoryForCopyFromNASSharedFolder(string directoryFullName)
        {
            if (!Directory.Exists(directoryFullName))
            {
                Directory.CreateDirectory(directoryFullName);
            }
        }



        private void WriteException(Exception ex)
        {
            BuildDirectoryTemp();

            const string csProjectID = "CopyFromNASSharedFolder";
            string sFile = Path.Combine(msDirectory, $"{csProjectID}-Exception.txt");
            string sMsg = string.Format(
                $"{csProjectID}-Exception, " +
                $"{DateTime.Now.ToString("HH:mm:ss")}, {ex.Message}");
            EventLog.WriteEntry(sMsg);
            using (StreamWriter sw = new StreamWriter(sFile, true, Encoding.UTF8))
            {
                sw.WriteLine(sMsg);
                sw.WriteLine(ex.Message);
                sw.WriteLine(ex.StackTrace);
            }
        }

        private void WriteExceptionExt(Exception ex, string exMessage)
        {
            BuildDirectoryTemp();

            const string csProjectID = "CopyFromNASSharedFolder";
            string sFile = Path.Combine(msDirectory, $"{csProjectID}-Exception.txt");
            string sMsg = string.Format(
                $"{csProjectID}-Exception, " +
                $"{DateTime.Now.ToString("HH:mm:ss")}, {ex.Message}, {exMessage}");
            EventLog.WriteEntry(sMsg);
            using (StreamWriter sw = new StreamWriter(sFile, true, Encoding.UTF8))
            {
                sw.WriteLine(sMsg);
                sw.WriteLine(ex.Message);
                sw.WriteLine(ex.StackTrace);
            }
        }

        private void WriteLog(string strMessage)
        {
            BuildDirectoryTemp();

            const string csProjectID = "CopyFromNASSharedFolder";
            string sFile = Path.Combine(msDirectory, $"{csProjectID}-DebugLog.txt");
            string sMsg = string.Format(
                $"{csProjectID}-DebugLog, " +
                $"{DateTime.Now.ToString("HH:mm:ss")}, {strMessage}");
            EventLog.WriteEntry(sMsg);
            using (StreamWriter sw = new StreamWriter(sFile, true, Encoding.UTF8))
            {
                sw.WriteLine(sMsg);
                sw.WriteLine(strMessage);
                //sw.WriteLine(ex.StackTrace);
            }
        }


        private void setTaskAtFixedTime()
        {
            DateTime now = DateTime.Now;

            if(t != null)
            {
                t.Dispose();
            }

            DateTime oneOClock = DateTime.Today.AddHours(startTime); //Run Time
            if (now > oneOClock)
            {
                oneOClock = oneOClock.AddDays(dayTimeStamp);
                //oneOClock = oneOClock.AddHours(1.0);
            }
            int msUntilFour = (int)(oneOClock - now).TotalMilliseconds;
            if (msUntilFour < -1)
            {
                msUntilFour = int.MaxValue;
            }
            
            t = new System.Threading.Timer(doAt10AM);

            //if (msUntilFour < 0)
            //{
            //    msUntilFour = -msUntilFour;
            //}

            t.Change(msUntilFour, Timeout.Infinite);
            
        }

        //要執行的任務
        private void doAt10AM(object state)
        {

            //bool isSucceed2 = false;
            //int nasFolderCount = 0;

            try
            {

                string[] dirz = Directory.GetDirectories(srcNasDirectory);

                //sort dirz[]
                DirectoryInfo di = new DirectoryInfo(srcNasDirectory);

                DirectoryInfo[] arrDir = di.GetDirectories();

                SortAsFolderCreationTime(ref arrDir);

                if (arrDir.Length > 0)
                {

                    //dirz.

                    try
                    {

                        foreach (DirectoryInfo dir in arrDir)
                        {
                            string[] strFolderNames = dir.FullName.Split('\\');
                            currentFolder = dir.Name;

                            string[] strCurFolderSplit = currentFolder.Split('_');

                            string xmlFilePath = Path.Combine(dir.FullName, "ArchivingSessionInfo.xml");
                            string strDateEnd = "";
                            strSrcDirPath = dir.FullName;
                            strTargetDirPath = Path.Combine(copyToTargetDir, currentFolder);

                            burnDiscParameters = strTargetDirPath;

                            //string strMovetoDelDirPath = Path.Combine(@"E:\TempToDelete", currentFolder);

                            if (!(strCurFolderSplit[0] is "NeedDel"))
                            {
                                XmlDocument xmlDoc = new XmlDocument();
                                xmlDoc.Load(xmlFilePath);

                                strDateEnd = GetDateEnd(xmlDoc);
                                if (strDateEnd != "")
                                {


                                    string[] dateTimes = strDateEnd.Split('T');
                                    DateTime dateOfRec = Convert.ToDateTime(dateTimes[0]);
                                    TimeSpan timeSpan = TimeSpan.FromDays(overDays);

                                    if (DateTime.Today.Subtract(dateOfRec) > timeSpan)
                                    {
                                        //backgroundWorker2.WorkerReportsProgress = true;
                                        backgroundWorker2.RunWorkerAsync();

                                        while (this.backgroundWorker2.IsBusy)
                                        {

                                            // Keep UI messages moving, so the form remains 
                                            // responsive during the asynchronous operation.
                                            Application.DoEvents();
                                        }
                                        break;

                                    }
                                }

                            }

                        }
                        setTaskAtFixedTime();

                    }
                    catch (Exception ex1)
                    {

                        //Console.WriteLine(ex1.Message);
                        WriteException(ex1);


                    }
                }
                else
                {
                    setTaskAtFixedTime();
                }


            }

            catch (Exception ex2)
            {

                //Console.WriteLine(ex2.Message);
                WriteException(ex2);
            }
        }

        /// <summary>
        /// C#按檔案夾夾建立時間排序（順序）
        /// </summary>
        /// <param name="dirs">待排序檔案夾數組</param>
        private void SortAsFolderCreationTime(ref DirectoryInfo[] dirs)
        {
            Array.Sort(dirs, delegate (DirectoryInfo x, DirectoryInfo y) { return x.CreationTime.CompareTo(y.CreationTime); });
        }

        /// <summary>
        /// DirectoryCopy(sour,desti,bool copysubdirectory);
        /// </summary>
        private bool DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }
            try
            {
                DirectoryInfo[] dirs = dir.GetDirectories();

                // If the destination directory doesn't exist, create it.       
                Directory.CreateDirectory(destDirName);

                // Get the files in the directory and copy them to the new location.
                FileInfo[] files = dir.GetFiles();
                foreach (FileInfo file in files)
                {
                    string tempPath = Path.Combine(destDirName, file.Name);
                    file.CopyTo(tempPath, false);
                }

                // If copying subdirectories, copy them and their contents to new location.
                if (copySubDirs)
                {
                    foreach (DirectoryInfo subdir in dirs)
                    {
                        string tempPath = Path.Combine(destDirName, subdir.Name);
                        DirectoryCopy(subdir.FullName, tempPath, copySubDirs);
                    }
                }
                return true;
            }
            catch (Exception ex210410A)
            {
                WriteException(ex210410A);

                return false;
            }


        }

        public string GetDateEnd(XmlDocument xmlDoc)
        {
            //擷取節點
            XmlNodeList xmlNodeList = xmlDoc.SelectNodes("/ArchivingSessionView");
            //擷取將該節點下的 子節點的值
            foreach (XmlNode node in xmlNodeList)
            {
                //方法: node.SelectSingleNode("節點名稱").InnerText;
                string strDateEnd = node.SelectSingleNode("DateEnd").InnerText;

                return strDateEnd;
            }
            //
            return "";
        }


        //static void BurnData(string strDataFolder)
        //{

        //    // Use ProcessStartInfo class
        //    ProcessStartInfo startInfo = new ProcessStartInfo();
        //    startInfo.CreateNoWindow = false;
        //    startInfo.UseShellExecute = false;

        //    //設置啟動動作,確保以管理員身份運行
        //    startInfo.Verb = "runas";

        //    //Give the name as NeroCmd 
        //    startInfo.FileName = @"C:\IMAPI2BurnCD\BurnMedia.exe";
        //    //make the window Hidden
        //    startInfo.WindowStyle = ProcessWindowStyle.Hidden;
        //    //Send the Source and destination as Arguments to the process
        //    //startInfo.Arguments = @"@C:\nerocmd\ParameterFile.txt";
        //    //string strDirPath = strDirectory + "\\" + strCurrDirName + ".iso";

        //    startInfo.Arguments = @"  " + strDataFolder;
        //    try
        //    {
        //        // Start the process with the info we specified.
        //        // Call WaitForExit and then the using statement will close.
        //        using (Process exeProcess2 = Process.Start(startInfo))
        //        {
        //            //if (!exeProcess2.WaitForExit(3600000))
        //            //{
        //            //    exeProcess2.Close();
        //            //    exeProcess2.Dispose();
        //            //}

        //            if (!exeProcess2.CloseMainWindow())
        //            {
        //                exeProcess2.Close();
        //                exeProcess2.Dispose();
        //            }


        //        }

        //    }
        //    catch (Exception exp)
        //    {
        //        throw exp;
        //    }



        //}



        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            double doubleLimitSize = 3.6 * 1024 * 1024 * 1024;

            try
            {
                isSucceed = DirectoryCopy(strSrcDirPath, strTargetDirPath, true);

                long longDirSize = CalculateDirectorySize(strTargetDirPath);
                WriteLog("Directory Size: " + strTargetDirPath + " ==> " + longDirSize.ToString());

                if (Convert.ToDouble(longDirSize) >= doubleLimitSize)
                {
                    throw new Exception("Directory size is over DVD size: " + strTargetDirPath + " ==> " + longDirSize.ToString());
                }

            }
            catch (Exception ex1211A)
            {
                WriteLog(DateTime.Now.ToString() + " : 複製資料匣失敗 ( " + ex1211A.Message + " ) ==> " +
                    Path.GetDirectoryName(strSrcDirPath));
            }


        }

        /// <summary>
        /// Check directory size
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static long CalculateDirectorySize(string DirRoute)
        {
            try
            {
                Type tp = Type.GetTypeFromProgID("Scripting.FileSystemObject");
                object fso = Activator.CreateInstance(tp);
                object fd = tp.InvokeMember("GetFolder", System.Reflection.BindingFlags.InvokeMethod, null, fso, new object[] { DirRoute });
                long ret = Convert.ToInt64(tp.InvokeMember("Size", System.Reflection.BindingFlags.GetProperty, null, fd, null));
                Marshal.ReleaseComObject(fso);
                return ret;
            }
            catch
            {
                return 0;
            }
        }


        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {

            //string strDelDir = (String)e.Argument;
            //try
            //{
            //    DirectoryInfo dirInfo = new DirectoryInfo(strDelDir);

            //    dirInfo.Delete(true);
            //}
            //catch(Exception ex)
            //{
            //    WriteException(ex);
            //}

        }



        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                string chgFolderName = "NeedDel_" + currentFolder;
                string markalreadyCopyPath = Path.Combine(srcNasDirectory, chgFolderName);
                Directory.Move(strSrcDirPath, markalreadyCopyPath);


                debugCount++;
                WriteLog(DateTime.Now.ToString() + "==> Copy Record Data finished. DebugCount = " + debugCount);

                Thread.Sleep(1800000);

                backgroundWorker3.RunWorkerAsync();

                while (this.backgroundWorker3.IsBusy)
                {

                    // Keep UI messages moving, so the form remains 
                    // responsive during the asynchronous operation.
                    Application.DoEvents();
                }
            }
            catch (Exception ex)
            {
                WriteException(ex);
            }






        }

        private void backgroundWorker3_DoWork(object sender, DoWorkEventArgs e)
        {

            //burnDiscParameters = strTargetDirPath;

            //220608A 
            if (strTargetDirPath != "")
            {

                //
                // Determine the current recording devices
                //
                MsftDiscMaster2 discMaster = null;
                try
                {
                    discMaster = new MsftDiscMaster2();

                    if (!discMaster.IsSupportedEnvironment)
                        return;
                    foreach (string uniqueRecorderId in discMaster)
                    {
                        discRecorder = new MsftDiscRecorder2();
                        discRecorder.InitializeDiscRecorder(uniqueRecorderId);
                        
                        //devicesComboBox.Items.Add(discRecorder2);
                    }

                    

                    //if (devicesComboBox.Items.Count > 0)
                    //{
                    //    devicesComboBox.SelectedIndex = 0;
                    //}
                }
                catch (COMException ex)
                {
                    WriteException(ex);
                    return;
                }
                finally
                {
                    if (discMaster != null)
                    {
                        Marshal.ReleaseComObject(discMaster);
                    }
                }


                //
                // Create the volume label based on the current date
                //
                DateTime now = DateTime.Now;
                string discLabel = now.Year + "_" + now.Month + "_" + now.Day;

                //labelStatusText.Text = string.Empty;
                //labelFormatStatusText.Text = string.Empty;

                //
                // Select no verification, by default
                //
                //comboBoxVerification.SelectedIndex = 0;
                try
                {
                    if (burnDiscParameters != "")
                    {
                        directoryItem = new DirectoryItem(burnDiscParameters);
                        //listBoxFiles.Items.Add(directoryItem);
                        

                        //EnableBurnButton();
                        UpdateCapacity();
                        
                        RunBurnProcess();

                        Thread.Sleep(1800000);

                    }
                    else
                    {
                        UpdateCapacity();
                    }
                }
                catch (Exception ex)
                {
                    WriteException(ex);
                }


                //try
                //{

                //    //UpdateCapacity();

                //    RunBurnProcess();

                //    Thread.Sleep(1200000);
                //}
                //catch (Exception ex220606C)
                //{
                //    WriteException(ex220606C);
                //}
            }
        }

        private void UpdateCapacity()
        {
            //
            // Get the text for the Max Size
            //
            //if (_totalDiscSize == 0)
            //{
            //    labelTotalSize.Text = "0MB";
            //    return;
            //}

            //labelTotalSize.Text = _totalDiscSize < 1000000000 ?
            //string.Format("{0}MB", _totalDiscSize / 1000000) :
            //string.Format("{0:F2}GB", (float)_totalDiscSize / 1000000000.0);

            //
            // Calculate the size of the files
            //
            Int64 totalMediaSize = 0;
            //foreach (IMediaItem mediaItem in listBoxFiles.Items)
            //{

            mediaItem = directoryItem;
            totalMediaSize += mediaItem.SizeOnDisc;
            //}

            //if (totalMediaSize == 0)
            //{
            //    progressBarCapacity.Value = 0;
            //    progressBarCapacity.ForeColor = SystemColors.Highlight;
            //}
            //else
            //{
            //    var percent = (int)((totalMediaSize * 100) / _totalDiscSize);
            //    if (percent > 100)
            //    {
            //        progressBarCapacity.Value = 100;
            //        progressBarCapacity.ForeColor = Color.Red;
            //    }
            //    else
            //    {
            //        progressBarCapacity.Value = percent;
            //        progressBarCapacity.ForeColor = SystemColors.Highlight;
            //    }
            //}
        }

        private void backgroundWorker3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 22/6/6 修改

            //再次設定

            try
            {
                if (discRecorder != null)
                {
                    Marshal.ReleaseComObject(discRecorder);
                    discFormatData = null;
                    fileSystemImage = null;
                    mediaItem = null;
                    _burnData = new BurnData();
                }

                if (directoryItem != null)
                {
                    directoryItem = null;
                }
            }
            catch (Exception ex)
            {
                WriteExceptionExt(ex, "再次設定錯誤...");
            
            }

            nasFolderCount = 0;

            isSucceed = false;

            
            
            //UpdateCapacity();

            WriteLog(DateTime.Now.ToString() + "==> Burn DVD finished.");

            //22/05/22 Delete > 30 day folder

            int delNum = 0;

            //sort dirz[]
            DirectoryInfo di2 = new DirectoryInfo(copyToTargetDir);

            DirectoryInfo[] arrDir2 = di2.GetDirectories();

            SortAsFolderCreationTime(ref arrDir2);

            if (arrDir2.Length > 2)
            {
                //dirz.

                try
                {
                    foreach (DirectoryInfo dir in arrDir2)
                    {
                        string[] strFolderNames = dir.FullName.Split('\\');
                        currentFolder = dir.Name;

                        string[] strCurFolderSplit = currentFolder.Split('_');

                        string xmlFilePath = Path.Combine(dir.FullName, "ArchivingSessionInfo.xml");
                        string strDateEnd = "";
                        strSrcDirPath = dir.FullName;
                        strTargetDirPath = Path.Combine(copyToTargetDir, currentFolder);

                        //string strMovetoDelDirPath = Path.Combine(@"E:\TempToDelete", currentFolder);

                        if (strCurFolderSplit[0] is "NeedDel")
                        {
                            XmlDocument xmlDoc = new XmlDocument();
                            xmlDoc.Load(xmlFilePath);

                            strDateEnd = GetDateEnd(xmlDoc);
                            if (strDateEnd != "")
                            {


                                string[] dateTimes = strDateEnd.Split('T');
                                DateTime dateOfRec = Convert.ToDateTime(dateTimes[0]);
                                TimeSpan timeSpan = TimeSpan.FromDays(overDays);

                                if (DateTime.Today.Subtract(dateOfRec) > TimeSpan.FromDays(30))
                                {
                                    string strDelDir = dir.FullName;
                                    try
                                    {
                                        DirectoryInfo dirInfo = new DirectoryInfo(strDelDir);

                                        dirInfo.Delete(true);

                                        WriteLog("刪除目錄成功: " + dir.FullName);
                                        delNum++;
                                    }
                                    catch (Exception ex)
                                    {
                                        WriteException(ex);
                                    }

                                }


                            }

                        }
                        else
                        {
                            DirectoryInfo burnFolder = new DirectoryInfo(burnDiscParameters);
                            if (dir.Name != burnFolder.Name)
                            {
                                string chgFolderName = "NeedDel_" + dir.Name;
                                string markalreadyCopyPath = Path.Combine(copyToTargetDir, chgFolderName);
                                Directory.Move(strTargetDirPath, markalreadyCopyPath);
                            }
                            
                        }
                        if (delNum > 1 || arrDir2.Length <= 1)
                        {
                            break;
                        }
                    }
                    
                }
                catch (Exception ex1)
                {

                    //Console.WriteLine(ex1.Message);
                    WriteException(ex1);


                }
            }

            strTargetDirPath = "";
            burnDiscParameters = "";
            _isBurning = false;
            _closeMedia = false;
            _ejectMedia = false;


        }

     


        private void RunBurnProcess()
        {
            try
            {
                

                //var directoryItem = new DirectoryItem(strTargetDirPath);

                nasFolderCount++;

                //UpdateCapacity();

                WriteLog("Start to burn disc.");





                _isBurning = true;
                _closeMedia = true;
                _ejectMedia = true;



                //discRecorder = discRecorder2;
                _burnData.uniqueRecorderId = discRecorder.ActiveDiscRecorder;

                backgroundBurnWorker.RunWorkerAsync(_burnData);
            }
            catch(Exception ex)
            {
                WriteException(ex);
            }
            

        }

        private void backgroundBurnWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            MsftDiscRecorder2 discRecorder3 = null;
            discFormatData = null;

            try
            {
                //
                // Create and initialize the IDiscRecorder2 object
                //
                discRecorder3 = new MsftDiscRecorder2();
                var burnData = (BurnData)e.Argument;
                discRecorder3.InitializeDiscRecorder(burnData.uniqueRecorderId);
                

                //
                // Create and initialize the IDiscFormat2Data
                //
                discFormatData = new MsftDiscFormat2Data
                {
                    Recorder = discRecorder,
                    ClientName = ClientName,
                    ForceMediaToBeClosed = _closeMedia
                };

                //
                // Set the verification level
                //
                var burnVerification = (IBurnVerification)discFormatData;
                burnVerification.BurnVerificationLevel = _verificationLevel;

                //
                // Check if media is blank, (for RW media)
                //
                object[] multisessionInterfaces = null;
                if (!discFormatData.MediaHeuristicallyBlank)
                {
                    multisessionInterfaces = discFormatData.MultisessionInterfaces;
                }

                //
                // Create the file system
                //
                IStream fileSystem;
                if (!CreateMediaFileSystem(discRecorder3, multisessionInterfaces, out fileSystem))
                {
                    e.Result = -1;
                    return;
                }

                //
                // add the Update event handler
                //
                discFormatData.Update += discFormatData_Update;

                //
                // Write the data here
                //
                try
                {                    
                    discFormatData.SetWriteSpeed(4, false);
                    discFormatData.Write(fileSystem);
                    e.Result = 0;
                }
                catch (COMException ex)
                {
                    e.Result = ex.ErrorCode;
                    WriteExceptionExt(ex , "IDiscFormat2Data.Write failed");
                    //MessageBox.Show(ex.Message, "IDiscFormat2Data.Write failed",
                    //    MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                finally
                {
                    if (fileSystem != null)
                    {
                        Marshal.FinalReleaseComObject(fileSystem);
                    }
                }

                //
                // remove the Update event handler
                //
                discFormatData.Update -= discFormatData_Update;

                if (_ejectMedia)
                {
                    discRecorder3.EjectMedia();
                    _ejectMedia = false;
                }
            }
            catch (COMException exception)
            {
                //
                // If anything happens during the format, show the message
                //
                WriteException(exception);

            }
            finally
            {
                if (discRecorder != null)
                {
                    Marshal.ReleaseComObject(discRecorder3);
                }

                if (discFormatData != null)
                {
                    Marshal.ReleaseComObject(discFormatData);
                }
            }
        }

        /// <summary>
        /// Completed the "Burn" thread
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundBurnWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            WriteLog((int)e.Result == 0 ? (DateTime.Now.ToString() + " : Finished Burning Disc!") : (DateTime.Now.ToString() + " : Error Burning Disc!"));


            _isBurning = false;



            //string[] strFolderNames;
            string currentFolderName = "";

            if (burnDiscParameters != "")
            {
                string[] strTmpFolderNames = burnDiscParameters.Split('\\');
                string currentTmpFolderName = strTmpFolderNames[strTmpFolderNames.Length - 1];

                currentFolderName = currentTmpFolderName;
            }



            if ((int)e.Result != 0)
            {
                if (burnDiscParameters != "")
                {
                    WriteLog("false," + "Error_Burning Disc : " + DateTime.Now.ToString() + "Error_Burning Folder : \r\n " + burnDiscParameters);

                    //string[] strFolderNames = burnDiscParameters.Split('\\');
                    //string currentFolderName = strFolderNames[strFolderNames.Length - 1];


                }
                else
                {
                    WriteLog("false," + "Error_Burning Disc : " + DateTime.Now.ToString());
                }

                

            }
            else
            {
                WriteLog("true," + "Finished Burning Disc : " + DateTime.Now.ToString());
               

            }

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="progress"></param>
        void discFormatData_Update([In, MarshalAs(UnmanagedType.IDispatch)] object sender, [In, MarshalAs(UnmanagedType.IDispatch)] object progress)
        {
            //
            // Check if we've cancelled
            //
            if (backgroundBurnWorker.CancellationPending)
            {
                var format2Data = (IDiscFormat2Data)sender;
                format2Data.CancelWrite();
                return;
            }

            var eventArgs = (IDiscFormat2DataEventArgs)progress;

            _burnData.task = BURN_MEDIA_TASK.BURN_MEDIA_TASK_WRITING;

            // IDiscFormat2DataEventArgs Interface
            _burnData.elapsedTime = eventArgs.ElapsedTime;
            _burnData.remainingTime = eventArgs.RemainingTime;
            _burnData.totalTime = eventArgs.TotalTime;


            // IWriteEngine2EventArgs Interface
            _burnData.currentAction = eventArgs.CurrentAction;
            _burnData.startLba = eventArgs.StartLba;
            _burnData.sectorCount = eventArgs.SectorCount;
            _burnData.lastReadLba = eventArgs.LastReadLba;
            _burnData.lastWrittenLba = eventArgs.LastWrittenLba;
            _burnData.totalSystemBuffer = eventArgs.TotalSystemBuffer;
            _burnData.usedSystemBuffer = eventArgs.UsedSystemBuffer;
            _burnData.freeSystemBuffer = eventArgs.FreeSystemBuffer;

            //
            // Report back to the UI
            //
            //backgroundBurnWorker.ReportProgress(0, _burnData);
        }

        private void fileSystemImage_Update([In, MarshalAs(UnmanagedType.IDispatch)] object sender,
            [In, MarshalAs(UnmanagedType.BStr)] string currentFile, [In] int copiedSectors, [In] int totalSectors)
        {
            var percentProgress = 0;
            if (copiedSectors > 0 && totalSectors > 0)
            {
                percentProgress = (copiedSectors * 100) / totalSectors;
            }

            if (!string.IsNullOrEmpty(currentFile))
            {
                var fileInfo = new FileInfo(currentFile);
                _burnData.statusMessage = "Adding \"" + fileInfo.Name + "\" to image...";

                //
                // report back to the ui
                //
                _burnData.task = BURN_MEDIA_TASK.BURN_MEDIA_TASK_FILE_SYSTEM;
                //backgroundBurnWorker.ReportProgress(percentProgress, _burnData);
            }

        }

        private bool CreateMediaFileSystem(IDiscRecorder2 discRecorder, object[] multisessionInterfaces, out IStream dataStream)
        {
            fileSystemImage = null;
            try
            {
                fileSystemImage = new MsftFileSystemImage();
                fileSystemImage.ChooseImageDefaults(discRecorder);
                fileSystemImage.FileSystemsToCreate =
                    FsiFileSystems.FsiFileSystemJoliet | FsiFileSystems.FsiFileSystemISO9660;

                DateTime now = DateTime.Now;

                fileSystemImage.VolumeName = now.Year + "_" + now.Month + "_" + now.Day; 

                fileSystemImage.Update += fileSystemImage_Update;

                //
                // If multisessions, then import previous sessions
                //
                //if (multisessionInterfaces != null)
                //{
                //    fileSystemImage.MultisessionInterfaces = multisessionInterfaces;
                //    fileSystemImage.ImportFileSystem();
                //}

                //
                // Get the image root
                //
                IFsiDirectoryItem rootItem = fileSystemImage.Root;

                //
                // Add Files and Directories to File System Image
                //
                //foreach (IMediaItem mediaItem in listBoxFiles.Items)
                //{
                //    //
                //    // Check if we've cancelled
                //    //
                //    if (backgroundBurnWorker.CancellationPending)
                //    {
                //        break;
                //    }

                //
                // Add to File System
                //
                //mediaItem = (IMediaItem)directoryItem;
                mediaItem.AddToFileSystem(rootItem);
                //}

                fileSystemImage.Update -= fileSystemImage_Update;

                //
                // did we cancel?
                //
                //if (backgroundBurnWorker.CancellationPending)
                //{
                //    dataStream = null;
                //    return false;
                //}

                dataStream = fileSystemImage.CreateResultImage().ImageStream;
            }
            catch (COMException exception)
            {
                //MessageBox.Show(this, exception.Message, "Create File System Error",
                //    MessageBoxButtons.OK, MessageBoxIcon.Error);
                WriteException(exception);
                dataStream = null;
                return false;
            }
            finally
            {
                if (fileSystemImage != null)
                {
                    Marshal.ReleaseComObject(fileSystemImage);
                }
            }

            return true;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //WriteLog("刪除目錄成功。");
        }
    }
}
