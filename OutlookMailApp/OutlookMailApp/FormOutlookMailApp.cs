// Released as open source by NCC Group Plc - https://www.nccgroup.trust/
// Developed by Soroush Dalili (@irsdl)
// Released under AGPL see LICENSE for more information
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Email.Mail;
using Aspose.Email.Outlook;
using System.Threading;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading.Tasks;
using System.Security.Cryptography;
using System.Diagnostics;

namespace OutlookMailApp
{
    public partial class FormOutlookMailApp : Form
    {
        TimestampLogging timestampLogger;
        SimpleLogging payloadsListLogger;
        SimpleLogging payloadsProcessedTestcaseFilesLogger;
        static long caseID = 0;
        private ManualResetEvent mre = new ManualResetEvent(true);
        private long paused = 0;
        private int processTimeoutInSecond = 10;
        private DateTime lastKillRequest = DateTime.Now;
        private string additionalStatusDetails = "";

        public FormOutlookMailApp()
        {
            InitializeComponent();
        }

        private void genrateSampleTemplate()
        {
            string strFrom = textBoxFrom.Text;
            string strTo = textBoxTo.Text;
            string strCC = textBoxCC.Text;
            string strSubject = textBoxSubject.Text;
            string strTemplate = textBoxTemplate.Text;
            string strWorkPath = textBoxWorkPath.Text;
            string strPreProcessedFolder = textBoxPreProcessedFolder.Text;
            string strProcessedFolder = textBoxProcessedFolder.Text;
            string strTestcasePrefix = textBoxTestCasePrefix.Text;
            string strLogfilePrefix = textBoxLogfilePrefix.Text;
            int intKeepOpen = StrToIntDef(textBoxKeepOpen.Text, 3);
            processTimeoutInSecond = StrToIntDef(textBoxProcessTimeout.Text, 10);

            textBoxKeepOpen.Text = intKeepOpen.ToString();
            Boolean isURLEncoded = checkBoxURLDecode.Checked;
            Boolean isAutoRun = checkBoxAutoRun.Checked;

            timestampLogger = new TimestampLogging(strWorkPath + "\\" + strLogfilePrefix);
            payloadsProcessedTestcaseFilesLogger = new SimpleLogging(strWorkPath + "\\" + textBoxProcessedTestcaseFilesFilename.Text);
            //logger.log("==Begin==");

            // to kill it on timeout: http://stackoverflow.com/questions/1410602/how-do-set-a-timeout-for-a-method
            Action<object, string, string, string, int> longMethod_showCloseMoveEmail = showCloseMoveEmail; 

            try
            {


                string strPreProcessedFolderFullPath = strWorkPath + "\\" + strPreProcessedFolder + "\\";
                string strProcessedFolderFullPath = strWorkPath + "\\" + strProcessedFolder + "\\";
                System.IO.Directory.CreateDirectory(strPreProcessedFolderFullPath);
                System.IO.Directory.CreateDirectory(strProcessedFolderFullPath);

                if (isURLEncoded)
                {
                    timestampLogger.log("URL Decoding...");
                    strTemplate = Uri.UnescapeDataString(strTemplate);
                }

                MailMessage msg = new MailMessage();

                // Set recipients information
                if (!String.IsNullOrEmpty(strFrom))
                    msg.From = strFrom;
                if (!String.IsNullOrEmpty(strTo))
                    msg.To = strTo;
                if (!String.IsNullOrEmpty(strCC))
                    msg.CC = strCC;

                // Set the subject
                msg.Subject = strSubject;

                // Set HTML body
                msg.HtmlBody = strTemplate;

                // Add an attachment
                // msg.Attachments.Add(new Aspose.Email.Mail.Attachment("test.txt"));

                // Local filenames
                string strTestCaseFileName = MakeValidFileName(strTestcasePrefix + strSubject + ".msg");
                string strInProgressFilePath = strPreProcessedFolderFullPath + strTestCaseFileName;

                //msg.Save(strFullFilePath,SaveOptions.DefaultMsg);

                MapiMessage outlookMsg = MapiMessage.FromMailMessage(msg);
                outlookMsg.SetMessageFlags(MapiMessageFlags.MSGFLAG_SUBMIT);


                while (File.Exists(strInProgressFilePath))
                {
                    strInProgressFilePath = strPreProcessedFolderFullPath + Guid.NewGuid() + "_" + strTestCaseFileName;
                }
                timestampLogger.log("Saving the " + strTestCaseFileName + " file in: " + strInProgressFilePath);
                outlookMsg.Save(strInProgressFilePath);

                if (isAutoRun)
                {
                    new Thread(delegate()
                    {
                        //showCloseMoveEmail(strInProgressFilePath, strProcessedFolderFullPath, strTestCaseFileName, intKeepOpen);

                        // to kill it on timeout: http://stackoverflow.com/questions/1410602/how-do-set-a-timeout-for-a-method
                        object monitorSync = new object();
                        bool timedOut;
                        lock (monitorSync)
                        {
                            longMethod_showCloseMoveEmail.BeginInvoke(monitorSync, strInProgressFilePath, strProcessedFolderFullPath, strTestCaseFileName, intKeepOpen, null, null);
                            timedOut = !Monitor.Wait(monitorSync, TimeSpan.FromSeconds(processTimeoutInSecond));
                        }
                        if (timedOut)
                        {
                            object _lock = new object();
                            lock (_lock)
                            {
                                killProcess("OUTLOOK", processTimeoutInSecond);
                                timestampLogger.log("Process killed due to " + processTimeoutInSecond.ToString() + " seconds timeout");
                            }
                        }

                    }).Start();
                }
            }
            catch (Exception e)
            {
                timestampLogger.log("Error in genrateSampleTemplate(): " + e.StackTrace);
            }
            //logger.log("==End==");
        }

        private void genrateTestCases()
        {
            string strFrom = textBoxFrom.Text;
            string strTo = textBoxTo.Text;
            string strCC = textBoxCC.Text;
            string strSubject = textBoxSubject.Text;
            string strTemplate = textBoxTemplate.Text;
            string strWorkPath = textBoxWorkPath.Text;
            string strPreProcessedFolder = textBoxPreProcessedFolder.Text;
            string strProcessedFolder = textBoxProcessedFolder.Text;
            string strTestcasePrefix = textBoxTestCasePrefix.Text;
            string strLogfilePrefix = textBoxLogfilePrefix.Text;
            string strPayloadPattern = textBoxPayloadPattern.Text;
            int intKeepOpen = StrToIntDef(textBoxKeepOpen.Text, 3);
            processTimeoutInSecond = StrToIntDef(textBoxProcessTimeout.Text, 10);
            textBoxKeepOpen.Text = intKeepOpen.ToString();
            Boolean isURLEncoded = checkBoxURLDecode.Checked;
            Boolean isAutoRun = checkBoxAutoRun.Checked;

            

            timestampLogger = new TimestampLogging(strWorkPath + "\\" + strLogfilePrefix);
            payloadsListLogger = new SimpleLogging(strWorkPath + "\\" + textBoxPayloadsListFilename.Text);
            payloadsProcessedTestcaseFilesLogger = new SimpleLogging(strWorkPath + "\\" + textBoxProcessedTestcaseFilesFilename.Text);
            //logger.log("==Begin==");

            try
            {
                caseID = 0;
                string[] arrPrefix = UniqueTextArrayFromFile(textBoxPrefix.Text);
                string[] arrSuffix = UniqueTextArrayFromFile(textBoxSuffix.Text);
                string[] arrFormula = UniqueTextArrayFromFile(textBoxFormula.Text);
                string[] arrSchemes = UniqueTextArrayFromFile(textBoxSchemes.Text);
                string[] arrTargets = UniqueTextArrayFromFile(textBoxTargets.Text);
                string[] arrSpecialFormula = UniqueTextArrayFromFile(textBoxSpecialFormula.Text);

                string strPreProcessedFolderFullPath = strWorkPath + "\\" + strPreProcessedFolder + "\\";
                string strProcessedFolderFullPath = strWorkPath + "\\" + strProcessedFolder + "\\";
                System.IO.Directory.CreateDirectory(strPreProcessedFolderFullPath);
                System.IO.Directory.CreateDirectory(strProcessedFolderFullPath);

                if (isURLEncoded)
                {
                    timestampLogger.log("URL Decoding...");
                    strTemplate = Uri.UnescapeDataString(strTemplate);
                }

                string strTemplateSha1Sig = SHA1FromString(strTemplate);

                int intPrefixParalDegree = StrToIntDef(textBoxThreadPrefixHigh.Text,1);
                int intTargetsParalDegree = StrToIntDef(textBoxThreadTargetsHigh.Text, 2);
                int intSchemesParalDegree = StrToIntDef(textBoxThreadSchemesHigh.Text, 5);
                int intFormulaParalDegree = StrToIntDef(textBoxThreadFormulaHigh.Text, 5);
                int intSuffixParalDegree = StrToIntDef(textBoxThreadSuffixHigh.Text, 1);
                if (isAutoRun)
                {
                    intPrefixParalDegree = StrToIntDef(textBoxThreadPrefixLow.Text, 1);
                    intTargetsParalDegree = StrToIntDef(textBoxThreadTargetsLow.Text, 1);
                    intSchemesParalDegree = StrToIntDef(textBoxThreadSchemesLow.Text, 3);
                    intFormulaParalDegree = StrToIntDef(textBoxThreadFormulaLow.Text, 3);
                    intSuffixParalDegree = StrToIntDef(textBoxThreadSuffixLow.Text, 1);
                }
                

                long totalEvents = arrPrefix.Length * arrTargets.Length * arrSchemes.Length * arrFormula.Length * arrSuffix.Length;
                progressBarStatus.Value = 0;
                setLabelStatus(0, totalEvents);
                ResetCaseID();

                // to kill it on timeout: http://stackoverflow.com/questions/1410602/how-do-set-a-timeout-for-a-method
                Action<object, string, string, string, int> longMethod_showCloseMoveEmail = showCloseMoveEmail;

                Parallel.ForEach(arrPrefix, new ParallelOptions { MaxDegreeOfParallelism = intPrefixParalDegree }, (strPrefix, loopStatePrefix) =>
                    {
                        Parallel.ForEach(arrTargets, new ParallelOptions { MaxDegreeOfParallelism = intTargetsParalDegree }, (strTarget, loopStateTargets) =>
                            {
                                Parallel.ForEach(arrSchemes, new ParallelOptions { MaxDegreeOfParallelism = intSchemesParalDegree }, (strScheme, loopStateScheme) =>
                                    {
                                        Parallel.ForEach(arrFormula, new ParallelOptions { MaxDegreeOfParallelism = intFormulaParalDegree }, (strFormula, loopStateFormula) =>
                                            {
                                                Parallel.ForEach(arrSuffix, new ParallelOptions { MaxDegreeOfParallelism = intSuffixParalDegree }, (strSuffix, loopStateSuffix) =>
                                                    {
                                                        if (Interlocked.Read(ref paused) == 1)
                                                            mre.WaitOne();

                                                        long currentCaseID = GetNextValue();

                                                        string strPayload = strFormula.Replace("<$scheme$>", strScheme);
                                                        strPayload = strPayload.Replace("<$target$>", strTarget);
                                                        strPayload = strPrefix + strPayload + strSuffix + currentCaseID.ToString();
                                                        payloadsListLogger.log(strTemplateSha1Sig + "," + strPayload);

                                                        string strTempTemplate = System.Text.RegularExpressions.Regex.Replace(strTemplate, strPayloadPattern, strPayload);
                                                        string strTempSubject = System.Text.RegularExpressions.Regex.Replace(strSubject, strPayloadPattern, strPayload);
                                                        string strTempFrom = System.Text.RegularExpressions.Regex.Replace(strFrom, strPayloadPattern, strPayload);
                                                        string strTempTo = System.Text.RegularExpressions.Regex.Replace(strTo, strPayloadPattern, strPayload);
                                                        string strTempCC = System.Text.RegularExpressions.Regex.Replace(strCC, strPayloadPattern, strPayload);


                                                        MailMessage msg = new MailMessage();

                                                        // Set recipients information
                                                        if (!String.IsNullOrEmpty(strTempFrom))
                                                            msg.From = strTempFrom;
                                                        if (!String.IsNullOrEmpty(strTempTo))
                                                            msg.To = strTempTo;
                                                        if (!String.IsNullOrEmpty(strTempCC))
                                                            msg.CC = strTempCC;

                                                        // Set the subject
                                                        msg.Subject = strTempSubject;

                                                        // Set HTML body
                                                        msg.HtmlBody = strTempTemplate;

                                                        // Add an attachment
                                                        // msg.Attachments.Add(new Aspose.Email.Mail.Attachment("test.txt"));

                                                        // Local filenames
                                                        string strTestCaseFileName = MakeValidFileName(strTestcasePrefix + currentCaseID.ToString() + "-" + strTempSubject + ".msg");
                                                        string strInProgressFilePath = strPreProcessedFolderFullPath + strTestCaseFileName;

                                                        //msg.Save(strFullFilePath,SaveOptions.DefaultMsg);

                                                        MapiMessage outlookMsg = MapiMessage.FromMailMessage(msg);
                                                        outlookMsg.SetMessageFlags(MapiMessageFlags.MSGFLAG_SUBMIT);


                                                        while (File.Exists(strInProgressFilePath))
                                                        {
                                                            strInProgressFilePath = strPreProcessedFolderFullPath + Guid.NewGuid() + "_" + strTestCaseFileName;
                                                        }
                                                        timestampLogger.log("Saving the " + strTestCaseFileName + " file in: " + strInProgressFilePath);
                                                        outlookMsg.Save(strInProgressFilePath);

                                                        if (isAutoRun)
                                                        {
                                                            mre.WaitOne();
                                                            //showCloseMoveEmail(strInProgressFilePath, strProcessedFolderFullPath, strTestCaseFileName, intKeepOpen);

                                                            // to kill it on timeout: http://stackoverflow.com/questions/1410602/how-do-set-a-timeout-for-a-method
                                                            object monitorSync = new object();
                                                            bool timedOut;
                                                            lock (monitorSync)
                                                            {
                                                                longMethod_showCloseMoveEmail.BeginInvoke(monitorSync, strInProgressFilePath, strProcessedFolderFullPath, strTestCaseFileName, intKeepOpen, null, null);
                                                                timedOut = !Monitor.Wait(monitorSync, TimeSpan.FromSeconds(processTimeoutInSecond));
                                                            }
                                                            if (timedOut)
                                                            {
                                                                object _lock = new object();
                                                                lock (_lock)
                                                                {
                                                                    killProcess("OUTLOOK", processTimeoutInSecond);
                                                                    timestampLogger.log("Process killed due to " + processTimeoutInSecond.ToString() + " seconds timeout");
                                                                }
                                                            }
                                                            
                                                            
                                                            setLabelStatus(currentCaseID, totalEvents);
                                                            if (currentCaseID % 10 == 0 || currentCaseID == totalEvents)
                                                            {
                                                                setLabelStatus(currentCaseID, totalEvents);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (currentCaseID % 100 == 0 || currentCaseID == totalEvents)
                                                            {
                                                                setLabelStatus(currentCaseID, totalEvents);
                                                            }
                                                        
                                                        }
                                                    }
                                                );
                                            }
                                        );
                                    }
                                );
                            }
                        );
                    }

                    );

            }
            catch (Exception e)
            {
                timestampLogger.log("Error in genrateTestCases(): " + e.StackTrace);
                MessageBox.Show("An error occured!", "Error!");
            }

            //logger.log("==End==");
        }


        private void genrateSpecialTestCases()
        {
            string strFrom = textBoxFrom.Text;
            string strTo = textBoxTo.Text;
            string strCC = textBoxCC.Text;
            string strSubject = textBoxSubject.Text;
            string strTemplate = textBoxTemplate.Text;
            string strWorkPath = textBoxWorkPath.Text;
            string strPreProcessedFolder = textBoxPreProcessedFolder.Text;
            string strProcessedFolder = textBoxProcessedFolder.Text;
            string strTestcasePrefix = textBoxTestCasePrefix.Text;
            string strLogfilePrefix = textBoxLogfilePrefix.Text;
            string strPayloadPattern = textBoxPayloadPattern.Text;
            int intKeepOpen = StrToIntDef(textBoxKeepOpen.Text, 3);
            processTimeoutInSecond = StrToIntDef(textBoxProcessTimeout.Text, 10);
            textBoxKeepOpen.Text = intKeepOpen.ToString();
            Boolean isURLEncoded = checkBoxURLDecode.Checked;
            Boolean isAutoRun = checkBoxAutoRun.Checked;



            timestampLogger = new TimestampLogging(strWorkPath + "\\" + strLogfilePrefix);
            payloadsListLogger = new SimpleLogging(strWorkPath + "\\" + textBoxPayloadsListFilename.Text);
            payloadsProcessedTestcaseFilesLogger = new SimpleLogging(strWorkPath + "\\" + textBoxProcessedTestcaseFilesFilename.Text);
            //logger.log("==Begin==");

            try
            {
                caseID = 0;
                string[] arrPrefix = UniqueTextArrayFromFile(textBoxPrefix.Text);
                string[] arrSuffix = UniqueTextArrayFromFile(textBoxSuffix.Text);
                string[] arrFormula = UniqueTextArrayFromFile(textBoxFormula.Text);
                string[] arrSchemes = UniqueTextArrayFromFile(textBoxSchemes.Text);
                string[] arrTargets = UniqueTextArrayFromFile(textBoxTargets.Text);
                string[] arrSpecialFormula = UniqueTextArrayFromFile(textBoxSpecialFormula.Text);

                string strPreProcessedFolderFullPath = strWorkPath + "\\" + strPreProcessedFolder + "\\";
                string strProcessedFolderFullPath = strWorkPath + "\\" + strProcessedFolder + "\\";
                System.IO.Directory.CreateDirectory(strPreProcessedFolderFullPath);
                System.IO.Directory.CreateDirectory(strProcessedFolderFullPath);

                if (isURLEncoded)
                {
                    timestampLogger.log("URL Decoding...");
                    strTemplate = Uri.UnescapeDataString(strTemplate);
                }

                string strTemplateSha1Sig = SHA1FromString(strTemplate);

                int intPrefixParalDegree = StrToIntDef(textBoxThreadPrefixHigh.Text, 1);
                int intTargetsParalDegree = StrToIntDef(textBoxThreadTargetsHigh.Text, 2);
                int intSpecialFormulaParalDegree = StrToIntDef(textBoxThreadSpecialFormulaHigh.Text, 5);
                int intSuffixParalDegree = StrToIntDef(textBoxThreadSuffixHigh.Text, 1);
                if (isAutoRun)
                {
                    intPrefixParalDegree = StrToIntDef(textBoxThreadPrefixLow.Text, 1);
                    intTargetsParalDegree = StrToIntDef(textBoxThreadTargetsLow.Text, 1);
                    intSpecialFormulaParalDegree = StrToIntDef(textBoxThreadSpecialFormulaLow.Text, 3);
                    intSuffixParalDegree = StrToIntDef(textBoxThreadSuffixLow.Text, 1);
                }

                long totalEvents = arrPrefix.Length * arrTargets.Length * arrSpecialFormula.Length * arrSuffix.Length;
                progressBarStatus.Value = 0;
                setLabelStatus(0, totalEvents);
                ResetCaseID();

                // to kill it on timeout: http://stackoverflow.com/questions/1410602/how-do-set-a-timeout-for-a-method
                Action<object, string, string, string, int> longMethod_showCloseMoveEmail = showCloseMoveEmail;

                Parallel.ForEach(arrPrefix, new ParallelOptions { MaxDegreeOfParallelism = intPrefixParalDegree }, (strPrefix, loopStatePrefix) =>
                {
                    Parallel.ForEach(arrTargets, new ParallelOptions { MaxDegreeOfParallelism = intTargetsParalDegree }, (strTarget, loopStateTargets) =>
                    {
                        Parallel.ForEach(arrSpecialFormula, new ParallelOptions { MaxDegreeOfParallelism = intSpecialFormulaParalDegree }, (strSpecialFormula, loopSpecialFormula) =>
                        {
                            Parallel.ForEach(arrSuffix, new ParallelOptions { MaxDegreeOfParallelism = intSuffixParalDegree }, (strSuffix, loopStateSuffix) =>
                                {
                                    if (Interlocked.Read(ref paused) == 1)
                                        mre.WaitOne();

                                    long currentCaseID = GetNextValue();
                                    string strPayload = strSpecialFormula.Replace("<$target$>", strTarget);
                                    strPayload = strPrefix + strPayload + strSuffix + currentCaseID.ToString();
                                    payloadsListLogger.log(strTemplateSha1Sig + "," + strPayload);

                                    string strTempTemplate = System.Text.RegularExpressions.Regex.Replace(strTemplate, strPayloadPattern, strPayload);
                                    string strTempSubject = System.Text.RegularExpressions.Regex.Replace(strSubject, strPayloadPattern, strPayload);
                                    string strTempFrom = System.Text.RegularExpressions.Regex.Replace(strFrom, strPayloadPattern, strPayload);
                                    string strTempTo = System.Text.RegularExpressions.Regex.Replace(strTo, strPayloadPattern, strPayload);
                                    string strTempCC = System.Text.RegularExpressions.Regex.Replace(strCC, strPayloadPattern, strPayload);


                                    MailMessage msg = new MailMessage();

                                    // Set recipients information
                                    if (!String.IsNullOrEmpty(strTempFrom))
                                        msg.From = strTempFrom;
                                    if (!String.IsNullOrEmpty(strTempTo))
                                        msg.To = strTempTo;
                                    if (!String.IsNullOrEmpty(strTempCC))
                                        msg.CC = strTempCC;

                                    // Set the subject
                                    msg.Subject = strTempSubject;

                                    // Set HTML body
                                    msg.HtmlBody = strTempTemplate;

                                    // Add an attachment
                                    // msg.Attachments.Add(new Aspose.Email.Mail.Attachment("test.txt"));

                                    // Local filenames
                                    string strTestCaseFileName = MakeValidFileName(strTestcasePrefix + currentCaseID.ToString() + "-" + strTempSubject + ".msg");
                                    string strInProgressFilePath = strPreProcessedFolderFullPath + strTestCaseFileName;

                                    //msg.Save(strFullFilePath,SaveOptions.DefaultMsg);

                                    MapiMessage outlookMsg = MapiMessage.FromMailMessage(msg);
                                    outlookMsg.SetMessageFlags(MapiMessageFlags.MSGFLAG_SUBMIT);


                                    while (File.Exists(strInProgressFilePath))
                                    {
                                        strInProgressFilePath = strPreProcessedFolderFullPath + Guid.NewGuid() + "_" + strTestCaseFileName;
                                    }
                                    timestampLogger.log("Saving the " + strTestCaseFileName + " file in: " + strInProgressFilePath);
                                    outlookMsg.Save(strInProgressFilePath);

                                    if (isAutoRun)
                                    {
                                        mre.WaitOne();
                                        //showCloseMoveEmail(strInProgressFilePath, strProcessedFolderFullPath, strTestCaseFileName, intKeepOpen);

                                        // to kill it on timeout: http://stackoverflow.com/questions/1410602/how-do-set-a-timeout-for-a-method
                                        object monitorSync = new object();
                                        bool timedOut;
                                        lock (monitorSync)
                                        {
                                            longMethod_showCloseMoveEmail.BeginInvoke(monitorSync, strInProgressFilePath, strProcessedFolderFullPath, strTestCaseFileName, intKeepOpen, null, null);
                                            timedOut = !Monitor.Wait(monitorSync, TimeSpan.FromSeconds(processTimeoutInSecond));
                                        }
                                        if (timedOut)
                                        {
                                            object _lock = new object();
                                            lock (_lock)
                                            {
                                                killProcess("OUTLOOK", processTimeoutInSecond);
                                                timestampLogger.log("Process killed due to " + processTimeoutInSecond.ToString() + " seconds timeout");
                                            }
                                        }


                                        if (currentCaseID % 10 == 0 || currentCaseID == totalEvents)
                                        {
                                            setLabelStatus(currentCaseID, totalEvents);
                                        }
                                    }
                                    else
                                    {
                                        if (currentCaseID % 100 == 0 || currentCaseID == totalEvents)
                                        {
                                            setLabelStatus(currentCaseID, totalEvents);
                                        }
                                    }
                                }
                                );
                        }
                            );
                    }
                        );

                }

               );

            }
            catch (Exception e)
            {
                timestampLogger.log("Error in genrateSpecialTestCases(): " + e.StackTrace);
                MessageBox.Show("An error occured!", "Error!");
            }
            //logger.log("==End==");
        }


        private void RunTestCases()
        {
            string strWorkPath = textBoxWorkPath.Text;
            string strPreProcessedFolder = textBoxPreProcessedFolder.Text;
            string strProcessedFolder = textBoxProcessedFolder.Text;
            string strTestcasePrefix = textBoxTestCasePrefix.Text;
            string strLogfilePrefix = textBoxLogfilePrefix.Text;
            buttonControl.Enabled = true;
            int intKeepOpen = StrToIntDef(textBoxKeepOpen.Text, 3);
            processTimeoutInSecond = StrToIntDef(textBoxProcessTimeout.Text, 10);
            textBoxKeepOpen.Text = intKeepOpen.ToString();

            timestampLogger = new TimestampLogging(strWorkPath + "\\" + strLogfilePrefix);
            string strPreProcessedFolderFullPath = strWorkPath + "\\" + strPreProcessedFolder + "\\";
            string strProcessedFolderFullPath = strWorkPath + "\\" + strProcessedFolder + "\\";

            payloadsProcessedTestcaseFilesLogger = new SimpleLogging(strWorkPath + "\\" + textBoxProcessedTestcaseFilesFilename.Text);

            // to kill it on timeout: http://stackoverflow.com/questions/1410602/how-do-set-a-timeout-for-a-method
            Action<object, string, string, string, int> longMethod_showCloseMoveEmail = showCloseMoveEmail;

            int intTestcaseRunnerThread = StrToIntDef(textBoxThreadTestcaseRunner.Text, 10);

            if (Directory.Exists(strPreProcessedFolderFullPath))
            {
                string[] files = Directory.GetFiles(strPreProcessedFolderFullPath, "*.msg", SearchOption.AllDirectories);

                long totalEvents = files.Length;
                ResetCaseID();
                setLabelStatus(0, totalEvents);
                bool timedOut = false;
                Parallel.ForEach(files, new ParallelOptions { MaxDegreeOfParallelism = intTestcaseRunnerThread }, (strTestCaseFilePath, loopStateFiles) =>
                {
                    string strTestCaseFileName = Path.GetFileName(strTestCaseFilePath);
                    mre.WaitOne();
                    //showCloseMoveEmail(strTestCaseFilePath, strProcessedFolderFullPath, strTestCaseFileName, intKeepOpen);

                    // to kill it on timeout: http://stackoverflow.com/questions/1410602/how-do-set-a-timeout-for-a-method
                    object monitorSync = new object();
                    timedOut = false;
                    lock (monitorSync) {
                        longMethod_showCloseMoveEmail.BeginInvoke(monitorSync,strTestCaseFilePath, strProcessedFolderFullPath, strTestCaseFileName, intKeepOpen, null, null);
                        timedOut = !Monitor.Wait(monitorSync, TimeSpan.FromSeconds(processTimeoutInSecond));
                    }
                    if (timedOut) {
                        object _lock = new object();
                        lock (_lock)
                        {
                            killProcess("OUTLOOK", processTimeoutInSecond);
                            timestampLogger.log("Process killed due to " + processTimeoutInSecond.ToString() + " seconds timeout");
                        }
                    }
                        
                    long currentCaseID = GetNextValue();
                    if (currentCaseID % intTestcaseRunnerThread == 0 || currentCaseID == totalEvents)
                    {
                        setLabelStatus(currentCaseID, totalEvents);
                    }
                });

                if (timedOut)
                {
                    additionalStatusDetails = "Completed with errors: use Testcase Runner again!";
                }
            }
            else
            {
                MessageBox.Show("Directory not found: " + strPreProcessedFolderFullPath);
            }
        }

        private void textBoxTemplate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                if (sender != null)
                    ((TextBox)sender).SelectAll();
            }
        }

        public void showCloseMoveEmail(object monitorSync, String strTestCaseFilePath, String strProcessedFolderFullPath, String strTestCaseFileName, int intWaitSeconds)
        {
            try
            {
                Microsoft.Office.Interop.Outlook.Application app;
                try { 
                    app = (Microsoft.Office.Interop.Outlook.Application)Marshal.GetActiveObject("Outlook.Application"); 
                }
                catch (Exception err) {
                    app = new Microsoft.Office.Interop.Outlook.Application();
                }


                var item = app.Session.OpenSharedItem(strTestCaseFilePath) as Microsoft.Office.Interop.Outlook.MailItem;
                // Show it
                item.Display(false);
                // Wait for it
                Thread.Sleep(intWaitSeconds * 1000);
                //Close it
                item.Close(Microsoft.Office.Interop.Outlook.OlInspectorClose.olDiscard);
                //Release it
                Marshal.ReleaseComObject(item);
                //Move it
                string strFinalFilePath = strProcessedFolderFullPath + strTestCaseFileName;
                while (File.Exists(strFinalFilePath))
                {
                    strFinalFilePath = strProcessedFolderFullPath + Guid.NewGuid() + "_" + strTestCaseFileName;
                }
                timestampLogger.log("Moving the " + strTestCaseFilePath + " file to " + strFinalFilePath);
                File.Move(strTestCaseFilePath, strFinalFilePath);
                payloadsProcessedTestcaseFilesLogger.log(SHA1FromFile(strFinalFilePath) + "," + strTestCaseFileName);
                lock (monitorSync)
                {
                    Monitor.Pulse(monitorSync);
                }
            }
            catch (Exception e)
            {
                timestampLogger.log("Error in showCloseMoveEmail(): " + e.StackTrace);
            }
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            browseFolder(textBoxWorkPath);
        }

        private void buttonBrowseSuffix_Click(object sender, EventArgs e)
        {
            browseFileName(textBoxSuffix);
        }

        private void buttonBrowsePrefix_Click(object sender, EventArgs e)
        {
            browseFileName(textBoxPrefix);
        }

        private void buttonBrowseFormula_Click(object sender, EventArgs e)
        {
            browseFileName(textBoxFormula);
        }

        private void buttonBrowseSchemes_Click(object sender, EventArgs e)
        {
            browseFileName(textBoxSchemes);
        }

        private void buttonBrowseTargets_Click(object sender, EventArgs e)
        {
            browseFileName(textBoxTargets);
        }

        private void buttonBrowseSpecialFormula_Click(object sender, EventArgs e)
        {
            browseFileName(textBoxSpecialFormula);
        }

        private void browseFileName(TextBox textBox)
        {
            openFileDialog1.Multiselect = false;
            openFileDialog1.InitialDirectory = System.IO.Path.GetDirectoryName(textBox.Text);
            openFileDialog1.FileName = "";

            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox.Text = openFileDialog1.FileName;
            }
        }

        private void browseFolder(TextBox textBox)
        {
            folderBrowserDialog1.SelectedPath = textBox.Text;
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void FormOutlookMailApp_Load(object sender, EventArgs e)
        {
            string currentPath = Directory.GetCurrentDirectory();
            string currentResourcePath = currentPath + "\\resources\\";
            textBoxPrefix.Text = currentResourcePath + "prefix.txt";
            textBoxSuffix.Text = currentResourcePath + "suffix.txt";
            textBoxFormula.Text = currentResourcePath + "formula.txt";
            textBoxSchemes.Text = currentResourcePath + "schemes.txt";
            textBoxTargets.Text = currentResourcePath + "targets-all.txt";
            textBoxSpecialFormula.Text = currentResourcePath + "special_schemes_formula.txt";
            textBoxWorkPath.Text = currentPath + "\\work\\";
            toolStripStatusLabelStatus.Text = "Ready";
            toolStripStatusLabelCommandName.Text = "";
            toolStripStatusLabelAdditionalStatus.Text = "";
        }

        private void buttonImportTemplate_Click(object sender, EventArgs e)
        {
            openFileDialog1.Multiselect = false;
            openFileDialog1.InitialDirectory = Directory.GetCurrentDirectory() + "\\resources\\";
            openFileDialog1.FileName = "template.html";

            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                string file = openFileDialog1.FileName;
                try
                {
                    textBoxTemplate.Text = File.ReadAllText(file);

                }
                catch (IOException err)
                {
                    timestampLogger.log("Error in buttonImportTemplate_Click(): " + err.StackTrace);
                }

            }
        }

        private void buttonGenSamlpleTemplate_Click(object sender, EventArgs e)
        {
            toolStripStatusLabelCommandName.Text = "Generating a sample testcase";
            disableAllButtons();
            genrateSampleTemplate();
            resetBoxes();
        }




        public static long GetNextValue()
        {
            return Interlocked.Increment(ref caseID);
        }

        public static void ResetCaseID()
        {
            Interlocked.Exchange(ref caseID, 0);
        }

        private void setLabelStatus(long currentCounter, long totalEvents)
        {
            int percent = 0;
            if (totalEvents > 0)
            {
                percent = (int)Math.Abs(100 * currentCounter / totalEvents);
            }
            else if (totalEvents == 0)
            {
                percent = 100;
            }
            string strStatus = currentCounter.ToString() + "/" + totalEvents.ToString() + " - %" + percent.ToString();

            if (this.statusStrip.InvokeRequired)
            {
                this.statusStrip.BeginInvoke((MethodInvoker)delegate()
                {

                    this.labelStatus.Text = strStatus;
                    this.progressBarStatus.Value = percent;

                });
            }
            else
            {
                this.labelStatus.Text = strStatus;
                this.progressBarStatus.Value = percent;
            }
        }


        public void ResumePause()
        {
            if (Interlocked.Read(ref paused) == 1)
            {
                Interlocked.Exchange(ref paused, 0);
                mre.Set();
                buttonControl.Text = "Click to Pause";
                toolStripStatusLabelStatus.Text = "Running";
            }
            else
            {
                Interlocked.Exchange(ref paused, 1);
                mre.Reset();
                buttonControl.Text = "Click to Resume";
                toolStripStatusLabelStatus.Text = "Paused";
            }
        }

        public static int StrToIntDef(string s, int @default)
        {
            int number;
            if (int.TryParse(s, out number))
                return number;
            return @default;
        }

        public static string[] UniqueTextArrayFromFile(string strFilePath)
        {
            string[] readText = File.ReadAllLines(strFilePath);
            return FixEmptyStringArray(RemoveDuplicates(readText));
        }

        public static string[] RemoveDuplicates(string[] s)
        {
            HashSet<string> set = new HashSet<string>(s);
            string[] result = new string[set.Count];
            set.CopyTo(result);
            return result;
        }

        public static string[] FixEmptyStringArray(string[] arrInput)
        {
            if (arrInput.Length == 0)
            {
                Array.Resize(ref arrInput, 1);
                arrInput[0] = "";
            }
            return arrInput;
        }

        private static string MakeValidFileName(string name)
        {
            string invalidChars = System.Text.RegularExpressions.Regex.Escape(new string(System.IO.Path.GetInvalidFileNameChars()));
            string invalidRegStr = string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);
            return System.Text.RegularExpressions.Regex.Replace(name, invalidRegStr, "_");
        }

        private void disableAllButtons()
        {
            buttonGenNormalFormula.Enabled = false;
            buttonGenCustomSchemesFormula.Enabled = false;
            buttonTestcaseRunner.Enabled = false;
            checkBoxAutoRun.Enabled = false;
            textBoxProcessTimeout.Enabled = false;
            buttonControl.Enabled = true;
            groupBoxThreads.Enabled = false;
            toolStripStatusLabelStatus.Text = "Running";
            textBoxKeepOpen.Enabled = false;
            additionalStatusDetails = "";
            this.TopMost = true;
            ResetCaseID();
        }

        private void resetBoxes()
        {
            if (this.buttonGenNormalFormula.InvokeRequired)
            {
                this.buttonGenNormalFormula.BeginInvoke((MethodInvoker)delegate()
                {
                    buttonGenNormalFormula.Enabled = true;
                    buttonGenCustomSchemesFormula.Enabled = true;
                    buttonTestcaseRunner.Enabled = true;
                    checkBoxAutoRun.Enabled = true;
                    textBoxProcessTimeout.Enabled = true;
                    buttonControl.Enabled = false;
                    textBoxKeepOpen.Enabled = true;
                    groupBoxThreads.Enabled = true;
                    toolStripStatusLabelStatus.Text = "Completed";
                    buttonControl.Text = "Click to Pause";
                    toolStripStatusLabelAdditionalStatus.Text = additionalStatusDetails;
                    this.TopMost = false;
                });
            }
            else
            {
                buttonGenNormalFormula.Enabled = true;
                buttonGenCustomSchemesFormula.Enabled = true;
                buttonTestcaseRunner.Enabled = true;
                checkBoxAutoRun.Enabled = true;
                textBoxProcessTimeout.Enabled = true;
                buttonControl.Enabled = false;
                textBoxKeepOpen.Enabled = true;
                groupBoxThreads.Enabled = true;
                toolStripStatusLabelStatus.Text = "Completed";
                buttonControl.Text = "Click to Pause";
                toolStripStatusLabelAdditionalStatus.Text = additionalStatusDetails;
                this.TopMost = false;
            }

            
            additionalStatusDetails = "";
            Interlocked.Exchange(ref paused, 0);
            ResetCaseID();
        }
        private void cmdGenNormalFormula_Click(object sender, EventArgs e)
        {
            //disableAllButtons(checkBoxAutoRun.Checked); there is a bug in running testcase runner immediately after generation TODO
            disableAllButtons();
            toolStripStatusLabelCommandName.Text = "Generating normal testcases";
            new Thread(delegate()
            {
                genrateTestCases();
                resetBoxes();
            }).Start();
        }

        private void buttonGenCustomSchemesFormula_Click(object sender, EventArgs e)
        {
            //disableAllButtons(checkBoxAutoRun.Checked); there is a bug in running testcase runner immediately after generation TODO
            disableAllButtons();
            toolStripStatusLabelCommandName.Text = "Generating special testcases";
            new Thread(delegate()
            {
                genrateSpecialTestCases();
                resetBoxes();
            }).Start();
        }

        private void buttonControl_Click(object sender, EventArgs e)
        {
            ResumePause();
        }

        private void buttonTestcaseRunner_Click(object sender, EventArgs e)
        {
            disableAllButtons();
            buttonTestcaseRunner.Enabled = false;
            toolStripStatusLabelCommandName.Text = "Iterating testcases";
            new Thread(delegate()
            {
                RunTestCases(); 
                resetBoxes();
            }).Start();
        }

        private void checkBoxAutoRun_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxAutoRun.Checked)
            {
                buttonTestcaseRunner.Enabled = false;
            }
            else
            {
                buttonTestcaseRunner.Enabled = true;
            }
        }

        public static string SHA1FromString(string stringToHash)
        {
            using (var sha1 = new SHA1Managed())
            {
                return BitConverter.ToString(sha1.ComputeHash(Encoding.UTF8.GetBytes(stringToHash))).Replace("-", string.Empty);
            }
        }

        public static string SHA1FromFile(string strFilePath){

            string sendCheckSum = "";
            using (FileStream stream = File.OpenRead(strFilePath))
            {
                using (SHA1Managed sha = new SHA1Managed())
                {
                    byte[] checksum = sha.ComputeHash(stream);
                    sendCheckSum = BitConverter.ToString(checksum)
                        .Replace("-", string.Empty);
                }
            }
            return sendCheckSum;
        }



        private void killProcess(string strProcessName, int processTimeoutInSecond)
        {
            DateTime endTime = DateTime.Now;
            TimeSpan span = endTime.Subtract(lastKillRequest);
            int killDiff = 0;
            if(processTimeoutInSecond!=-1)
                killDiff = (processTimeoutInSecond > 10) ? processTimeoutInSecond : 10; // so it can be buggy if someone set the processTimeoutInSecond to something lower than 10
            if (span.Seconds > killDiff)
            {
                foreach (var process in Process.GetProcessesByName(strProcessName))
                {
                    try
                    {
                        process.Kill();
                        lastKillRequest = DateTime.Now;
                    }
                    catch (Exception e)
                    {
                        if (timestampLogger != null)
                        {
                            timestampLogger.log("Error in killProcess(): " + e.StackTrace);
                        }
                    }
                    
                }
            }
        }

        private void buttonKillOutlook_Click(object sender, EventArgs e)
        {
            killProcess("OUTLOOK",-1);
        }


    }
}
