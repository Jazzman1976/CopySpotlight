using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Threading.Tasks;
using System.Timers;

namespace CopySpotlight
{
    public static class FileExtensions
    {
        public static Task DeleteAsync(this FileInfo fi)
        {
            return Task.Factory.StartNew(() => fi.Delete());
        }
    }

    public partial class CopySpotlight : ServiceBase
    {
        private static Timer _timer;

        public CopySpotlight()
        {
            // call base default method
            InitializeComponent();

            // init eventlog
            string srcName = "CopySpotlightSource";
            string logName = "CopySpotlightLog";
            eventLog1 = new EventLog();
            if (!EventLog.SourceExists(srcName))
            {
                EventLog.CreateEventSource(srcName, logName);
            }
            eventLog1.Source = srcName;
            eventLog1.Log = logName;

            // init timer
            double time = 1000 * 60 * 60 * 6; // ms, s, h, 6 => 6 hours
            _timer = new Timer();
            _timer.Interval = time;
            _timer.Elapsed += new ElapsedEventHandler(this.OnTimer);

            // react on shutdown and start in windows fast start mode
            this.CanHandlePowerEvent = true;

            eventLog1.WriteEntry("New CopySpotlight object created.");
        }

        protected override void OnStart(string[] args)
        {
            DoCopySpotlight();
            _timer.Start();
            eventLog1.WriteEntry("Service has been started.");
        }

        private void OnTimer(object sender, ElapsedEventArgs args)
        {
            eventLog1.WriteEntry("Timed events are executed.");
            DoCopySpotlight();
        }

        private async void DoCopySpotlight()
        {
            // collects log information
            List<string> log = new List<string>();

            // find target folder
            string userPicturesFolderPath = Environment.ExpandEnvironmentVariables(
                "%UserProfile%\\Pictures\\");
            string picturesFolderPath = null;
            if (Directory.Exists(userPicturesFolderPath))
            {
                picturesFolderPath = userPicturesFolderPath;
                log.Add("Using local user picture folder (not onedrive).");
            }
            else
            {
                string onedrivePicturesFolderPath = Environment
                    .ExpandEnvironmentVariables("%OneDrive%\\Pictures\\");
                if (Directory.Exists(onedrivePicturesFolderPath))
                {
                    picturesFolderPath = onedrivePicturesFolderPath;
                    log.Add("Using OneDrive picture folder.");
                }
            }
            if (string.IsNullOrEmpty(picturesFolderPath))
            {
                eventLog1.WriteEntry("No target Picture folder found.\n"
                    + string.Join("\n", log), EventLogEntryType.Warning);
                return;
            }
            picturesFolderPath += "CopySpotlight\\";
            log.Add($"The used pictures folder path is: {picturesFolderPath}");
            Directory.CreateDirectory(picturesFolderPath);

            // get files
            string spotlightFolderPath = Environment.ExpandEnvironmentVariables(
                "%LocalAppData%\\Packages\\"
                + "Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\\LocalState\\Assets\\");
            log.Add("I check for new files...");
            string[] files = Directory.GetFiles(spotlightFolderPath);
            if (!files.Any())
            {
                eventLog1.WriteEntry("No Spotlight files found.\n"
                    + string.Join("\n", log), EventLogEntryType.Warning);
                return;
            }
            log.Add($"{files.Length} files were found in path {spotlightFolderPath}");

            // copy files
            log.Add("I start main process to copy the new files to the target locations.");
            bool hasNewFiles = false;
            bool teamsUpdate = false;
            try
            {
                foreach (string file in files)
                {
                    log.Add("\n" + $"I start investigation of file {file}.");

                    #region copy all new files ----------------------------------------------------

                    string[] filePathParts = file.Split("\\".ToCharArray());
                    string fileName = filePathParts.Last();
                    string jpgFile = picturesFolderPath + fileName + ".jpg";
                    log.Add($"I use the filename '{jpgFile}' in target locations.");

                    // continue with next file if file still exists
                    if (File.Exists(jpgFile))
                    {
                        log.Add($"The file '{jpgFile}' still exists. I continue with next file.");
                        continue;
                    }

                    // skip if not larger than 200 KB
                    long minFilesizeToCopy = 200 * 1000;
                    FileInfo infoSpotlightFile = new FileInfo(file);
                    if (infoSpotlightFile.Length < minFilesizeToCopy)
                    {
                        log.Add($"The file {file} is not larger than "
                            + $"{minFilesizeToCopy / 1000} KB. I will skip it.");
                        continue;
                    }

                    // copy all new file to jpg file in target folder
                    // await it, because it could not be read otherwise in next code block
                    // from now on jpgFile exists at target location
                    log.Add($"Start copiing file '{file}' to '{jpgFile}'...");
                    await CopyFileAsync(file, jpgFile);
                    hasNewFiles = true;
                    log.Add($"Copiing file '{file}' to '{jpgFile}' has been finished.");

                    // read image details
                    log.Add($"Reading file information of file '{jpgFile}'...");
                    bool isHdLandscape = false;
                    bool isHdPortrait = false;
                    bool isJpgImage = false;
                    try
                    {
                        FileInfo info = new FileInfo(jpgFile);
                        using (Image jpgImage = Image.FromFile(jpgFile))
                        {
                            log.Add("Image object has been created.");
                            if (jpgImage != null && info.Length >= minFilesizeToCopy)
                            {
                                isJpgImage
                                    = jpgImage.RawFormat.Equals(ImageFormat.Jpeg)
                                    && jpgImage.Width > 0
                                    && jpgImage.Height > 0;

                                if (isJpgImage)
                                {
                                    isHdLandscape = jpgImage.Width == 1920 && jpgImage.Height == 1080;
                                    isHdPortrait = jpgImage.Width == 1080 && jpgImage.Height == 1920;
                                }

                                jpgImage.Dispose(); // unlock file
                            }
                        }
                        log.Add($"It {(isJpgImage ? "is" : "is not")} an image.");
                        log.Add($"It {(isHdLandscape ? "is" : "is not")} a HD-Landscape image.");
                        log.Add($"It {(isHdPortrait ? "is" : "is not")} a HD-Portrait image.");

                        // continue with next image if not an HD image 
                        if (!isHdLandscape && !isHdPortrait)
                        {
                            log.Add("Deleting the copied file; it is not a HD image.");
                            await info.DeleteAsync();
                            log.Add("Continue with next file.");
                            continue;
                        }
                    }
                    catch (Exception exc)
                    {
                        log.Add($"An exception occured: {exc.Message}");
                    }

                    #endregion


                    #region update latest landscape or portrait -----------------------------------

                    try
                    {
                        // check if landscape or portrait
                        string latestJpgFileLandscape = picturesFolderPath + "latest-landscape.jpg";
                        string latestJpgFilePortrait = picturesFolderPath + "latest-portrait.jpg";
                        string latestTarget = isHdLandscape
                            ? latestJpgFileLandscape
                            : latestJpgFilePortrait;

                        // if no target file exists create it and continue with next image
                        if (!File.Exists(latestTarget))
                        {
                            // await it, because it could not be read otherwise in next code block
                            await CopyFileAsync(jpgFile, latestTarget);
                            log.Add($"The file {latestTarget} didn't exist and has been created.");
                        }

                        // if a target file exists, overwrite it if newer
                        else if (File.GetCreationTime(latestTarget).Ticks
                            <= File.GetCreationTime(jpgFile).Ticks)
                        {
                            //File.Copy(jpgFile, latestTarget, true);
                            await CopyFileAsync(jpgFile, latestTarget);
                            log.Add($"The exising file {latestTarget} has been updated.");
                        }
                    }
                    catch (Exception exc)
                    {
                        log.Add("An error occured while working on the latest image in picture " +
                            $"folder. I continue with next image.\nException: {exc.Message}");
                        continue;
                    }

                    #endregion


                    #region update background image to use in MS Teams ----------------------------

                    try
                    {
                        // skip if not landscape or Teams path is not found
                        string teamsFolderPath = Environment.ExpandEnvironmentVariables(
                            "%appdata%\\Microsoft\\Teams\\");
                        if (!isHdLandscape || !Directory.Exists(teamsFolderPath))
                        {
                            if (!isHdLandscape)
                            {
                                log.Add("The image is not relevant as Teams background.");
                            }
                            else
                            {
                                log.Add("Can't find MS Teams installation.");
                            }
                            log.Add("I will continue with the next image.");
                            continue;
                        }

                        // Teams path and files
                        string teamsBackgroundsFolderPath = Environment.ExpandEnvironmentVariables(
                            "%appdata%\\Microsoft\\Teams\\Backgrounds\\");
                        string teamsUploadsFolderPath = Environment.ExpandEnvironmentVariables(
                            "%appdata%\\Microsoft\\Teams\\Backgrounds\\Uploads\\");
                        string latestBackgroundTeamsFile
                            = teamsUploadsFolderPath + "latest-landscape.jpg";
                        string latestBackgroundTeamsThumbFile
                            = teamsUploadsFolderPath + "latest-landscape_thumb.jpg";

                        // assure folders
                        Directory.CreateDirectory(teamsBackgroundsFolderPath);
                        Directory.CreateDirectory(teamsUploadsFolderPath);

                        // create files if not existing
                        if (!File.Exists(latestBackgroundTeamsFile))
                        {
                            // await it, because it could not be read otherwise in next code block
                            await CopyFileAsync(jpgFile, latestBackgroundTeamsFile);
                            log.Add($"The file {latestBackgroundTeamsFile} didn't exist and has " +
                                $"been created.");

                            // create thumbnail img
                            Image latestBackgroundImage = Image.FromFile(latestBackgroundTeamsFile);
                            if (File.Exists(latestBackgroundTeamsThumbFile))
                            {
                                File.Delete(latestBackgroundTeamsThumbFile);
                                log.Add($"I have deleted the existing file " +
                                    $"{latestBackgroundTeamsThumbFile} before creating a new " +
                                    $"version of that file.");
                            }
                            Image thumbImage = new Bitmap(latestBackgroundImage, new Size(280, 158));
                            thumbImage.Save(latestBackgroundTeamsThumbFile);
                            log.Add($"A new thumbnail image {latestBackgroundTeamsThumbFile} " +
                                $"has been created.");

                            // raise flag
                            teamsUpdate = true;
                        }

                        // if a target file exists, overwrite it if newer
                        else if (File.GetCreationTime(latestBackgroundTeamsFile).Ticks
                            <= File.GetCreationTime(jpgFile).Ticks)
                        {
                            //File.Copy(jpgFile, latestTarget, true);
                            await CopyFileAsync(jpgFile, latestBackgroundTeamsFile);

                            // update thumbnail img
                            Image latestBackgroundImage = Image.FromFile(latestBackgroundTeamsFile);
                            if (File.Exists(latestBackgroundTeamsThumbFile))
                            {
                                File.Delete(latestBackgroundTeamsThumbFile);
                            }
                            Image thumbImage = new Bitmap(latestBackgroundImage, new Size(280, 158));
                            thumbImage.Save(latestBackgroundTeamsThumbFile);

                            teamsUpdate = true;
                        }
                    }
                    catch (Exception exc)
                    {
                        log.Add($"An error occured while working on MS Teams background image. " +
                            $"I will skip it and continue with the next image.\n" +
                            $"Error: {exc.Message}");
                        continue;
                    }

                    #endregion
                }
            }
            catch (Exception exc)
            {
                eventLog1.WriteEntry($"An error occured: {exc.Message} \n"
                    + string.Join("\n", log), EventLogEntryType.Error);
            }

            // log info
            string result = hasNewFiles
                ? "New files were found and copied to target location."
                : "No new files found to copy.";
            string teams = teamsUpdate
                ? $"Background image in Teams has been added/updated."
                : "Background image in Teams has not been added/updated.";
            eventLog1.WriteEntry(
                result + "\n"
                + teams + "\n"
                + "Spotlight folder path: " + spotlightFolderPath + "\n"
                + "Target Picture folder path: " + picturesFolderPath + "\n\n"
                + string.Join("\n", log));
        }

        protected override void OnContinue()
        {
            base.OnContinue();
            DoCopySpotlight();
            _timer.Start();
            eventLog1.WriteEntry("Service has been continued.");
        }

        /// <summary>
        /// This is called when I turn on the comuter and windows uses the "fast boot" option.<br/>
        /// It is also called after wake up from sleep.<br/>
        /// Services are set to some kind of sleep in this mode and are not turned off/on.<br/>
        /// Requires: "this.CanHandlePowerEvent = true" in the constructor.
        /// </summary>
        protected override bool OnPowerEvent(PowerBroadcastStatus powerStatus)
        {
            switch (powerStatus)
            {
                case PowerBroadcastStatus.ResumeSuspend:
                    DoCopySpotlight();
                    _timer.Start();
                    break;
            }
            eventLog1.WriteEntry(
                $"OnPowerEvent: {Enum.GetName(typeof(PowerBroadcastStatus), powerStatus)}");
            return base.OnPowerEvent(powerStatus);
        }

        protected override void OnStop()
        {
            _timer.Stop();
            eventLog1.WriteEntry("Service has been stopped.");
        }

        private async Task CopyFileAsync(string sourcePath, string destinationPath)
        {
            using (Stream source = File.Open(sourcePath, FileMode.Open))
            {
                using (Stream destination = File.Create(destinationPath))
                {
                    await source.CopyToAsync(destination);
                }
            }
        }
    }
}
