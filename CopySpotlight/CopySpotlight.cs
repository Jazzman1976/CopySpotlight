using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Threading.Tasks;
using System.Timers;

namespace CopySpotlight
{
    public partial class CopySpotlight : ServiceBase
    {
        private Timer _timer;

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
            // find target folder
            string userPicturesFolderPath = Environment.ExpandEnvironmentVariables(
                "%UserProfile%\\Pictures\\");
            string picturesFolderPath = null;
            if (Directory.Exists(userPicturesFolderPath))
            {
                picturesFolderPath = userPicturesFolderPath;
            }
            else
            {
                string onedrivePicturesFolderPath = Environment
                    .ExpandEnvironmentVariables("%OneDrive%\\Pictures\\");
                if (Directory.Exists(onedrivePicturesFolderPath))
                {
                    picturesFolderPath = onedrivePicturesFolderPath;
                }
            }
            if (string.IsNullOrEmpty(picturesFolderPath))
            {
                eventLog1.WriteEntry("No target Picture folder found.", EventLogEntryType.Warning);
                return;
            }
            picturesFolderPath += "CopySpotlight\\";
            Directory.CreateDirectory(picturesFolderPath);

            // get files
            string spotlightFolderPath = Environment.ExpandEnvironmentVariables(
                "%LocalAppData%\\Packages\\"
                + "Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\\LocalState\\Assets\\");
            string[] files = Directory.GetFiles(spotlightFolderPath);
            if (!files.Any())
            {
                eventLog1.WriteEntry("No Spotlight files found.", EventLogEntryType.Warning);
                return;
            }

            // copy files
            bool hasNewFiles = false;
            bool teamsUpdate = false;
            try
            {
                foreach (string file in files)
                {
                    #region copy all new files ----------------------------------------------------

                    string[] filePathParts = file.Split("\\".ToCharArray());
                    string fileName = filePathParts.Last();
                    string jpgFile = picturesFolderPath + fileName + ".jpg";

                    // continue with next file if file still exists
                    if (File.Exists(jpgFile))
                    {
                        continue;
                    }

                    // copy all new file to jpg file in target folder
                    // await it, because it could not be read otherwise in next code block
                    // from now on jpgFile exists at target location
                    await CopyFileAsync(file, jpgFile);
                    hasNewFiles = true;

                    // parse to Image object
                    Image jpgImage = Image.FromFile(jpgFile);
                    if (jpgImage == null)
                    {
                        // assure image
                        continue;
                    }

                    // continue with next image if not an HD image
                    bool isHdLandscape = jpgImage.Width == 1920 && jpgImage.Height == 1080;
                    bool isHdPortrait = jpgImage.Width == 1080 && jpgImage.Height == 1920;
                    jpgImage.Dispose(); // unlock file
                    if (!isHdLandscape && !isHdPortrait)
                    {
                        File.Delete(jpgFile);
                        continue;
                    }

                    #endregion


                    #region update latest landscape or portrait -----------------------------------

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
                    }

                    // if a target file exists, overwrite it if newer
                    else if (File.GetCreationTime(latestTarget).Ticks
                        <= File.GetCreationTime(jpgFile).Ticks)
                    {
                        //File.Copy(jpgFile, latestTarget, true);
                        await CopyFileAsync(jpgFile, latestTarget);
                    }

                    #endregion


                    #region update background image to use in MS Teams ----------------------------

                    // skip if not landscape or Teams path is not found
                    string teamsFolderPath = Environment.ExpandEnvironmentVariables(
                        "%appdata%\\Microsoft\\Teams\\");
                    if (!isHdLandscape || !Directory.Exists(teamsFolderPath))
                    {
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

                        // create thumbnail img
                        Image latestBackgroundImage = Image.FromFile(latestBackgroundTeamsFile);
                        if (File.Exists(latestBackgroundTeamsThumbFile))
                        {
                            File.Delete(latestBackgroundTeamsThumbFile);
                        }
                        Image thumbImage = new Bitmap(latestBackgroundImage, new Size(280, 158));
                        thumbImage.Save(latestBackgroundTeamsThumbFile);

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

                    #endregion
                }
            }
            catch (Exception exc)
            {
                eventLog1.WriteEntry(exc.Message, EventLogEntryType.Error);
            }

            // log info
            string result = hasNewFiles
                ? "New files were found and copied to target location."
                : "No new files found to copy.";
            string teams = teamsUpdate
                ? $"Background image in Teams has been added/updated."
                : "Background image in Teams has not been added/updated.";
            eventLog1.WriteEntry(
                result + "\n\r" 
                + teams + "\n\r"
                + "Spotlight folder path: " + spotlightFolderPath + "\n\r"
                + "Target Picture folder path: " + picturesFolderPath);
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
