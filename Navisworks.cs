using System;
using System.IO;
using System.Linq;
using System.Text;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Controls;
using ReviewConverter.Properties;

using NwApplication = Autodesk.Navisworks.Api.Automation.NavisworksApplication;

namespace ReviewConverter
{
    class NavisWorks
    {
        public static void NwdPublish(string input, string output, int vers = 2012)
        {
            // Set the job variables
            var startDir = Directory.GetParent(input);
            var jobName = Path.GetFileName(input);
            var jobTitle = Path.GetFileNameWithoutExtension(input);
            var jobDate = new FileInfo(input).CreationTime.ToString("yyyy-MM-dd");
            var tempDir = Path.Combine(Environment.GetEnvironmentVariable("TEMP"), jobTitle);
            var tempOutput = Path.Combine(tempDir, Path.GetFileName(output));
            if (jobName == null)
            {
                Program.ConsoleLog("Error setting job variables");
                return;
            }

            // Clear the temp directory
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, true);
            Directory.CreateDirectory(tempDir);

            // Get all the working files
            var workingFiles = startDir.EnumerateFiles(jobTitle + "*", SearchOption.TopDirectoryOnly).ToList();

            // Remove NWC and NWD files
            for (var i = workingFiles.Count - 1; i >= 0; i--)
            {
                if (!workingFiles[i].Extension.Equals(".nwc", StringComparison.CurrentCultureIgnoreCase) &&
                    !workingFiles[i].Extension.Equals(".nwd", StringComparison.CurrentCultureIgnoreCase))
                    continue;
                workingFiles.RemoveAt(i);
            }

            // Copy the work files to the temp directory
            Program.ConsoleLog("Copying working files...");
            workingFiles.ForEach(x => x.CopyTo(Path.Combine(tempDir, x.Name)));

            // Initialize the API
            Program.ConsoleLog("Initializing NavisWorks...");
            Program.ConsoleFreeze("Loading Navisworks modules...");
            ApplicationControl.ApplicationType = ApplicationType.SingleDocument;
            ApplicationControl.Initialize();
            if (!ApplicationControl.IsInitialized)
            {
                Program.ConsoleLog("Unable to initialize Navisworks!");
                return;
            }
            Program.ConsoleThaw();

            // Get the publish properties from the provided dictionary
            var publishProps = GetPublishProperties(string.Format("{0}_{1}", jobName, jobDate));
            if (publishProps == null)
            {
                Program.ConsoleLog("Failed to create publish properties!");
                goto Terminate;
            }

            // Create a document control
            using (var docControl = new DocumentControl())
            {
                // Set the control as the primary document
                docControl.SetAsMainDocument();
                var nwDoc = docControl.Document;

                // Try opening the document
                Program.ConsoleLog("Opening {0}...", jobName);
                try
                { nwDoc.OpenFile(Path.Combine(tempDir, jobName)); }
                catch (DocumentFileException ex)
                {
                    Program.ConsoleLog(ex.Message);
                    goto Terminate;
                }

                // Try publishing the document
                Program.ConsoleLog("Publishing {0}...", Path.GetFileName(output));
                try { nwDoc.PublishFile(tempOutput, publishProps); }
                catch (DocumentFileException ex)
                {
                    Program.ConsoleLog(ex.Message);
                    goto Terminate;
                }

                // Convert to a lower format if necessary
                if (vers < 2012)
                {
                    Program.ConsoleLog("Converting {0} to {1} format...", Path.GetFileName(output), vers);
                    try
                    {
                        nwDoc.OpenFile(tempOutput);
                        nwDoc.SaveFile(tempOutput, GetVersion(vers));
                    }
                    catch (DocumentFileException ex)
                    {
                        Program.ConsoleLog(ex.Message);
                    }
                }
            }

            // Terminate the API
            Terminate:
            Program.ConsoleLog("Closing NavisWorks...");
            Program.ConsoleFreeze("Unloading Navisworks modules...");
            ApplicationControl.Terminate();
            Program.ConsoleThaw();

            // Copy the file to the final location
            //Program.ConsoleLog("Copying {0} to {1}", Path.GetFileName(output), Path.GetDirectoryName(output));
            try { File.Copy(tempOutput, output, true); }
            catch (IOException ex) { Program.ConsoleLog(ex.Message); }
        }
        public static void MergeFiles(string[] sources, string destination, int vers = 2012)
        {
            // Initialize the API
            Program.ConsoleLog("Initializing NavisWorks...");
            ApplicationControl.ApplicationType = ApplicationType.SingleDocument;
            ApplicationControl.Initialize();

            // Create a document control
            using (var nwDocControl = new DocumentControl())
            {
                // Set the control as the primary document
                nwDocControl.SetAsMainDocument();
                var nwDoc = nwDocControl.Document;

                // Merge
                Program.ConsoleLog("Merging files...");
                if (nwDoc.TryMergeFiles(sources))
                {
                    Program.ConsoleLog("Merge Successful!");

                    // Save
                    Program.ConsoleLog("Saving {0}...", destination);
                    if (nwDoc.TrySaveFile(destination, GetVersion(vers)))
                    {
                        Program.ConsoleLog("{0} Saved successfully!", destination);
                    }
                }
            }

            // Terminate the API
            Program.ConsoleLog("Closing NavisWorks...");
            ApplicationControl.Terminate();
        }

        private static PublishProperties GetPublishProperties(string jobTitle)
        {
            var returnProperties = new PublishProperties
            {
                Author = Settings.Default.NWD_AUTHOR,
                Comments = Settings.Default.NWD_COMMENTS,
                Copyright = Settings.Default.NWD_COPYRIGHT,
                Keywords = Settings.Default.NWD_KEYWORDS,
                PublishDate = DateTime.Now,
                PublishedFor = Settings.Default.NWD_PUBLISHEDFOR,
                Publisher = Settings.Default.NWD_PUBLISHER,
                Subject = Settings.Default.NWD_SUBJECT,
                Title = jobTitle,
                AllowResave = true
            };

            return returnProperties;
        }
        private static DocumentFileVersion GetVersion(int vers)
        {
            switch (vers)
            {
                case 2010: return DocumentFileVersion.Navisworks2010;
                case 2011: return DocumentFileVersion.Navisworks2011;
                case 2012: return DocumentFileVersion.Navisworks2012;
                default: return DocumentFileVersion.Current;
            }
        }
        public static DocumentFileVersion GetVersion(string filePath)
        {
            // Create the byte buffer
            var buffer = new byte[1024];

            // Create the search pattern (Navisworks:)
            byte[] pattern = { 110, 97, 118, 105, 115, 119, 111, 114, 107, 115, 58 };

            // Create the version byte array
            var version = new byte[3];

            // Load the NWD file
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                // Read the bytes into the buffer
                fs.Read(buffer, 0, buffer.Length);

                // Get the pattern start position
                int patternStart = IndexOfSequence(buffer, pattern, 0);

                // If the pattern wasn't found in the header
                if (patternStart == -1)
                {
                    // Go the the end of the file
                    fs.Seek(-buffer.Length, SeekOrigin.End);

                    // Read the bytes into the buffer
                    fs.Read(buffer, 0, buffer.Length);

                    // Get the pattern start position
                    patternStart = IndexOfSequence(buffer, pattern, 0);

                    // Return default if the pattern is still not found
                    if (patternStart == -1) return DocumentFileVersion.Current;
                }

                // Get the version start position
                int versionStart = patternStart + pattern.Length;

                // Copy the version bytes (3) into the version byte array
                Array.Copy(buffer, versionStart, version, 0, version.Length);
            }

            // Get the integer representation of the version
            var vers = int.Parse(Encoding.ASCII.GetString(version));

            // Return the DocumentFileVersion parsed from the integer
            return (DocumentFileVersion)vers;
        }

        // Returns the start position of a byte pattern within a byte array
        private static int IndexOfSequence(byte[] buffer, byte[] pattern, int startIndex)
        {
            // Find a starting point matching the first pattern element
            var i = Array.IndexOf(buffer, pattern[0], startIndex);

            // Iterate through the length of the buffer
            while (i >= 0 && i <= buffer.Length - pattern.Length)
            {
                // Create a byte block of the pattern size
                var segment = new byte[pattern.Length];

                // Copy the bytes from the ith position into the segment
                Buffer.BlockCopy(buffer, i, segment, 0, pattern.Length);

                // Return the current position if the sequeence is a match
                if (segment.SequenceEqual(pattern))
                    return i;

                // Otherwise increment i a full pattern length
                i = Array.IndexOf(buffer, pattern[0], i + pattern.Length);
            }

            // Return no pattern found
            return -1;
        }
    }
}
