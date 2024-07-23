using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using System.IO;
using System.Reflection;
using Newtonsoft.Json.Linq;
using ExtensionBlocks;

namespace KNOCKOUT
{
    internal class KNOCKOUT
    {

        /// <summary>
        /// Represents an entry for an application with an ID and a Name.
        /// </summary>
        public class ApplicationEntry
        {
            /// <summary>
            /// Gets or sets the unique identifier for the application.
            /// </summary>
            public string ID { get; set; }

            /// <summary>
            /// Gets or sets the name of the application.
            /// </summary>
            public string Name { get; set; }
        }

        /// <summary>
        /// Entry point for the application. Extracts various types of recent files, folders, and device information.
        /// </summary>
        /// <param name="args">Command line arguments.</param>
        static void Main(string[] args)
        {
            // Print the tool banner
            PrintBanner();

            // Create an instance of RegistryHelper
            RegistryHelper registryHelper = new RegistryHelper();

            // Extract recently used Office files and folders
            ExtractRecentOfficeFilesAndFolders(registryHelper);

            // Extract recent file and folder shortcuts
            ExtractRecentLNKFilesAndFolders();

            // Extract recent URL files and folders
            ExtractRecentURLFilesAndFolders();

            // Extract recent Explorer files and folders using the registry
            ExtractRecentExplorerFilesAndFolders(registryHelper);

            // Extract frequently accessed files from Jump Lists
            ExtractFrequentFilesFromJumpLists();

            // Extract information about recently connected USB storage devices using the registry
            ExtractRecentlyConnectedUSBStorage(registryHelper);

            // Extract and print browser favorites
            ExtractBrowserFavorites();

            // Extract last run software from the UserAssist registry key
            ExtractUserAssist(registryHelper);
        }

        /// <summary>
        /// Prints the application banner to the console.
        /// </summary>
        private static void PrintBanner()
        {
            Console.WriteLine(@"
                ██   ██ ███    ██  ██████   ██████ ██   ██  ██████  ██    ██ ████████ 
                ██  ██  ████   ██ ██    ██ ██      ██  ██  ██    ██ ██    ██    ██    
                █████   ██ ██  ██ ██    ██ ██      █████   ██    ██ ██    ██    ██    
                ██  ██  ██  ██ ██ ██    ██ ██      ██  ██  ██    ██ ██    ██    ██    
                ██   ██ ██   ████  ██████   ██████ ██   ██  ██████   ██████     ██    
            ");
        }

        /// <summary>
        /// Extracts recently used office files and folders from the registry and prints them to the console.
        /// </summary>
        /// <param name="registryHelper">The helper which is used to access and query the registry.</param>
        /// <returns></returns>
        /// <exception cref="Exception">Thrown if a registry key could not be accessed.</exception>
        private static void ExtractRecentOfficeFilesAndFolders(RegistryHelper registryHelper)
        {
            Console.WriteLine($"[i] Extracting recently used Office files and folders");

            try
            {
                string registryPath = @"Software\Microsoft\Office";
                IntPtr hKey = registryHelper.OpenRegistryKey(RegistryHelper.HKEY_CURRENT_USER, registryPath);
                List<string> subKeys = registryHelper.EnumerateSubKeys(hKey);

                foreach (string subKey in subKeys)
                {
                    // Regex pattern to match office versions
                    string pattern = @"\d+\.\d+"; 
                    // Create a Regex object with the pattern
                    Regex regex = new Regex(pattern);
                    // Use the Matches method to find all matches in the input string
                    MatchCollection matches = regex.Matches(subKey);

                    foreach (Match match in matches)
                    {
                        Console.WriteLine($"[+] Found Office Version {match.Value} in the Registry");
                        string officeVersionKeyPath = registryPath + @"\" + match.Value;

                        IntPtr officeVersionKey = registryHelper.OpenRegistryKey(RegistryHelper.HKEY_CURRENT_USER, officeVersionKeyPath);
                        List<string> mruKeys = registryHelper.FindMRUKeys(officeVersionKey, officeVersionKeyPath);
                        List<string> mruValues = ParseMruData(mruKeys, registryHelper);
                        foreach (string mruValue in mruValues)
                        {
                            Console.WriteLine(mruValue);
                        }
                        registryHelper.CloseRegistryKey(officeVersionKey);
                    }
                }
                registryHelper.CloseRegistryKey(hKey);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        /// <summary>
        /// Extracts recently accessed documents, typed paths and launched executables from the registry and prints them to the console.
        /// </summary>
        /// <param name="registryHelper">The helper which is used to access and query the registry.</param>
        /// <exception cref="Exception">Thrown if a registry key could not be accessed.</exception>
        private static void ExtractRecentExplorerFilesAndFolders(RegistryHelper registryHelper)
        {
            // Output the action being performed to the console.
            Console.WriteLine("\n[i] Extracting recent explorer files and folders");

            // List of registry paths to be checked.
            List<string> registryPaths = new List<string>
            {
                //@"Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32\OpenSavePidlMRU", // Excluded for now as the parsing logic is more complex
                @"Software\Microsoft\Windows\CurrentVersion\Explorer\RecentDocs",
                @"Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU",
                @"Software\Microsoft\Windows\CurrentVersion\Explorer\TypedPaths"
            };

            // Iterate through each registry path.
            foreach (string path in registryPaths)
            {
                try
                {
                    // Special handling for RecentDocs since its format is more readable after specific treatment.
                    // This treatment process is inspired by the RegRipper tool.
                    if (path.Contains("RecentDocs"))
                    {
                        // Open the registry key.
                        IntPtr hKey = registryHelper.OpenRegistryKey(RegistryHelper.HKEY_CURRENT_USER, path);

                        // Retrieve the binary values from the registry key.
                        Dictionary<string, string> subValues = registryHelper.RetrieveRegistryValuesBinary(hKey);
                        foreach (KeyValuePair<string, string> subValue in subValues)
                        {
                            // Output the value to the console.
                            Console.WriteLine(subValue.Value); 
                        }

                        // Process all subkeys as they represent file extensions.
                        try
                        {
                            List<string> subKeys = registryHelper.EnumerateSubKeys(hKey);

                            foreach (string subKey in subKeys)
                            {
                                Console.WriteLine($"\n[i] Processing extension '{subKey}'");

                                // Construct the path to the subkey.
                                string subSubKeyPath = $"{path}\\{subKey}";

                                // Open the subkey.
                                IntPtr hSubKey = registryHelper.OpenRegistryKey(RegistryHelper.HKEY_CURRENT_USER, subSubKeyPath);

                                // Retrieve the binary values from the subkey.
                                Dictionary<string, string> subExtensionValues = registryHelper.RetrieveRegistryValuesBinary(hSubKey);
                                foreach (KeyValuePair<string, string> subExtensionValue in subExtensionValues)
                                {
                                    // Output the value to the console.
                                    Console.WriteLine(subExtensionValue.Value); 
                                }

                                // Close the subkey.
                                registryHelper.CloseRegistryKey(hSubKey);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Output error message if there's an issue accessing the registry path.
                            Console.WriteLine($"Error accessing registry path {path}: {ex.Message}");
                        }

                        // Close the main registry key.
                        registryHelper.CloseRegistryKey(hKey);
                    }
                    else if (path.Contains("RunMRU"))
                    {
                        // Handle RunMRU registry key.
                        Console.WriteLine($"\n[i] Extracting RunMRU keys from explorer");

                        // Open the registry key.
                        IntPtr hKey = registryHelper.OpenRegistryKey(RegistryHelper.HKEY_CURRENT_USER, path);

                        // Retrieve the string values from the registry key.
                        Dictionary<string, string> subValues = registryHelper.RetrieveRegistryValues(hKey);
                        foreach (KeyValuePair<string, string> subValue in subValues)
                        {
                            // Output the string to the console.
                            Console.WriteLine($"{subValue.Value}");
                        }

                        // Close the registry key.
                        registryHelper.CloseRegistryKey(hKey);
                    }
                    // Else block to handle the remaining registry paths (e.g., TypedPaths).
                    else
                    {
                        // Output the registry path being processed.
                        Console.WriteLine($"Extracting from registry path: {path}");

                        // Open the registry key.
                        IntPtr hKey = registryHelper.OpenRegistryKey(RegistryHelper.HKEY_CURRENT_USER, path);

                        // Retrieve the string values from the registry key.
                        Dictionary<string, string> subValues = registryHelper.RetrieveRegistryValues(hKey);
                        foreach (KeyValuePair<string, string> subValue in subValues)
                        {
                            string key = subValue.Key;
                            string value = subValue.Value;
                            Console.WriteLine(value); // Output the value to the console.
                        }

                        // Close the registry key.
                        registryHelper.CloseRegistryKey(hKey);
                    }
                }
                catch (Exception ex)
                {
                    // Output error message if there's an issue accessing the registry path.
                    Console.WriteLine($"Error accessing registry path {path}: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Extracts information about recently used folders, files and office documents and prints them to the console.
        /// </summary>
        private static void ExtractRecentURLFilesAndFolders()
        {
            Console.WriteLine($"\n[i] Extracting internet shortcuts");

            string userFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string desktopPath = Path.Combine(userFolder, "Desktop");
            string downloadsPath = Path.Combine(userFolder, "Downloads");
            string documentsPath = Path.Combine(userFolder, "Documents");

            List<string> directories = new List<string>
            {
                desktopPath,
                downloadsPath,
                documentsPath
            };

            List<string> urlFiles = new List<string>();

            foreach (string directory in directories)
            {
                if (Directory.Exists(directory))
                {
                    try
                    {
                        urlFiles.AddRange(Directory.GetFiles(directory, "*.url", SearchOption.AllDirectories));
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }

            Console.WriteLine($"[+] Found {urlFiles.Count} internet shortcuts");

            // Process the .url files
            foreach (string urlFile in urlFiles)
            {
                string url = ParseUrlShortcut(urlFile);
                Console.WriteLine($"{urlFile} : {url}");
            }
        }

        /// <summary>
        /// Extracts information about recently used folders, files, and office documents and prints them to the console.
        /// </summary>
        private static void ExtractRecentLNKFilesAndFolders()
        {
            // Get the path to the user's profile folder
            string userFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

            // Define paths to the Recent, Desktop, and Office Recent directories
            string recentPath = Path.Combine(userFolder, "AppData", "Roaming", "Microsoft", "Windows", "Recent");
            string desktopPath = Path.Combine(userFolder, "Desktop");
            string officeRecentPath = Path.Combine(userFolder, "AppData", "Roaming", "Microsoft", "Office", "Recent");

            // List of directories to search for .lnk files
            List<string> directories = new List<string>
            {
                recentPath,
                desktopPath,
                officeRecentPath
            };

            // List to store found .lnk files
            List<string> lnkFiles = new List<string>();

            Console.WriteLine("[i] Extracting recent file and folder shortcuts");

            // Iterate through each directory to find .lnk files
            foreach (string directory in directories)
            {
                // Check if the directory exists
                if (Directory.Exists(directory))
                {
                    // Add all .lnk files from the directory to the lnkFiles list
                    lnkFiles.AddRange(Directory.GetFiles(directory, "*.lnk", SearchOption.AllDirectories));
                }
            }

            // Print the number of found .lnk files
            Console.WriteLine($"[+] Found {lnkFiles.Count} recent shortcuts");

            // Process each found .lnk file
            foreach (string lnkFile in lnkFiles)
            {
                // Create a WinShortcut object to get the target path of the shortcut
                var shortcut = new WinShortcut(lnkFile);
                // Print the path of the .lnk file and its target path
                Console.WriteLine($"{lnkFile} : {shortcut.TargetPath}");
            }
        }

        /// <summary>
        /// Extracts and prints browser favorites from Microsoft Edge.
        /// </summary>
        /// <remarks>
        /// This function extracts the favorites from the Microsoft Edge browser by
        /// locating the Bookmarks file in the user profile directory, parsing it,
        /// and printing each favorite URL to the console.
        /// </remarks>
        private static void ExtractBrowserFavorites()
        {
            Console.WriteLine("\n[i] Extracting Browser Favorites!");
            try
            {
                // Get the path to the user's profile folder.
                string userFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

                // Extract Edge favorites from the user's profile folder.
                List<string> edgeFavorites = ExtractEdgeFavorites(userFolder);

                // Print each extracted favorite URL.
                foreach (string edgeFavorite in edgeFavorites)
                {
                    Console.WriteLine(edgeFavorite);
                }
            }
            catch (Exception ex)
            {
                // Print any exceptions that occur during the extraction process.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Extracts frequently used applications based on so called "destinations" which are present in a hashed format.
        /// </summary>
        private static void ExtractFrequentFilesFromJumpLists()
        {
            Console.WriteLine("\n[i] Analyzing recent automatic destination jumplists");

            // Get the current user's home directory
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

            // Construct paths for the destinations
            string automaticDestinationsPath = Path.Combine(userProfile, @"AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations");

            // Process automatic destinations
            if (Directory.Exists(automaticDestinationsPath))
            {
                List<string> appIDs = new List<string> { };
                string[] filePaths = Directory.GetFiles(automaticDestinationsPath);

                foreach (string filePath in filePaths)
                {
                    if (File.Exists(filePath))
                    {
                        // Generate the FileInfo object
                        FileInfo fileInfo = new FileInfo(filePath);
                        // Extract the base file name without extension
                        string fileBaseName = fileInfo.Name.Replace(fileInfo.Extension, "");
                        // Add the file hash to deducplicate it later
                        appIDs.Add(fileBaseName);
                    }
                }

                // Create a HashSet from the List to remove duplicates
                appIDs = new List<string>(new HashSet<string>(appIDs));

                List<ApplicationEntry> knownApplicationIDs = ReadAppIDsFromEmbeddedResource("KNOCKOUT.AppIdlist.csv");

                List<ApplicationEntry> resolvedAppIDs = new List<ApplicationEntry>();

                // Resolve all APP IDs
                foreach (string appID in appIDs)
                {
                    ApplicationEntry application = ResolveAppID(appID, knownApplicationIDs);
                    if (application.Name != "Unknown")
                    {
                        resolvedAppIDs.Add(application);
                    }
                }

                // Sort the list by Application name using LINQ
                List<ApplicationEntry> sortedAppIDs = resolvedAppIDs
                    .OrderBy(entry => entry.Name)
                    .ToList();

                foreach (ApplicationEntry sortedAppID in sortedAppIDs)
                {
                    Console.WriteLine($"{sortedAppID.Name} - {sortedAppID.ID}");
                }
            }
            else
            {
                Console.WriteLine($"Directory does not exist: {automaticDestinationsPath}");
            }
        }

        /// <summary>
        /// Extracts and analyzes information from the UserAssist registry key for current user's profile.
        /// Retrieves details about recently accessed applications, including their execution counts,
        /// last run timestamps, and focus usage times. Outputs results to the console.
        /// </summary>
        /// <param name="registryHelper">Helper class for interacting with the Windows Registry.</param>
        private static void ExtractUserAssist(RegistryHelper registryHelper)
        {
            Console.WriteLine($"\n[i] Extracting UserAssist registry key");

            const string userAssistKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\UserAssist";

            try
            {
                // Open the UserAssist registry key under HKEY_CURRENT_USER
                IntPtr hKey = registryHelper.OpenRegistryKey(RegistryHelper.HKEY_CURRENT_USER, userAssistKeyPath);
                List<string> subKeys = registryHelper.EnumerateSubKeys(hKey);

                // Iterate through each subkey under UserAssist
                foreach (string subKey in subKeys)
                {
                    string guid = null;
                    var run = 0;
                    try
                    {
                        // Open the subkey under UserAssist identified by 'subKey'
                        string guidSubKeyPath = $"{userAssistKeyPath}\\{subKey}";
                        IntPtr hKeyGuid = registryHelper.OpenRegistryKey(RegistryHelper.HKEY_CURRENT_USER, guidSubKeyPath);
                        List<string> subKeysGuid = registryHelper.EnumerateSubKeys(hKeyGuid);

                        // Check if the subkey contains "Count" to process its values
                        if (subKeysGuid.Contains("Count"))
                        {
                            string countSubKeyPath = $"{userAssistKeyPath}\\{subKey}\\Count";
                            IntPtr hKeyCount = registryHelper.OpenRegistryKey(RegistryHelper.HKEY_CURRENT_USER, countSubKeyPath);

                            // Retrieve raw registry values under the Count subkey
                            Dictionary<string, byte[]> subKeysCount = registryHelper.RetrieveRegistryValuesRaw(hKeyCount);

                            // Process each value under the Count subkey
                            foreach (KeyValuePair<string, byte[]> subKeyCount in subKeysCount)
                            {
                                try
                                {
                                    // Decode the registry key using ROT13 transformation
                                    string keyDecoded = UserAssist.Rot13Transform(subKeyCount.Key);

                                    // Extract GUID from the decoded key using regex
                                    guid = Regex.Match(keyDecoded, @"\b[A-F0-9]{8}(?:-[A-F0-9]{4}){3}-[A-F0-9]{12}\b", RegexOptions.IgnoreCase).Value;

                                    // Resolve folder name from GUID
                                    var foldername = Utils.GetFolderNameFromGuid(guid);

                                    // Replace GUID placeholders with folder names, handle "Unmapped" case
                                    if (!foldername.Contains("Unmapped") && guid != "")
                                    {
                                        keyDecoded = keyDecoded.Replace($"{{{guid}}}", foldername);
                                    }

                                    DateTimeOffset? lastRun = null;
                                    int? focusCount = null;
                                    TimeSpan focusTime = new TimeSpan();

                                    // Parse binary values to extract run information
                                    if (subKeyCount.Value.Length >= 16)
                                    {
                                        run = BitConverter.ToInt32(subKeyCount.Value, 4);
                                        lastRun = DateTimeOffset.FromFileTime(BitConverter.ToInt64(subKeyCount.Value, 8));

                                        // Handle Windows 7 and newer format for additional data
                                        if (subKeyCount.Value.Length >= 68)
                                        {
                                            focusCount = BitConverter.ToInt32(subKeyCount.Value, 8);
                                            focusTime = TimeSpan.FromMilliseconds(BitConverter.ToInt32(subKeyCount.Value, 12));
                                            lastRun = DateTimeOffset.FromFileTime(BitConverter.ToInt64(subKeyCount.Value, 60));
                                        }
                                    }

                                    // Check if lastRun is valid (after 1970)
                                    if (lastRun?.Year < 1970)
                                    {
                                        lastRun = null;
                                    }

                                    // Output information about the binary and its run details
                                    Console.WriteLine($"\n[+] Binary: {keyDecoded}");
                                    if (lastRun.HasValue)
                                    {
                                        Console.WriteLine($"Last Run: {lastRun.ToString()}");
                                    }
                                    if (focusCount > 0)
                                    {
                                        Console.WriteLine($"Focus Count: {focusCount}, Time: {focusTime.ToString(@"d'd, 'h'h, 'mm'm, 'ss's'")}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // Handle exceptions from processing individual values
                                    Console.WriteLine ($"[!] Error: {ex.Message}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Extracts recently connected USB Storage devices from the registry and prints them to the console.
        /// </summary>
        /// <param name="registryHelper">The helper which is used to access and query the registry.</param>
        /// <returns></returns>
        /// <exception cref="Exception">Thrown if no USB storage devices have been connected to the system before.</exception>
        private static void ExtractRecentlyConnectedUSBStorage(RegistryHelper registryHelper)
        {
            Console.WriteLine("\n[i] Extracting recently connected USB storage devices");

            try
            {
                List<string> attributesToList = new List<string>{"HardwareID", "FriendlyName"};

                string registryPath = @"SYSTEM\CurrentControlSet\Enum\USBSTOR";
                IntPtr hKey = registryHelper.OpenRegistryKey(RegistryHelper.HKEY_LOCAL_MACHINE, registryPath);
                List<string> subKeys = registryHelper.EnumerateSubKeys(hKey);

                foreach (string subKey in subKeys)
                {
                    string registryPathUSBStorage = $"{registryPath}\\{subKey}";
                    IntPtr hKeyUSBStorage = registryHelper.OpenRegistryKey(RegistryHelper.HKEY_LOCAL_MACHINE, registryPathUSBStorage);
                    List<string> subKeysUSBStorage = registryHelper.EnumerateSubKeys(hKeyUSBStorage);
                    foreach (string subKeyUSBStorage in subKeysUSBStorage)
                    {
                        string registryPathUSBStorageSerial = $"{registryPath}\\{subKey}\\{subKeyUSBStorage}";
                        IntPtr hKeyUSBStorageSerial = registryHelper.OpenRegistryKey(RegistryHelper.HKEY_LOCAL_MACHINE, registryPathUSBStorageSerial);

                        Console.WriteLine($"Serial Number: {subKeyUSBStorage.Split('&')[0]}");

                        Dictionary<string, string> subValuesUSB = registryHelper.RetrieveRegistryValues(hKeyUSBStorageSerial);
                        foreach (KeyValuePair<string, string> subValueUSB in subValuesUSB)
                        {
                            if (attributesToList.Contains(subValueUSB.Key))
                            {
                                Console.WriteLine($"{subValueUSB.Key}: {subValueUSB.Value}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Console.WriteLine("Possibly no USB Storage was connected to the system before or access to the registry is blocked.");
            }
        }

        /// <summary>
        /// Resolves application names by matching known application hashes based on the file https://raw.githubusercontent.com/kacos2000/Jumplist-Browser/master/AppIdlist.csv.
        /// </summary>
        /// <param name="appID">The appID which was gathered by walking the file system and extracting the file base name.</param>
        /// <param name="knownApplicationIDs">The key/value pair list of known applications in order to resolve known applications.</param>
        /// <returns>Returns an ApplicationEntry which contains either the resolved application name and hash or the application name "Unknown".</returns>
        private static ApplicationEntry ResolveAppID(string appID, List<ApplicationEntry> knownApplicationIDs)
        {
            ApplicationEntry entry = knownApplicationIDs.FirstOrDefault(e => e.ID == appID.ToUpper());

            if (entry == null)
            {
                return new ApplicationEntry { ID = appID, Name = "Unknown" };
            }
            else
            {
                return entry;
            }
        }
 
        /// <summary>
        /// Reads the embedded CSV file containing known application hashes for comparison to file artifacts.
        /// </summary>
        /// <param name="resourceName">The name of the embedded resource file containing the key value pairs.</param>
        /// <returns>Returns a list of Key/Value pairs containing the application name as well as the known hash.</returns>
        private static List<ApplicationEntry> ReadAppIDsFromEmbeddedResource(string resourceName)
        {
            List<ApplicationEntry> entries = new List<ApplicationEntry>();

            var assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                string headerLine = reader.ReadLine(); // Read the header line

                // Read the data lines
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    string[] fields = line.Split(',');

                    // Assuming the CSV has exactly two fields: ID and Application
                    if (fields.Length == 2)
                    {
                        entries.Add(new ApplicationEntry
                        {
                            ID = fields[0].Trim(),
                            Name = fields[1].Trim()
                        });
                    }
                }
            }

            return entries;
        }

        /// <summary>
        /// Parses MRU (Most Recently Used) data from specified registry keys and returns a list of unique, sorted filenames.
        /// </summary>
        /// <param name="mruKeys">A list of registry key paths containing MRU data.</param>
        /// <param name="registryHelper">An instance of RegistryHelper to interact with the registry.</param>
        /// <returns>A list of unique, sorted MRU filenames.</returns>
        private static List<string> ParseMruData(List<string> mruKeys, RegistryHelper registryHelper)
        {
            // List to store the MRU filenames
            List<string> mruFilenames = new List<string>();

            // Iterate through each MRU registry key path
            foreach (string mruKeyPath in mruKeys)
            {
                // Open the MRU registry key
                IntPtr mruKey = registryHelper.OpenRegistryKey(RegistryHelper.HKEY_CURRENT_USER, mruKeyPath);

                // Retrieve the MRU values from the registry key
                Dictionary<string, string> mruValues = registryHelper.RetrieveRegistryValues(mruKey);

                // Iterate through each MRU value
                foreach (KeyValuePair<string, string> mruValue in mruValues)
                {
                    string key = mruValue.Key;
                    string value = mruValue.Value;

                    // If the value contains an asterisk, split and add the filename part
                    if (value.Contains("*"))
                    {
                        mruFilenames.Add(value.Split('*')[1]);
                    }
                    else
                    {
                        // Otherwise, add the entire value
                        mruFilenames.Add(value);
                    }
                }

                // Close the MRU registry key
                registryHelper.CloseRegistryKey(mruKey);
            }

            // Sort the list of MRU filenames
            mruFilenames.Sort();

            // Remove duplicate entries to make the list unique
            mruFilenames = mruFilenames.Distinct().ToList();

            // Return the list of unique, sorted MRU filenames
            return mruFilenames;
        }

        /// <summary>
        /// Parses the URL from a given URL shortcut file (.url).
        /// </summary>
        /// <param name="filePath">The path to the URL shortcut file.</param>
        /// <returns>The URL from the shortcut file.</returns>
        /// <exception cref="FileNotFoundException">Thrown if the shortcut file does not exist.</exception>
        public static string ParseUrlShortcut(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("URL shortcut file not found", filePath);
            }

            string url = null;
            foreach (var line in File.ReadLines(filePath))
            {
                if (line.StartsWith("URL=", StringComparison.OrdinalIgnoreCase))
                {
                    url = line.Substring(4);
                    break;
                }
            }

            if (url == null)
            {
                throw new InvalidOperationException("URL not found in the shortcut file.");
            }

            return url;
        }

        /// <summary>
        /// Extracts Microsoft Edge favorites from the Bookmarks file.
        /// </summary>
        /// <param name="userProfilePath">The path to the user profile directory.</param>
        /// <returns>A list of favorite URLs.</returns>
        public static List<string> ExtractEdgeFavorites(string userProfilePath)
        {
            string bookmarksFilePath = Path.Combine(userProfilePath, "AppData", "Local", "Microsoft", "Edge", "User Data", "Default", "Bookmarks");
            if (!File.Exists(bookmarksFilePath))
            {
                throw new FileNotFoundException("Bookmarks file not found", bookmarksFilePath);
            }

            string jsonContent = File.ReadAllText(bookmarksFilePath);
            JObject bookmarks = JObject.Parse(jsonContent);

            List<string> favorites = new List<string>();
            ExtractEdgeFavoritesFromNode(bookmarks["roots"]["bookmark_bar"], favorites);
            ExtractEdgeFavoritesFromNode(bookmarks["roots"]["other"], favorites);

            return favorites;
        }

        /// <summary>
        /// Recursively extracts favorite URLs from a JSON node.
        /// </summary>
        /// <param name="node">The JSON node to extract URLs from.</param>
        /// <param name="favorites">The list to store the extracted URLs.</param>
        private static void ExtractEdgeFavoritesFromNode(JToken node, List<string> favorites)
        {
            if (node["children"] != null)
            {
                foreach (var child in node["children"])
                {
                    ExtractEdgeFavoritesFromNode(child, favorites);
                }
            }
            else if (node["type"] != null && node["type"].Value<string>() == "url")
            {
                favorites.Add(node["url"].Value<string>());
            }
        }
    }
}
