using System;
// for directory info
using System.IO;
// for process.start
using System.Diagnostics;
// for creating a .lnk file
using IWshRuntimeLibrary;


namespace folderCreate
{
    class Program
    {
        // better ui for exiting
        static void ExitPrompt()
        {
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }

        // function to enter an insecure password, unencrypted, securely
        static void pwPrompt(string secretPW)
        {
            // initialize password string for while loop
            string pw = "";
            // while the password is incorrect loop through request, or type exit to quit
            while (!pw.Equals(secretPW))
            {
                // reset string after iteration
                pw = "";
                Console.WriteLine("Enter Password or type [Exit] to exit:");
                // for reading in keystrokes
                ConsoleKeyInfo key;
                // obstruct and read in keystrokes
                do
                {
                    key = Console.ReadKey(true);
                    // as long as a Backspace or Enter wasn't used
                    if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                    {
                        // append keystroke to pw string
                        pw += key.KeyChar;
                        Console.Write("*");
                    }
                    else
                    {
                        // If backspace is used, remove a space from the pw and correct the number of * displayed
                        if (key.Key == ConsoleKey.Backspace && pw.Length > 0)
                        {
                            // reduce the characters in pw by 1
                            pw = pw.Substring(0, (pw.Length - 1));
                            // backspace, erase by using a space, and backspace again
                            Console.Write("\b \b");
                        }
                    }
                // exit on Enter
                } while (key.Key != ConsoleKey.Enter);
                Console.WriteLine("");
                // Console.WriteLine(pw); *** Line used for debugging

                // on exit sequence quit application
                if (pw.Equals("Exit", StringComparison.InvariantCultureIgnoreCase))
                {
                    Environment.Exit(0);
                }
            }
        }

        static void CopyFolders(string srcDir, string destDir)
        {
            // check if src directory exists
            if (!Directory.Exists(srcDir))
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: " + srcDir
                    );
            }

            // Create the parent directory
            Directory.CreateDirectory(destDir);

            // Now create all subdirectories
            foreach (string subDir in Directory.EnumerateDirectories(srcDir))
            {
                // EnumerateDirectories returns a full file path, split out the final folder from the string and add to new target path
                string subFolder = subDir.Substring(subDir.LastIndexOf("\\"));
                Directory.CreateDirectory($@"{destDir}\{subFolder}");
                // also need to check for folders or files in the subdirectories
                foreach (string subChild in Directory.EnumerateDirectories(subDir))
                {
                    string childFolder = subChild.Substring(subChild.LastIndexOf("\\"));
                    Directory.CreateDirectory($@"{destDir}\{subFolder}\{childFolder}");

                    // also move any files contained in subChild to childFolder
                    foreach (string childDir in Directory.EnumerateFiles(subChild))
                    {
                        string childFileName = childDir.Substring(childDir.LastIndexOf("\\"));
                        System.IO.File.Copy(childDir, $"{destDir}{subFolder}{childFolder}{childFileName}", true);
                    }
                }
                // Copy any files in the subfolders over
                foreach (string subFile in Directory.EnumerateFiles(subDir))
                {
                    string subFileName = subFile.Substring(subDir.LastIndexOf("\\"));
                    System.IO.File.Copy(subFile, $"{destDir}{subFileName}", true);
                }
            }
        }

        static void ShortcutBuilder(string lnkLoc, string physDir)
        {
            // startup path = physDir
            var shell = new WshShell();
            // create shortcut in the physDir with the name of the last folder in lnkLoc plus the file extension .lnk;
            var shortcutObject = (IWshShortcut)shell.CreateShortcut(physDir + lnkLoc.Substring(lnkLoc.LastIndexOf("\\")) + ".lnk");
            // direct the link to the destination folder
            shortcutObject.TargetPath = lnkLoc;
            shortcutObject.Save();
        }

        static void Main()
        {
            pwPrompt("H0Ld3nItD0wn!");

            // prompt for the case name to use
            Console.WriteLine("Please enter the Case Name, commas will be ignored:");
            string caseName = Console.ReadLine();
            // clean up caseName to remove commas, they break the acrobat JS save function for appending emails
            while(caseName.IndexOf(",") > -1) {
                caseName = caseName.Remove(caseName.IndexOf(","), 1);
            }
            // prompt for the HL number assigned
            Console.WriteLine("Please enter the HL Number:");
            string hlNum = Console.ReadLine();
            // combined for ease of use
            string newName = caseName + " " + hlNum;

            // confirm that what was entered is what they want to use
            Console.WriteLine($"You've entered: \"{newName}\" is this correct (Y/N)?");
            string confirm = Console.ReadLine();
            Console.WriteLine();
            // if accurate
            if (confirm.Equals("Y", StringComparison.InvariantCultureIgnoreCase))
            {
                // Create folder source string @"\\HOLDEN-FS01\Common\Cases\_NewCase Template"
                // testing string @"D:\Cases\_NewCase Template"
                string srcDir = @"\\HOLDEN-FS01\Common\Cases\_NewCase Template";
                // Create folder destination string @"\\HOLDEN-FS01\Common\Cases\" + newName + @"\"
                // testing string @"D:\Cases\" + newName + @"\"
                string destDir = $@"\\HOLDEN-FS01\Common\Cases\{newName}\";

                // Call copyFolders to generate directories
                CopyFolders(srcDir, destDir);
                Console.WriteLine($@"Folders for G:\Cases\{newName} have been created");

                // Create Folder \\HOLDENDATA\CaseMap\{caseName} and create shortcut for it in \\HOLDEN-FS01\Common\Cases\{newName}
                Directory.CreateDirectory($@"\\HOLDENDATA\CaseMap\{caseName}");
                Console.WriteLine($@"The K:\{caseName} folder has been created");

                // create shortcut to "\\HOLDEN-FS01\Common\Cases\{newName}\5 Medical Records (Original)\Medical Billing Affidavits" in...
                // \\HOLDEN-FS01\Common\Cases\{newName}\3 Pleadings
                ShortcutBuilder(Path.Combine(destDir, @"5 Medical Records (Original)\Medical Billing Affidavits"), Path.Combine(destDir, @"3 Pleadings\"));

                // create shortcut to \\HOLDENDATA\CaseMap\{caseName} folder inside \\HOLDEN-FS01\Common\Cases\{newName}\
                ShortcutBuilder($@"\\HOLDENDATA\CaseMap\{caseName}", destDir);
                Console.WriteLine("All shortcuts have been established");

                // Create icacls process command strings
                string icaclsModify = $"\"{destDir}*\" /grant:r \"HOLDENMCKENNA\\Limited Case Access\":(OI)(CI)(IO)M";
                string icaclsWrite = $"\"{destDir}*\" /grant:r \"HOLDENMCKENNA\\Limited Case Access\":(NP)W";
                // Console.WriteLine(icaclsModify); *** Line used for debugging
                // start icacls process and send parameters to run as
                var modfy = Process.Start("icacls.exe", icaclsModify);
                // wait for icacls to finish and confirm it ran successfully
                modfy.WaitForExit();
                if (modfy.ExitCode != 0)
                {
                    Console.WriteLine($"Unable to set modify permissions, Exit Code: {modfy.ExitCode}");
                    ExitPrompt();
                    Environment.Exit(-1);
                }
                // start icacls process and send parameters to run as
                var wrte = Process.Start("icacls.exe", icaclsWrite);
                // wait for icacls to finish and confirm it ran successfully
                wrte.WaitForExit();
                if (wrte.ExitCode != 0)
                {
                    Console.WriteLine($"Unable to set write permissions, Exit Code: {modfy.ExitCode}");
                    ExitPrompt();
                    Environment.Exit(-2);
                }

                // display confirmation that permissions have been set correctly
                Console.WriteLine($@"Permissions have been set for: G:\Cases\{newName}");
                Console.WriteLine();
            }
            else
            {
                Console.WriteLine("Now exiting, please try again.");
            }
            Program.ExitPrompt();
        }
    }
}