# Outlook VBA Email Archiving Script

### Overview

This project is a custom VBA script designed to enhance how Microsoft Outlook handles email archiving. It provides an organized, automated way to clean up your mailbox, particularly useful for users with limited storage in free email accounts.

The script operates by archiving mail items from selected folders by mirroring the folder structure of the mail account in the archive. The retention period of each folder can be set individually. It also checks for and handles duplicate items and provides real-time feedback during the archiving process.

### Why I Built This

As someone dealing with limited storage on free email accounts, Outlook's default archiving system did not meet my needs. I wanted something more customizable and capable of handling a high volume of emails. This script solves that problem by offering control over how emails are archived and sorted, ensuring that my accounts stay organized and within their storage limits.

### Features

- **Full Control Over What Gets Archived:**  
  The script allows detailed selection of accounts to be archived, the location of the archives, and the individual folders. The retention period can be set individually for each folder. One archive file will be created for every account. Setings are carried over between archive runs.

- **Supported Outlook Items:** 
  The script will archive folders containing mail items, meeting items, or report items. Local folders are excluded. 

- **Progress Feedback & Custom Userforms:**  
  A progress bar is displayed in a userform while the script processes a mail folder, giving you visual feedback and allowing you to monitor the archiving progress.

- **Duplicate Email Detection:**  
  The script has a built-in check for duplicate emails by comparing unique attributes. Duplicate items are moved to a specified "Duplicates" folder to avoid cluttering the archive.

- **Quality-Of-Life Features:**
  On each launch of Outlook, it automatically fetches new mail for non-Exchange accounts. If rules are configured for individual accounts, they will be executed before every archive run. The script creates a backup `.bak` file of the archive each time it runs to avoid data loss.

- **Standard Folders:**
  All script data is stored in the `LocalAppData/Outlook AutoArchive/` folder. The default folder for storing archives is the Outlook standard `Documents/Outlook Files/` folder.

### Usage Instructions

1. **Import the VBA project files** into your Outlook environment.
2. Add buttons to the Outlook UI calling the two Subs in the 'UI' module.
3. Create a certificate for the script. 
2. Customize the script settings, such as archive locations, folder retention periods, or the startup behavior.
3. Run the script using the custom buttons added to your Outlook UI. 
4. Review the logs or the debug window for detailed execution feedback. Logs and configuration files can be found in the `LocalAppData/Outlook AutoArchive/` folder.
5. If Autorun is enabled, the script will periodically prompt to archive mail accounts.
   
The script is signed with a self-signed code signing certificate to avoid any security prompts from Outlook when running macros. If you're integrating it into your own Outlook, you'll need to adjust the certificate setup as needed.

### Known issues

- When unloading an archive file from Outlook, it may not free up the file consistently. This can occur when moving an archive file to another location or creating a backup of the archive before the actual archive procedure begins. In these cases, it is necessary to unload all archives and restart Outlook.
- If Outlook is closed through the script (due to the first known issue), Outlook sometimes invalidates the certificate. To avoid this, import the certificate to the `Trusted Root Certification Authorities` folder. However, this is only recommended if you know what you're doing.
- `Archive Progress Module:` If the archive is large, checking for duplicates can take a long time. A delay of 100 ms between each file operation has been included for better readability. This delay can be removed if quicker execution is desired.
- The location of the userforms on screen is not consistent, so the user experience may vary.

### Personal Motivation

This project was born out of frustration with Outlook’s built-in archive functionality. I am not a big fan of VBA either, as it is quite tedious to work with in an outdated IDE. I am far from an expert in programming, but the script I created ended up working for me, even though the code might not be as pretty as it could have been.
