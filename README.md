# outlook-move-to-thread
Microsoft Outlook VBA to move emails to the same folder as the rest of the email thread

When this macro is run within the main window of Microsoft Outlook, the user will be prompted with a pop-up containing a list of folders that emails within the thread already reside in, excluding default folders such as "Inbox" and "Sent Items". The user picks a folder from the list and emails will be moved to the selected folder.

![Example of selecting the folder](images/select_folder.png)

If there would only be one folder in the list, then the emails will be moved without prompting the user and the macro displays a message box confirming the move.

![Example of moving emails](images/moved_emails.png)

## Installation
Open Outlook VBA window using a method such as Alt+F11.

Import files by selecting "File" -> "Import File...". Import `ListThread.bas` and `ListThreadFolders.frm`. `ListThreadFolders.frx` must be in the same directory as `ListThreadFolders.frm` when it is imported.

### Signing the Macro
If you receive a warning stating "The macros in this project are disabled," you probably have security settings that only allow signed macros.

To sign the macro, see [Digitally sign your VBA macro project | Microsoft Support](https://support.microsoft.com/en-us/office/digitally-sign-your-macro-project-956e9cc8-bbf6-4365-8bfa-98505ecd1c01).

After signing and saving the project, restart Outlook. You may get a warning because the certificate is self-signed (depending on your security settings). If you trust the publisher (yourself), the notice will not appear again.

## Usage
Run `MoveToThread` in `ListThread.bas` to start the macro.

For easy access to run the macro, create a [ribbon](https://support.microsoft.com/en-us/office/customize-the-ribbon-in-office-00f24ca7-6021-48d3-9514-a31a460ecb31) or [quick access](https://support.microsoft.com/en-us/office/customize-the-quick-access-toolbar-43fff1c9-ebc4-4963-bdbd-c2b6b0739e52) shortcut. The macro will be available under "Macros" as `Project1.MoveToThread`.

## Localization
The macro ignores default folders for the following languages:
- English (US)
- Portuguese (PT)

See commit [744c4ed](https://github.com/k4j8/outlook-move-to-thread/commit/744c4ed46bf76bec88b0b302f0f358a96817cef1#diff-934591ddc75022ac0b12b82d1a44f7d572d3724183229b594a0829036ea899e3) to understand how to localize for a language and then submit a pull request. Or, create an issue for additional languages.

## Limitations
The macro crashes if certain symbols such as the percent sign (%) or backslash (\\) are within a folder name.

## Contributing
Pull requests, issues, and feature suggestions are welcome. All code in pull requests must be tested in Outlook and exported directly from the program to ensure import compatibility.

## Acknowledgements
Much of the code has been copied and edited from various sites and forums. Credit to original sources is given in the code where applicable.
