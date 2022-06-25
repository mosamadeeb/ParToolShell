# ParToolShell
Windows shell extension that adds ParTool commands to File Explorer's context menu. Made in .NET using [SharpShell](https://github.com/dwmkerr/sharpshell).

***

# Installing
Download the MSI setup file from the [latest release](https://github.com/SutandoTsukai181/ParToolShell/releases/latest) and run it. Finish the setup and the program will be installed.

***

# Usage

Check the [Usage guide](https://github.com/SutandoTsukai181/ParToolShell/wiki/Usage) in the wiki.

***

# Uninstalling

You can uninstall ParShellTool from the **Apps & features** section in Windows settings, or from the **Uninstall a program** menu in Control Panel.

However, if you see a prompt asking for Windows Explorer to be closed, make sure to select **"Do not close applications"**. After the uninstaller is done, reboot your PC.

<img src="https://user-images.githubusercontent.com/52977072/175780203-293051a3-4ccf-46e3-aa3c-11410c7ea024.png" alt="image" width="350"/>

***

# Logging

ParToolShell logs the arguments that are used to run ParTool, as well as stack traces whenever an exception occurs. To view the logs, you need to follow [this guide](https://github.com/dwmkerr/sharpshell/blob/master/docs/logging/logging.md).

***

# Building
Make sure to update submodules.

Use `ParToolShell.sln` to open the solution. Build the `ParShellExtension` project, then the `ParShellSetup` project.

***

# Credits

dwmkerr for [SharpShell](https://github.com/dwmkerr/sharpshell).

Kaplas80 for [ParManager](https://github.com/Kaplas80/ParManager).

Kent for the menu icon.
