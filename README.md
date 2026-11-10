# What The Lock

What The Lock is a tool that integrates into Windows shell components (i.e. Explorer and file dialogs) to release currently locked (opened by another process) files. It checks for locks when shell components try to delete something, a context menu item is also provided.

<p align="center">
  <img src="./.github/TaskDialog.png" alt="Task dialog for file unlocking">
</p>
<p align="center">
  <img src="./.github/ContextMenu.png" alt="Context menu item">
</p>

### Features
- ...
- Built using Lazarus :fire:

### Installation
Download the installer from "Releases" and run it. The installer asks whether What The Lock should be installed
- for all users (requires administrator rights): installs to %PROGRAMFILES%\What The Lock
- only for the current user (no administrator rights required): installs to %LOCALAPPDATA%\Programs\What The Lock

When installed for the current user only, What The Lock is not available in elevated processes, since Windows ignores per-user COM registrations there.

Complete uninstallation is possible using "Programs and Features".

### Building
- Install Lazarus IDE

To use the build script "Build.bat" you need to configure variables at the top of the file.

### Thanks
- [Lazarus IDE](https://www.lazarus-ide.org)
- [Free Pascal](https://www.freepascal.org)
- [Font Awesome](https://fontawesome.com) for icons
- [MahdiSafsafi](https://github.com/MahdiSafsafi) for DDetours
