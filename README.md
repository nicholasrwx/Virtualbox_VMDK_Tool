## Overview
Virtualbox UI does not have a vmdk-to-sdcard mapping option. However, the documentation provides a way to use VBoxManage.exe to map an sdcard ( connected via builtin reader ) to a vmdk file.
This scripted tool automates and simplifies the process by making a user friendly console application to handle this scenario.

🔴 **```NOTE: Script finalization is still in progress.```** 

### Purpose Of Script
1) Primarily it was an opportunity to really learn and expand my powershell capabilities.
2) There should be an available option for this in Virtualbox UI if VBoxManage.exe has the capability and is providing a way.

### Youtube Source Describing Problem
https://youtu.be/hXIP97sBCXg

### VirtualBox Documentation On SD Card Mapping
[VB Manual - VirtualBox Advanced Storage Configuration](https://www.virtualbox.org/manual/topics/AdvancedTopics.html#adv-storage-config)

#### VM Menu:
![VM Menu](https://github.com/nicholasrwx/Virtualbox_VMDK_Tool/blob/main/Images/1-VM-Menu.png)
#### Device Menu:
![Device Menu](https://github.com/nicholasrwx/Virtualbox_VMDK_Tool/blob/main/Images/2-Device-Menu.png)
#### Action Menu:
![Action Menu](https://github.com/nicholasrwx/Virtualbox_VMDK_Tool/blob/main/Images/3-Action-Menu.png)

## COMMANDS ##
- **Discard Saved State:**
  - `.\VBoxManage discardstate vm-guid`
- **Remove Attached:**
  - `.\VBoxManage storageattach vm-guid --storagectl controller-type --port # --device # --type hdd --medium none`
- **Close Disk:**
  - `.\VBoxManage closemedium disk path-to-file.vmdk -or- device-guid`
- **Delete By Filename:**
  - `@('path-to-file.vmdk','path-to-file-pt.vmdk') ForEach-Object { Remove-Item $_ --force }`
- **Delete By Extension:**
  - `Get-ChildItem -Path "target-dir" -Recurse -Filter "*.vmdk" | ForEach-Object { Remove-Item -Path $_.FullName -Force }`
- **Create New VMDK:**
  - **old version:** `.\VBoxManage internalcommands createrawvmdk -filename "desired-path-to-file.vdmk" -rawdisk \\.\PHYSICALDRIVEX -partitions 1`
  - **latest version:** `.\VBoxManage createmedium disk --filename "desired-path-to-file.vmdk" --format=VMDK --variant RawDisk --property RawDrive=\\.\PHYSICALDRIVEX --property Partitions=1`
- **Create Controller:**
  - `.\VBoxManage storagectl vm-guid --name "controller-name" --add controller-type --controller controller-bus`
- **Add Attached:**
  - `.\VBoxManage storageattach vm-guid --storagectl "SATA" --port # --device # --type hdd --medium path-to-file.vdmk -or- device-guid`

## Acquiring Physical Disk Information For Windows ##
- Get the entire physical disk path via cmd prompt:
  - `wmic diskdrive list brief`

- Get the disk drive number via powershell:
  - `Get-Disk`

- Get other information via diskpart utility using either cmd prompt or powershell:
  - diskpart
  - list disk
  - select disk #
  - detail disk
  - exit

## Physical Drive Addresses ##
- This is a DeviceID used to directly access a physical disk on a windows system:
  - `\\.\PHYSICALDRIVEX`

- Indicates a local device path:
  - `\\.`
- Denotes the physical disk connected to the system. X will be replaced with the number for the desired physical disk. ( 0, 1, 2, etc ):
  - `PHYSICALDRIVEX`
