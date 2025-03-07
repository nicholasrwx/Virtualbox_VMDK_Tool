## Overview

### Issue
When attempting to attach a built-in sdcard reader to a virtual machine, like you would any other usb device, I discovered it is not recognized as a usb device anymore.

### Solution
Since the built-in readers are no longer recognized as a USB device, you have to map it to a vmdk file and access the raw disk. Although you can create stand alone vmdk files in Virtualbox, the UI itself does not have a vmdk-to-sdcard mapping option. However, the documentation does provide a way to manually do this via VBoxManage.exe. This scripted tool automates and simplifies this process by making a user friendly console application to handle this scenario.

### Purpose Of Script
1) Map a vmdk to an sd-card or other storage device.
2) Remove a vmdk, that is crashing a vm, which can't be removed easily via the UI.
3) Virtualbox UI doesn't provide a straight forward way for this type of advanced configuration.
4) Opportunity to learn and advance my powershell capabilities.

### Youtube Source Describing Problem
https://youtu.be/hXIP97sBCXg

### Related VirtualBox Documentation
[VB Manual - VirtualBox Advanced Storage Configuration](https://www.virtualbox.org/manual/topics/AdvancedTopics.html#adv-storage-config)

### Screenshots:
#### VM Menu:
![VM Menu](https://github.com/nicholasrwx/Virtualbox_VMDK_Tool/blob/main/Images/1-VM-Menu.png)
#### Device Menu:
![Device Menu](https://github.com/nicholasrwx/Virtualbox_VMDK_Tool/blob/main/Images/2-Device-Menu.png)
#### Action Menu:
![Action Menu](https://github.com/nicholasrwx/Virtualbox_VMDK_Tool/blob/main/Images/3-Action-Menu.png)

### Commands:
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
  - **old version:**
    - 🔴 **```Interacts with physical devices in the desired way despite deprecataion.```**
    - `.\VBoxManage internalcommands createrawvmdk -filename "desired-path-to-file.vdmk" -rawdisk \\.\PHYSICALDRIVEX -partitions 1`
  - **latest version:**
    - 🔴 **```Doesn't produce the desired outcome as createmedium functions differently than createrawvmdk.```**
    - `.\VBoxManage createmedium disk --filename "desired-path-to-file.vmdk" --format=VMDK --variant RawDisk --property RawDrive=\\.\PHYSICALDRIVEX --property Partitions=1`
- **Create Controller:**
  - `.\VBoxManage storagectl vm-guid --name "controller-name" --add controller-type --controller controller-bus`
- **Add Attached:**
  - `.\VBoxManage storageattach vm-guid --storagectl "SATA" --port # --device # --type hdd --medium path-to-file.vdmk -or- device-guid`

### Physical Disk Information ( Windows )
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

### Physical Drive Addresses
- This is a DeviceID used to directly access a physical disk on a windows system:
  - `\\.\PHYSICALDRIVEX`

- Indicates a local device path:
  - `\\.`
- Denotes the physical disk connected to the system. X will be replaced with the number for the desired physical disk. ( 0, 1, 2, etc ):
  - `PHYSICALDRIVEX`

### Partition Files ( -pt.vmdk / .pt.vmdk ) ###
The partition file created alongside a vmdk file stores the partition table information for the virtualdisk. It does not store any actual data. It simply holds information related to the partition layout of the virtual disk. It is not a full disk image in itself but rather a small helper file that is used to manage or identify the disk’s partitioning structure. It is never the primary disk that you attach to the VM. It is a supplementary file which typically exists alongside the .vmdk file if partition information needs to be separated. It does not typically need to be attached to a virtual machine manually as it is handled automatically by VirtualBox as part of the disk management process.
