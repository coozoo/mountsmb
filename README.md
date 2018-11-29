# mountsmb
Windows VBS mounter for smb shares

# Intro
Suppose everyone has such a problem like some command line tools not able to work with remote windows shares.
So to make them work you nedd to map this share to some windows drive letter.
Actually this script made for this simple commandline interface.

# How to use
If you run it wihout parameters you will get some kind of help

`cscript mountsmb.vbs
`
Output:
`batch
cscript mountsmb.vbs
Microsoft (R) Windows Script Host Version 5.812
Copyright (C) Microsoft Corporation. All rights reserved.

Incorrect Arguments
   Example:
      mount
         mountsmb.vbs -mount -localDest f: -netShare \\127.0.0.1\c$\myfolder -isPerm TRUE -remoteUser weider -remotePass dart
      unmount
         mountsmb.vbs -unmount -localDest f:
`

Parameters:
  -mount - use to mount;
  -unmount - use to unmount;
  -localDest - local drive letter;
  -netShare - remote shared folder;
  
  Optional:
    -isPerm - make it permanent, means windows will remount it after reboot, default is FALSE;
    -remoteUser - user allowed to access this share;
    -remotePass - password for remote user;

