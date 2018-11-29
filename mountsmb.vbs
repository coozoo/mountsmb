'#######################################################################################
'#######
'#######	Author: Yuriy Kuzin
'#######
'#######################################################################################

dim ArgObj,argNames
Set ArgObj = WScript.Arguments

Dim objNetwork, strRemoteShare,strLocalDrive,strPer,strUsr,strPass,examplestring
strRemoteShare=""
strLocalDrive=""
strPer="FALSE"
strUsr=""
strPass=""
mount=false

argNames = "-netShare,-localDest,-isPerm,-remoteUser,-remotePass,-mount,-unmount"

examplestring=vbCrLf +"   Example:" + vbCrLf+ "      mount "+ vbCrLf+ "         "+ Wscript.ScriptName+ " -mount -localDest f: -netShare \\127.0.0.1\c$\myfolder -isPerm TRUE -remoteUser weider -remotePass dart "  + vbCrLf+ "      unmount "+ vbCrLf+ "         "+ Wscript.ScriptName+ " -unmount -localDest f:"

If WScript.Arguments.Count <> 0 Then
dim prevArgument
prevArgument=""
dim i
	For i = 0 To WScript.Arguments.Count-1
		if instr(argNames,ArgObj(i))=0 then
			if prevArgument="-netShare" then
				strRemoteShare=ArgObj(i)
			elseif prevArgument="-localDest" then
				destl = split(ArgObj(i),":")
				dest = destl(0)
				if len(dest)=1 then
					strLocalDrive=ucase(dest) + ":"
				else
					WScript.echo "Incorrect mount destination " + dest
				end if
			elseif prevArgument="-isPerm" then
				if LCase(ArgObj(i))="true" then
					strPer="TRUE"
				elseif LCase(ArgObj(i))="false" then
					strPer="FALSE"
				else
					WScript.echo "Unknown option for permanent argument. Should Be FALSE or TRUE"
				end if
			elseif prevArgument="-remoteUser" then
				strUsr=ArgObj(i)
			elseif prevArgument="-remotePass" then
				strPass=ArgObj(i)
			end if
		end if
		if ArgObj(i)="-mount" then
			mount=true
		elseif ArgObj(i)="-unmount" then
			mount=false
		end if
		prevArgument=ArgObj(i)
	Next
Else
	WScript.echo "Incorrect Arguments "+ examplestring
	WScript.quit
end if

if (strRemoteShare="" OR strLocalDrive="") AND mount=true then
	WScript.echo "Incorrect Arguments at least two args netShare localDest should be suplied" + examplestring
	WScript.quit
elseif strLocalDrive="" AND mount=false then
	WScript.echo "Incorrect Arguments localDest should be suplied" + examplestring
	WScript.quit
end if

WScript.echo ""

if mount=true then
	WScript.echo "Mounting " + strRemoteShare + " to " + strLocalDrive
	WScript.echo ""
	WScript.echo "Make Permanent " + strPer
	WScript.echo "Using remoteUser"
	WScript.echo "Using remotePass"
	MountDrive strLocalDrive,strRemoteShare,strPer,strUsr,strPass
else
	WScript.echo "Unmounting " + strLocalDrive
	UnMountDrive(strLocalDrive)
end if



'##################################################################################
Function MountDrive(strLocalDrive,strRemoteShare,strPer,strUsr,strPass)
	On Error Resume Next
	
	Dim objNetwork
	Set objNetwork = WScript.CreateObject("WScript.Network")
	dim filesys
	Set filesys = CreateObject("Scripting.FileSystemObject")
	If filesys.DriveExists(strLocalDrive)=false Then
		WScript.echo "Trying to mount: " + strLocalDrive
	else
		WScript.echo "Already mounted. Remounting! Disk" + strLocalDrive
		UnMountDrive strLocalDrive
		WScript.Sleep 3
		if filesys.DriveExists(strLocalDrive)=false then
			WScript.echo "Drive " + strLocalDrive + " unmounted"
		else
			WScript.echo "Unable to unmount drive: " + strLocalDrive + "  Try manually. Aborting."
			WScript.quit
		end if	
	end if 
	if strUsr<>"" then
		objNetwork.MapNetworkDrive strLocalDrive, strRemoteShare, strPer, strUsr, strPass
	else
		objNetwork.MapNetworkDrive strLocalDrive, strRemoteShare, strPer
	end if
	Set objNetwork = Nothing
	Set filesys	= Nothing
End Function

'##################################################################################
Function UnMountDrive(strLocalDrive)
	On Error Resume Next
	Dim objNetwork
	Set objNetwork = WScript.CreateObject("WScript.Network")
	dim filesys
	Set filesys = CreateObject("Scripting.FileSystemObject")
	Set oDrives = objNetwork.EnumNetworkDrives
	If filesys.DriveExists(strLocalDrive)=true Then
		WScript.echo "Trying to unmount: " + strLocalDrive
		objNetwork.RemoveNetworkDrive strLocalDrive
		WScript.Sleep 3
		if filesys.DriveExists(strLocalDrive)=false then
			WScript.echo "Drive " + strLocalDrive + " unmounted"
		else
			WScript.echo "Unable to unmount drive: " + strLocalDrive + "  Try manually."
		end if	
	else
		WScript.echo "Nothing to unmount: " + strLocalDrive
	end if
	
	Set objNetwork = Nothing
	Set filesys = Nothing
	Set oDrives = Nothing
End Function

