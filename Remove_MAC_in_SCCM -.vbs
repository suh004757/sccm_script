'	

Option Explicit

'	=====================================
'	The following need to be supplied ...
'	=====================================

Const MPP_SCCM_SERVER = "SCCM 서버 주소 (adm 권한 필요)"

'	=========
'	Constants
'	=========

'	VB constants

Const VB_OkOnly = 0
Const VB_OkCancel = 1
Const VB_Critical = 16
Const VB_Information = 64

'	ADS Authentication constants that can be used.

Const ADS_SECURE_AUTHENTICATION = &H1
Const ADS_USE_ENCRYPTION = &H2
Const ADS_USE_SSL = &H2
Const ADS_USE_SIGNING = &H40
Const ADS_USE_SEALING = &H80
Const ADS_USE_DELEGATION = &H100
Const ADS_SERVER_BIND = &H200

'	Script must be run using cScript (or "Open with command prompt")

If (UCase(Right(wScript.FullName, 12)) <> "\CSCRIPT.EXE") Then
	MsgBox "This tool must be run using cScript or right-click and ""Open with Command Prompt""", _
               VB_OkOnly + VB_Critical, _
	       "Clean Up"
	wScript.Quit
End If

'	=======
'	Welcome
'	=======

wScript.Echo ">>>Remove_Unknown_Computers.vbs, v1.0"
If (MsgBox ("This tool is designed to remove unknown computer entries created during Windows 10 Installation " _
          & " using PXE boot ", _
            VB_OkCancel + VB_Information, _
            "Remove_Unknown_Computers.vbs, v1.0") = 2) Then
	Bye()
End If

'	=========
'	Variables
'	=========

Dim bFound
Dim objResults, objWork, objCategory
Dim strReply

'	===========
'	Credentials
'	===========
'
'	This script may be executed using the credentials of the requesting user or they
'	can supply an alternate username & password.

wScript.Echo "-->Credentials"

Dim bCredentials,strUsername, strPassword

strReply = ""
Do While ((strReply <> "Y") AND (strReply <> "N"))
	wScript.Stdout.Write "   Do you want to run this script as yourself (y/n) "
	strReply = UCase(Left(wScript.Stdin.ReadLine(), 1))
	If (strReply = "") Then
		MsgBox "This script will attempt to delete unknown computer objects from SCCM, in order " _
		     & "to do this it must be run using credentials authorised to perform such actions. " _
		     & "If you have this authority enter y otherwise enter n and you will be asked for " _
		     & "a username & password to be used.", _
		       VB_OkOnly + VB_Information, _
		       "Clean Up"
	End If
Loop
bCredentials = (strReply = "N")
If (bCredentials) Then
	wScript.StdOut.Write "      Username: "
	strUsername = wScript.StdIn.ReadLine()
	Set objPassword = CreateObject("ScriptPW.Password") 
	wScript.StdOut.Write "      Password: "
	strPassword = objPassword.GetPassword()
'	strPassword = wScript.StdIn.ReadLine()
End If


'	===================================
'	System Center Configuration Manager
'	===================================
'
'	This script may be deleting computer from SCCM, this section establishes the
'	links to SCCM.

wScript.Echo "-->Connecting System Center Configuration Manager"

Dim bSccm
Dim objWmi, objServer, objSccm
Dim strSccmServer, strSccmSite

'	Connect to SCCM?
'	----------------

strReply = "Y"
'Do While ((strReply <> "Y") And (strReply <> "N"))
'	wScript.Stdout.Write "   Connect to SCCM? (y/n) "
'	strReply = UCase(Left(wScript.Stdin.ReadLine(), 1))
'	If (strReply = "") Then
'		MsgBox "Do you want to search for and delete computer objects from SCCM? ", _
'		       VB_OkOnly + VB_Information, _
'		       "Clean Up"
'	End If
'Loop
bSccm = (strReply = "Y")
If (bSccm) Then

	'	SCCM server
	'	-----------

	If (Not IsEmpty(MPP_SCCM_SERVER)) Then
		strSccmServer = MPP_SCCM_SERVER
		wScript.Echo "   SCCM Server: " & strSccmServer
	Else
		Do While (IsEmpty(strSccmServer))
			wScript.Stdout.Write "   SCCM Server: "
			strSccmServer = wScript.Stdin.ReadLine()
			If (strSccmServer = "") Then
				MsgBox "Please supply the name of the server where SCCM is installed.", _
				       VB_OkOnly + VB_Information, _
			               "Clean Up"
				strSccmServer = Empty
			End If
		Loop
	End If

	'	Using WMI, get the SCCM site code
	'	---------------------------------

	wScript.Echo "   Get SiteCode ..."

	Set objWmi = CreateObject("WbemScripting.SWbemLocator")

	On Error Resume Next
	If (Not bCredentials) Then
		Set objServer = objWmi.ConnectServer(strSccmServer, "root\sms")
	Else
		Set objServer = objWmi.ConnectServer(strSccmServer, "root\sms", strUsername, strPassword)
		objServer.Security_.ImpersonationLevel = 3
	End If
	If (Err.Number <> 0) Then
		wScript.Echo "***   Unable to get site code: " & Err.Description
		Bye
	End If
	On Error GoTo 0

	Set objResults = objServer.ExecQuery ("SELECT *" _
	                                    & "  FROM SMS_ProviderLocation" _
	                                    & "  WHERE Machine = '" & strSccmServer & "'" _
	                                    & "    AND ProviderForLocalSite = True")
	For each objWork in objResults 
		If (Not bCredentials) Then
			Set objSCCM = objWmi.ConnectServer(strSccmServer, _
			                                   "root\sms\site_" & objWork.SiteCode)
		Else
			Set objSCCM = objWmi.ConnectServer(strSccmServer, _
			                                   "root\sms\site_" & objWork.SiteCode, _
			                                   strUsername, strPassword)
			objSCCM.Security_.ImpersonationLevel = 3
		End If
		strSccmSite = objWork.SiteCode
		Exit For
	Next
	If (IsEmpty(strSccmSite)) Then
		wScript.Echo "***   Unable to get SiteCode"
		Bye
	End If

	wScript.Echo "      Found, SiteCode = " & strSccmSite
	wScript.Echo "-->Connected System Center Configuration Manager"
End If

'	======================
'	For each target PC ...
'	======================

'	Just loop ..

Do
	TargetLoop()
Loop

'	Executing this subroutine ...

Sub TargetLoop()
	Dim strTarget, strSplit
	Dim objData, objAdTarget, objAdParent, objSccmTarget
	Dim iSccm
	Dim bAdFound, bSccmFound

	'	Get MAC
	'	--------

	Do While (IsEmpty(strTarget))
		wScript.Stdout.Write "--> Provide MAC in format F4:30:B9:D4:E0:2B"
		wScript.Stdout.Write "-->Target MAC: "
		strTarget = wScript.StdIn.ReadLine()
		If (strTarget = "") Then
			MsgBox "Please supply the MAC of the computer to be deleted, or / to end.", _
			       VB_OkOnly + VB_Information, _
		               "Clean Up"
			strTarget = Empty
		End If
		If (strTarget = "/") Then
			Bye
		End If
	Loop



	'	Locate it in SCCM database?
	'	---------------------------
	'	There may be several of these, so we will count them at this stage, if we choose to delete
	'	then it will need to search again.

	bSccmFound = False
	If (bSccm) Then
		wScript.Echo "   Searching SCCM for " & strTarget & " ..."

		Set objResults = objSccm.ExecQuery ("SELECT ResourceID,Name" _
                		                  & "  FROM SMS_R_System" _
                                		  & "  WHERE MACAddresses Like '%" & strTarget & "%'")
		iSccm = 0
		For Each objWork in objResults
			iSccm = iSccm + 1
			wScript.Echo iSccm & " .Computer Name : " & objWork.Name
		Next
		bSccmFound = (iSccm <> 0)
		If (Not bSccmFound) Then
			wScript.Echo "**    " & strTarget & " not found"
		Else
			wScript.Echo "*     " & iSccm & " entries found"
		End If
	End If

	'	Anything found to delete?
	'	-------------------------

	If (Not (bAdFound Or bSccmFound)) Then
		Exit Sub
	End If

	'	Delete?
	'	-------

	strReply = ""
	Do While ((strReply <> "Y") And (strReply <> "N"))
		wScript.StdOut.Write "   Delete " & strTarget & "? (y/n)"
		strReply = UCase(Left(wScript.StdIn.ReadLine(), 1))
	Loop
	If (strReply = "N") Then
		wScript.Echo "*  Cancelled"
		Exit Sub
	End If



	'	Delete from SCCM
	'	----------------

	If (bSccmFound) Then
		wScript.Echo "   Deleting " & strTarget & " from SCCM ..."
		Set objResults = objSccm.ExecQuery ("SELECT ResourceID" _
                		                  & "  FROM SMS_R_System" _
                                		  & "  WHERE MACAddresses Like '%" & strTarget & "%'")
		For Each objWork in objResults
			Set objSccmTarget = objSCCM.Get ("SMS_R_System.ResourceID=" & objWork.ResourceID)

			On Error Resume Next
			objSccmTarget.Delete_
			If (Err.Number = 0) Then
				wScript.Echo "      Done"
			Else
				wScript.Echo "***   " & Err.Description
			End If
			On Error GoTo 0
		Next
	End If
End Sub

'	Bye
'	===

Sub Bye()
	wScript.Echo "<<<Bye"
'	wScript.Stdin.ReadLine()
	wScript.Quit
End Sub

