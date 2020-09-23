Attribute VB_Name = "Module1"
'Source code of Tape Manager v1.0
'Description :  A Backup Manager, using the backup assistant, making backup anytime you want with more option
'               than the backup assistant (You dont need anything else than the backup assistant coming with
'               windows XP/2003
'Programmer :   Sylvain Robichaud
'Date :         16/06/2005
'Copyright :    You can use this software in a company or at home without problem, you can edit this
'               code to make it better for you.  If you found a way to make that program better, and you
'               think people will like it, try to share it on planet-source-code.com like I did.  This software
'               cannot have a price, its free and will ever be free.  You cannot sell anything coming from this
'               software




'The way to make this program independant, is to find how the createfile on handle work, the tape engin use
'a handle to lock all other handle so other program cannot access the engin all at the same time.  Thats why I
'use closehandle to close the handle after I query the tape engin to know the status. I need to learn more
'about that tape engin

'Why I used Visual basic ? because this is one of the easiest programming, and I want any newbie to be able to
'edit this.




'To close the handle before making anything with ntbackup (otherwise the device will not responding)
Public Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long

'Get the tape status
Public Declare Function GetTapeStatus Lib "KERNEL32" (ByVal hDevice As Long) As Long
Public Declare Function CreateFile Lib "KERNEL32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Global Const FILE_SHARE_READ = &H1
Global Const FILE_SHARE_WRITE = &H2
Global Const OPEN_EXISTING = 3
Global Const GENERIC_READ = &H80000000
Global Const GENERIC_WRITE = &H40000000

Public blWeekSelect As Boolean
Public Function getDateWithDelimiter(strDelimiter As String)
    Dim strTemp As String
    getDateWithDelimiter = DateTime.Weekday(DateTime.Now) & strDelimiter & DateTime.Day(DateTime.Now) & strDelimiter & DateTime.Month(DateTime.Now) & strDelimiter & DateTime.Year(DateTime.Now)
End Function

Sub Main()
    mdiScreen.Show
End Sub
