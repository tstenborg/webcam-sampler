Attribute VB_Name = "WebcamSampler"

Option Explicit

Type Camera
  fileType As String
  label As String
  URL As String
  refreshRate As Long    'Refresh rate adopted (seconds).
  lastRefresh As Date
End Type

Dim GstrFolderName As String
Dim GlngFilesSuccessful As Long
Dim GlngFilesTried As Long
Dim GdtNow As Date

Declare PtrSafe Function BeepAPI Lib "kernel32" Alias "Beep" (ByVal Frequency As Long, ByVal Milliseconds As Long) As Long
Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Sub DownloadFileAPI()

    'N.B. Have computer sleep settings been disabled?
    'This program runs for an extended period.

    Const FILE_PREFIX = "C:\Webcam Screenshots\"
    Const MAX_INDEX = 3

    Dim CameraArray(MAX_INDEX) As Camera
    Dim strMessage As String
    Dim dtEnd As Date
    Dim dtStart As Date
    Dim x As Integer

    On Error GoTo ErrorHandler

    'Air Services Webcams, with five minute (300 second) sampling rates.
    CameraArray(0).label = "NORFOLK_N"
    CameraArray(0).URL = "https://weathercams.airservicesaustralia.com/wp-content/uploads/airports/200288/200288_360.jpg"
    CameraArray(0).refreshRate = 300
    '
    CameraArray(1).label = "NORFOLK_E"
    CameraArray(1).URL = "https://weathercams.airservicesaustralia.com/wp-content/uploads/airports/200288/200288_090.jpg"
    CameraArray(1).refreshRate = 300
    '
    CameraArray(2).label = "NORFOLK_S"
    CameraArray(2).URL = "https://weathercams.airservicesaustralia.com/wp-content/uploads/airports/200288/200288_180.jpg"
    CameraArray(2).refreshRate = 300
    '
    CameraArray(3).label = "NORFOLK_W"
    CameraArray(3).URL = "https://weathercams.airservicesaustralia.com/wp-content/uploads/airports/200288/200288_270.jpg"
    CameraArray(3).refreshRate = 300

    GstrFolderName = ""
    GlngFilesTried = 0
    GlngFilesSuccessful = 0

    'Parse camera image file types.
    For x = 0 To MAX_INDEX
        CameraArray(x).fileType = Right$(CameraArray(x).URL, 4)
    Next x

    'Syntax: DateSerial(year, month, day), TimeSerial(hour, minute, second)
    dtStart = DateSerial(2025, 10, 3) + TimeSerial(18, 45, 0)
    dtEnd = DateSerial(2025, 10, 3) + TimeSerial(19, 15, 0)
    strMessage = "Start time: " & CStr(dtStart) & vbNewLine & "End time: " & CStr(dtEnd) & vbNewLine & "Continue?"
    If MsgBox(strMessage, vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Run Time") <> vbYes Then
        strMessage = "Operation cancelled."
        Application.StatusBar = strMessage
        Exit Sub
    End If

    'Don't begin until the start time is reached.
    GdtNow = Now
    Do While GdtNow < dtStart
        'The shortest refresh rate is 15 sec, so set the wait between cycles = 15 sec.
        Application.Wait (Now + TimeSerial(0, 0, 15))
        GdtNow = Now
    Loop

    'Ensure the top-level folder exists.
    Call CreateFolder(FILE_PREFIX)
    'Create a timestamped subfolder to store images.
    GdtNow = Now
    GstrFolderName = FILE_PREFIX & "x" & Format(GdtNow, "ddd-hh") & Format(Floor(CDbl(Minute(GdtNow)), 5#), "00")
    Call CreateFolder(GstrFolderName)
    GstrFolderName = GstrFolderName & "\"

    'Create timestamped files.
    Do While Now < dtEnd

        'Ensure a valid Internet connection exists.
        Do While Not IsInternetConnected()
            'Frequency (Hertz) = 800, duration (milliseconds) = 500.
            BeepAPI 800, 500
        Loop

        For x = 0 To MAX_INDEX
            If DateAdd("s", CameraArray(x).refreshRate, CameraArray(x).lastRefresh) <= Now Then
                Call DownloadImage(CameraArray(x))
                CameraArray(x).lastRefresh = Now
            End If
        Next x

        'Alert if any files were missed.
        If GlngFilesTried <> GlngFilesSuccessful Then
            BeepAPI 800, 500
        End If

        'The shortest refresh rate is 15 sec, so set the wait between cycles = 15 sec.
        Application.Wait (Now + TimeSerial(0, 0, 15))
    Loop

    If GlngFilesTried = GlngFilesSuccessful Then
       strMessage = "All files downloaded."
    Else
       strMessage = "Error. Files tried = " & CStr(GlngFilesTried) & "." & vbCrLf & "Files ok = " & CStr(GlngFilesSuccessful) & "."
    End If

    Application.StatusBar = strMessage
    MsgBox strMessage, vbOKOnly, "Take Shots"

    Exit Sub

ErrorHandler:
    MsgBox Err.Number & ": " & Error.Description

End Sub

Function CreateFolder(strPath As String) As Boolean

   'To run this function, ensure that from VBA's Tools -> References, a check mark appears beside Microsoft Scripting Run-time.
   'N.B. FileSystemObject's CreateFolder method can only create one level of new folder at a time.

   Dim fso As New Scripting.FileSystemObject

   On Error GoTo ErrorHandler

   CreateFolder = False

   'Create the target folder if it doesn't exist.
   If Not fso.FolderExists(strPath) Then
       fso.CreateFolder strPath
   End If

   CreateFolder = True

   Exit Function

ErrorHandler:
    MsgBox Err.Number & ": " & Error.Description

End Function

Function Floor(dblX As Double, dblFactor As Double) As Long

    'dblX is the value you want to round.
    'Factor is the multiple to which you want to round.

    On Error GoTo ErrorHandler

    Floor = CLng(Int(dblX / dblFactor) * dblFactor)

    Exit Function

ErrorHandler:
    MsgBox Err.Number & ": " & Error.Description

End Function

Function FolderExists(strPath As String) As Boolean

   Dim fso As New FileSystemObject

   On Error GoTo ErrorHandler

   FolderExists = False
   If fso.FolderExists(strPath) Then
      FolderExists = True
   End If

   Exit Function

ErrorHandler:
    MsgBox Err.Number & ": " & Error.Description

End Function

Function IsInternetConnected() As Boolean

    Dim lngR As Long

    On Error GoTo ErrorHandler

    lngR = InternetGetConnectedState(0, 0&)
    If lngR = 0 Then
        IsInternetConnected = False
    Else
        If lngR <= 4 Then
            IsInternetConnected = True
        Else
            IsInternetConnected = False
        End If
    End If

    Exit Function

ErrorHandler:
    MsgBox Err.Number & ": " & Error.Description

End Function

Sub DownloadImage(tmpCamera As Camera)

   Dim strLocalFilePath As String

   On Error GoTo ErrorHandler

   GlngFilesTried = GlngFilesTried + 1

   strLocalFilePath = GstrFolderName & tmpCamera.label & Format(Now, " yyyy-mmm-dd hh\h mm\m ss\s") & tmpCamera.fileType
   If URLDownloadToFile(0, tmpCamera.URL, strLocalFilePath, 0, 0) = 0 Then
      Call DeleteUrlCacheEntry(tmpCamera.URL)
      GlngFilesSuccessful = GlngFilesSuccessful + 1
   End If

   Exit Sub

ErrorHandler:
    MsgBox Err.Number & ": " & Error.Description

End Sub

