' SPDX-License-Identifier: BSD-3-Clause
' Copyright (c) 2019, 2026 Travis Stenborg

Attribute VB_Name = "modWebcamSampler"

Option Explicit

' Custom type.
Public Type Camera
    FileType As String    ' 8 byte pointer.
    Label As String       ' 8 byte pointer.
    URL As String         ' 8 byte pointer.
    LastSample As Date    ' 8 bytes.
    SampleRate As Long    ' 4 bytes. Sample rate adopted (seconds).
End Type

' Windows API function declarations.

' VBA pre-processor directives are prefixed by a hash.
' Declarations incompatible with the current VBA version may appear in red.
' The code will however, still compile.
#If VBA7 Then
    ' Declarations for Microsoft Office 2010 and later.
    Public Declare PtrSafe Function BeepAPI Lib "kernel32" Alias "Beep" (ByVal Frequency As Long, ByVal Milliseconds As Long) As Long
    Public Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
    Public Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As LongPtr) As Long
    Public Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
        ByVal pCaller As LongPtr, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As LongPtr, _
        ByVal lpfnCB As LongPtr) As Long
#Else
    ' Declarations for Microsoft Office 2007 and earlier.
    Public Declare Function BeepAPI Lib "kernel32" Alias "Beep" (ByVal Frequency As Long, ByVal Milliseconds As Long) As Long
    Public Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
    Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
    Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
        ByVal pCaller As Long, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As Long, _
        ByVal lpfnCB As Long) As Long
#End If

' User notification icon guidelines.
' a) vbCritical:    used for bugs.
' b) vbExclamation: used for user errors, like invalid data entry.
' c) vbOKOnly:      used when providing information.
' d) vbQuestion:    used when a user needs to make a choice.

Sub DownloadFileAPI()

    ' N.B. Have computer sleep settings been disabled?
    ' This program runs for an extended period.

    Const PARENT_DIRECTORY As String = "C:\webcam-samples\"
    ' The Windows traditional file name length limit is 260 characters.
    Const WIN_PATH_LEN_LIMIT As Long = 260

    Dim CameraArray() As Camera
    Dim strEnd As String
    Dim strFolderName As String
    Dim strLabel As String
    Dim strMessage As String
    Dim strStart As String
    Dim dtEnd As Date
    Dim dtNow As Date
    Dim dtStart As Date
    Dim lngFilesSuccessful As Long
    Dim lngFilesTried As Long
    Dim lngMaxIndex As Long
    Dim lngMaxLabelLength As Long
    Dim lngNumWebcams As Long
    Dim lngX As Long

    On Error GoTo ErrorHandler

       ' Ensure details for at least one webcam are present.
       ' Miscellaneous configuration, remembering arrays are zero-indexed here.
1      lngNumWebcams = 4
2      lngMaxIndex = lngNumWebcams - 1
3      ReDim CameraArray(lngMaxIndex)

       ' Air Services Webcams, with five minute (300 second) sampling rates.
4      CameraArray(0).Label = "NORFOLK_N"
5      CameraArray(0).URL = "https://weathercams.airservicesaustralia.com/wp-content/uploads/airports/200288/200288_360.jpg"
6      CameraArray(0).SampleRate = 300
       '
7      CameraArray(1).Label = "NORFOLK_E"
8      CameraArray(1).URL = "https://weathercams.airservicesaustralia.com/wp-content/uploads/airports/200288/200288_090.jpg"
9      CameraArray(1).SampleRate = 300
       '
10     CameraArray(2).Label = "NORFOLK_S"
11     CameraArray(2).URL = "https://weathercams.airservicesaustralia.com/wp-content/uploads/airports/200288/200288_180.jpg"
12     CameraArray(2).SampleRate = 300
       '
13     CameraArray(3).Label = "NORFOLK_W"
14     CameraArray(3).URL = "https://weathercams.airservicesaustralia.com/wp-content/uploads/airports/200288/200288_270.jpg"
15     CameraArray(3).SampleRate = 300

       ' Ensure the top-level folder exists.
16     CreateFolder PARENT_DIRECTORY

       ' Determine maximum allowed length of webcam labels.
       ' a) Create a subfolder name for testing.
       ' b) Assume a local file path format: strFolderName & tmpCamera.Label & Format$(Now, "-yyyy-mmm-dd-hh\h-mm\m-ss\s") & tmpCamera.FileType
       ' c) Allow for file type extensions up to five characters, e.g., ".webp".
17     strFolderName = CreateTimestampedFolderName(PARENT_DIRECTORY)
18     lngMaxLabelLength = WIN_PATH_LEN_LIMIT - Len(strFolderName & Format$(Now, "-yyyy-mmm-dd-hh\h-mm\m-ss\s") & ".webp")

       ' Input validation and collation.
19     For lngX = 0 To lngMaxIndex

            ' Ensure a URL is present.
20          If Len(CameraArray(lngX).URL) = 0 Then
21              strMessage = "Missing URL (webcam " & (lngX + 1) & ")." & vbNewLine & "Operation cancelled."
22              DisplayValidationMessage "Take Samples", strMessage
23              Application.StatusBar = False    ' Return Status Bar control to Excel.
24              Exit Sub
25          End If

            ' Ensure the URL is a string.
26          If VarType(CameraArray(lngX).URL) <> vbString Then
27              strMessage = "Invalid non-text data (webcam " & (lngX + 1) & ")." & vbNewLine & "Operation cancelled."
28              DisplayValidationMessage "Take Samples", strMessage
29              Application.StatusBar = False    ' Return Status Bar control to Excel.
30              Exit Sub
31          End If

            ' Ensure a webcam label is present.
32          If Len(CameraArray(lngX).Label) = 0 Then
33              strMessage = "Missing webcam label (webcam " & (lngX + 1) & ")." & vbNewLine & "Operation cancelled."
34              DisplayValidationMessage "Take Samples", strMessage
35              Application.StatusBar = False    ' Return Status Bar control to Excel.
36              Exit Sub
37          End If

            ' Ensure the webcam label is a string.
38          If VarType(CameraArray(lngX).Label) <> vbString Then
39              strMessage = "Invalid non-text data (webcam " & (lngX + 1) & ")." & vbNewLine & "Operation cancelled."
40              DisplayValidationMessage "Take Samples", strMessage
41              Application.StatusBar = False    ' Return Status Bar control to Excel.
42              Exit Sub
43          End If

            ' Ensure the webcam label follows Windows file name rules.
            ' N.B. HasBadCharacters displays a relevant validation message.
44          strLabel = CameraArray(lngX).Label
45          If HasBadCharacters(strLabel) Then
46              Application.StatusBar = False    ' Return Status Bar control to Excel.
47              Exit Sub
48          End If

            ' The webcam label will get a folder prefix and date-time stamp.
            ' Ensure the total length doesn't exceed file path limits.
49          If Len(strLabel) > lngMaxLabelLength Then
                ' Display "<", not "<=", to users, for readability.
50              strMessage = "Webcam labels must be < " & (lngMaxLabelLength + 1) & " characters long." & vbNewLine & "Operation cancelled."
51              DisplayValidationMessage "Take Samples", strMessage
52              Application.StatusBar = False    ' Return Status Bar control to Excel.
53              Exit Sub
54          End If

            ' Ensure a sample rate is present.
55          If Len(CameraArray(lngX).SampleRate) = 0 Then
56              strMessage = "Missing sample rate (webcam " & (lngX + 1) & ")." & vbNewLine & "Operation cancelled."
57              DisplayValidationMessage "Take Samples", strMessage
58              Application.StatusBar = False    ' Return Status Bar control to Excel.
59              Exit Sub
60          End If

           ' Parse camera image file types.
61         CameraArray(lngX).FileType = Right$(CameraArray(lngX).URL, 4)
62     Next lngX

63     lngFilesTried = 0
64     lngFilesSuccessful = 0

       ' Process start and end times.
       ' Syntax: DateSerial(year, month, day), TimeSerial(hour, minute, second)
65     dtStart = DateSerial(2026, 10, 1) + TimeSerial(22, 36, 0)
66     dtEnd = DateSerial(2026, 10, 1) + TimeSerial(22, 36, 30)
67     strStart = Format$(dtStart, "d-mmm-yyyy hh:nn:ss")
68     strEnd = Format$(dtEnd, "d-mmm-yyyy hh:nn:ss")

69     strMessage = "Start time: " & strStart & vbNewLine & "End time: " & strEnd & vbNewLine & "Continue?"
70     If MsgBox(strMessage, vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Run Time") <> vbYes Then
71         strMessage = "Operation cancelled."
72         Application.StatusBar = strMessage
73         Exit Sub
74     End If

       ' Don't begin until the start time is reached.
75     dtNow = Now
76     Do While dtNow < dtStart
           ' The shortest refresh rate is 15 sec, so set the wait between cycles = 15 sec.
77         Application.Wait (Now + TimeSerial(0, 0, 15))
78         dtNow = Now
79     Loop
80     Application.StatusBar = "Taking samples. Webcam sampling will end at " & strEnd & "."

       ' Create a timestamped subfolder to store images.
81     dtNow = Now
82     strFolderName = CreateTimestampedFolderName(PARENT_DIRECTORY)
83     CreateFolder strFolderName

       ' Create timestamped files.
84     Do While Now < dtEnd

           ' Ensure a valid Internet connection exists.
85         Do While (Not IsInternetConnected()) And (Now < dtEnd)
               ' Frequency in Hertz = 800, duration in milliseconds = 500.
               ' N.B. Using parameter names in VBA is slower than positional calls.
86             BeepAPI 800, 500
87             Application.StatusBar = "Taking samples paused. No Internet connection."

               ' Use DoEvents to maintain some UI responsiveness.
88             DoEvents

               ' Pause to avoid thrashing the connection testing function.
               ' Also, an audio alert faster than every three seconds is annoying.
89             Application.Wait (Now + TimeSerial(0, 0, 3))
90         Loop

91         For lngX = 0 To lngMaxIndex
92             If DateAdd("s", CameraArray(lngX).SampleRate, CameraArray(lngX).LastSample) <= Now Then
93                 DownloadImage CameraArray(lngX), strFolderName, lngFilesTried, lngFilesSuccessful
94                 CameraArray(lngX).LastSample = Now
95             End If
96         Next lngX

           ' Alert if any samples were missed.
97         If lngFilesTried <> lngFilesSuccessful Then
               ' Frequency in Hertz = 800, duration in milliseconds = 500.
               ' N.B. Using parameter names in VBA is slower than positional calls.
98             BeepAPI 800, 500
99         End If

           ' The shortest refresh rate is 15 sec.
           ' So set the wait between cycles = 15 sec.
           ' Use DoEvents to maintain some UI responsiveness.
100        DoEvents
101        Application.Wait (Now + TimeSerial(0, 0, 15))
102    Loop

103    strMessage = "Samples tried:" & vbTab & lngFilesTried & "." & vbNewLine & "Samples taken:" & vbTab & lngFilesSuccessful & "." & vbNewLine & vbNewLine & "Webcam sampling complete."
104    If lngFilesSuccessful > 0 Then
105       strMessage = strMessage & vbNewLine & "See " & PARENT_DIRECTORY & "."
106    End If
107    MsgBox strMessage, vbOKOnly, "Take Samples"
108    Application.StatusBar = False    ' Return Status Bar control to Excel.

    Exit Sub

ErrorHandler:
    DisplayErrorMessage "DownloadFileAPI", Err.Description, Err.Number, Erl

End Sub

Function CreateFolder(ByVal strPath As String) As Boolean

    On Error GoTo ErrorHandler

1    CreateFolder = False

     ' Create the target folder if it doesn't exist.
     ' vbDirectory limits Dir$ to directory matches.
2    If Len(Dir$(strPath, vbDirectory)) = 0 Then
3        MkDir strPath
4    End If

5    CreateFolder = True

    Exit Function

ErrorHandler:
    DisplayErrorMessage "CreateFolder", Err.Description, Err.Number, Erl

End Function

Function CreateTimestampedFolderName(ByVal strFilePrefix As String) As String

    ' Create a date-time stamped folder name.

    On Error GoTo ErrorHandler

1    CreateTimestampedFolderName = strFilePrefix & "samples" & Format$(Now, "-yyyy-mmm-dd-hh\h-mm\m-ss\s") & "\"

    Exit Function

ErrorHandler:
    DisplayErrorMessage "CreateTimestampedFolderName", Err.Description, Err.Number, Erl

End Function

Function Floor(dblX As Double, dblFactor As Double) As Long

    ' dblX is the value you want to round.
    ' Factor is the multiple to which you want to round.

    On Error GoTo ErrorHandler

1    Floor = CLng(Int(dblX / dblFactor) * dblFactor)

    Exit Function

ErrorHandler:
    MsgBox Err.Number & ": " & Error.Description

End Function

Function FolderExists(strPath As String) As Boolean

    On Error GoTo ErrorHandler

     ' vbDirectory limits Dir$ to directory matches.
1    FolderExists = (Len(Dir$(strPath, vbDirectory)) > 0)

    Exit Function

ErrorHandler:
    DisplayErrorMessage "FolderExists", Err.Description, Err.Number, Erl

End Function

Function HasBadCharacters(ByVal strLabel As String) As Boolean

    ' This function checks if a webcam labels contains characters that are:
    ' a) forbidden in Windows file names, or
    ' b) problematic in URLs.

    Dim strMessage As String
    Dim lngX As Long

    On Error GoTo ErrorHandler

      ' Non-printable ASCII control characters.
      ' - ASCII: 0-31.
      ' - Character: DEL, ASCII: 127.

      ' Printable forbidden characters.
      ' - Character: ", ASCII: 34.
      ' - Character: *, ASCII: 42.
      ' - Character: /, ASCII: 47.
      ' - Character: :, ASCII: 58.
      ' - Character: <, ASCII: 60.
      ' - Character: >, ASCII: 62.
      ' - Character: ?, ASCII: 63.
      ' - Character: \, ASCII: 92.
      ' - Character: |, ASCII: 124.

      ' Problematic in URLs.
      ' - Character: SPACE, ASCII: 32.
      ' - Character: #, ASCII: 35.
      ' - Character: %, ASCII: 37.
      ' - Character: &, ASCII: 38.

      ' Check for forbidden characters.
1     For lngX = 1 To Len(strLabel)

2         Select Case Asc(Mid(strLabel, lngX, 1))
              Case 0 To 31, 127
                  ' Non-printable ASCII control characters.
3                 strMessage = "Webcam labels may not contain ASCII control character " & Asc(Mid(strLabel, lngX, 1)) & "."
4                 HasBadCharacters = True

5             Case 34, 35, 37, 38, 42, 47, 58, 60, 62, 63, 92, 124
                  ' a) Printable forbidden characters.
                  ' b) Other characters problematic for URLs.
6                 strMessage = "Webcam labels may not contain character: " & Mid(strLabel, lngX, 1)
7                 HasBadCharacters = True

8             Case 32
                  ' Spaces, problematic for URLs.
9                 strMessage = "Webcam labels may not contain spaces." & vbNewLine & "Consider '-' or '_' instead."
10                HasBadCharacters = True
11        End Select

12        If HasBadCharacters Then
13            strMessage = strMessage & vbNewLine & "Operation cancelled."
14            DisplayValidationMessage "Take Samples", strMessage
15            Exit Function
16        End If

17    Next lngX

18    HasBadCharacters = False

    Exit Function

ErrorHandler:
    DisplayErrorMessage "HasBadCharacters", Err.Description, Err.Number, Erl

End Function

Function IsInternetConnected() As Boolean

    Dim lngFlags As Long

    On Error GoTo ErrorHandler

1    IsInternetConnected = CBool(InternetGetConnectedState(lngFlags, 0&))

    Exit Function

ErrorHandler:
    DisplayErrorMessage "IsInternetConnected", Err.Description, Err.Number, Erl

End Function

Sub DisplayErrorMessage(ByVal strOperationName As String, ByVal strErrDescription As String, ByVal lngErrNumber As Long, ByVal intErrLine As Integer)

    On Error GoTo ErrorHandler

1    MsgBox "Error " & lngErrNumber & ": " & strErrDescription & ", on line " & intErrLine & ".", vbCritical, strOperationName & " Error"
2    Application.StatusBar = False    ' Return Status Bar control to Excel.

    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & ", on line " & Erl & ".", vbCritical, "DisplayErrorMessage Error"
    Application.StatusBar = False    ' Return Status Bar control to Excel.

End Sub

Sub DisplayValidationMessage(ByVal strOperationName As String, ByVal strValidationMessage As String)

    On Error GoTo ErrorHandler

1    MsgBox strValidationMessage, vbExclamation, strOperationName

    Exit Sub

ErrorHandler:
    DisplayErrorMessage "DisplayValidationMessage", Err.Description, Err.Number, Erl

End Sub

Sub DownloadImage(ByRef tmpCamera As Camera, ByVal strFolderName As String, ByRef lngFilesTried As Long, ByRef lngFilesSuccessful As Long)

    On Error GoTo ErrorHandler

1    lngFilesTried = lngFilesTried + 1

     ' Delete cached data here, from both successful and unsuccessful downloads.
2    DeleteUrlCacheEntry tmpCamera.URL

     ' a) Explicitly pass zeros as longs via 0&, instead of as integers via 0.
     ' b) The local file path is: strFolderName & tmpCamera.Label & Format$(Now, "-yyyy-mmm-dd-hh\h-mm\m-ss\s") & tmpCamera.FileType
3    If URLDownloadToFile( _
         0&, _
         tmpCamera.URL, _
         strFolderName & tmpCamera.Label & Format$(Now, "-yyyy-mmm-dd-hh\h-mm\m-ss\s") & tmpCamera.FileType, _
         0&, _
         0&) = 0 Then
4        lngFilesSuccessful = lngFilesSuccessful + 1
5    End If

    Exit Sub

ErrorHandler:
    DisplayErrorMessage "DownloadImage", Err.Description, Err.Number, Erl

End Sub

