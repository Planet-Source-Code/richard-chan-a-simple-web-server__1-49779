VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "0 Clients Connected."
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   4560
      TabIndex        =   7
      Top             =   360
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "IP Connected"
      Height          =   3135
      Left            =   4440
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Most Recent Request"
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4215
      Begin VB.TextBox Text2 
         Height          =   2295
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock sender 
      Index           =   0
      Left            =   3600
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock listener 
      Left            =   3600
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin VB.Label Label4 
      Caption         =   "Check log.txt for log [Click here]"
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Server developed by Richard Chan."
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Default document:"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Please use the wwwroot directory for the site directory."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**************************************
' Name: A Simple Web Server
' Description: a simple web server
'     using multiple winsocks and
'     packet sending to deliver files.
' By: Richard Chan
'
' *Please vote if you like the code*
'**************************************

' array to check if sock is sending
Dim sending(255) As Boolean
' array to stop a certain sending progress
Dim stopSend(255) As Boolean
' array to check if a certain sock is free
Dim freeConn(255) As Boolean

Private Sub Form_Load()
' check default document from file
    Dim def As String
    Open App.Path & "\conf.ini" For Input As #3
    Input #3, def
    Close #3
    Text1.Text = def
' loop to initialize winsocks
    Dim i As Integer
    For i = 0 To 255
        If i <> 0 Then Load sender(i)
        freeConn(i) = True
        sending(i) = False
        stopSend(255) = False
    Next
' have the listener sock start listening
    listener.Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim fileNum As Integer
' save default document selection
    fileNum = FreeFile
    Open App.Path & "\conf.ini" For Output As #fileNum
        Print #fileNum, Text1.Text
    Close #fileNum
' end the program directly in case some sock is still in use
    End
End Sub

Private Sub Label4_Click()
    Shell "notepad.exe " & App.Path & "\log.txt", vbMaximizedFocus
End Sub

Private Sub listener_Close()
' have the listener reset in case it closes in error, probably unnecessary
    listener.Close
    listener.Listen
End Sub

Private Sub listener_ConnectionRequest(ByVal requestID As Long)
' have a free sock connect to the request
    Dim i As Integer
    For i = 0 To 255
        If freeConn(i) Then
            freeConn(i) = False
            sender(i).Close
            sender(i).Accept requestID
' take log of connected ip
            List1.AddItem sender(i).RemoteHostIP
            Form1.Caption = List1.ListCount & " Clients Connected."
            Exit For
        End If
    Next
End Sub

Private Sub sendfile(file As String, sock As Integer)
' substitute for invalid characters
    file = Replace(file, "/", "\")
    file = Replace(file, "%20", " ")
' set up error handler in case there's error with the file
    On Error GoTo handler
    Dim fileNum As Integer
    Dim fileBin As String
    Dim fileSize As Long
    Dim sentSize As Long
    Dim i As Integer

' prepare opening the file requested
    fileSize = FileLen(file)
    fileNum = FreeFile
    Open file For Binary As #fileNum
        
' have it send 1024 bit a time
        fileBin = Space(1024)
        
        Do
' set sending for that sock to true so that we can check its progress
            sending(sock) = True
' get the packet
            Get #fileNum, , fileBin
' calculate amount sent
            sentSize = sentSize + Len(fileBin)
' check if it will be done or not in the next packet
            If sentSize > fileSize Then
                sender(sock).SendData Mid(fileBin, 1, Len(fileBin) - (sentSize - fileSize))
            Else
                sender(sock).SendData fileBin
            End If
            Do
' wait until sending is done -- the sending variable is changed by the sock's sendcomplete event
                DoEvents
' if it is to be stopped in the middle, send it to the error handler
                If stopSend(sock) Then GoTo handler
            Loop Until sending(sock) = False
            
        DoEvents
' keep sending until the file is sent
        Loop Until EOF(fileNum)
    
' close file and free sock
    Close (fileNum)
' remove ip log
    For i = 0 To List1.ListCount - 1
        If List1.List(i) = sender(sock).RemoteHostIP Then
            List1.RemoveItem i
            Form1.Caption = List1.ListCount & " Clients Connected."
            Exit For
        End If
    Next
    sender(sock).Close
    freeConn(sock) = True
    
    Exit Sub
handler:

' tell client about the error
    If sender(sock).State = 7 Then
        sending(sock) = True
        sender(sock).SendData "Internal Error" & vbNewLine
        Do
            DoEvents
        Loop Until sending(sock) = False
    End If
    
' close and free sock
' remove ip log
    For i = 0 To List1.ListCount - 1
        If List1.List(i) = sender(sock).RemoteHostIP Then
            List1.RemoveItem i
            Form1.Caption = List1.ListCount & " Clients Connected."
            Exit For
        End If
    Next
    sender(sock).Close
    freeConn(sock) = True
    stopSend(sock) = False
End Sub

Private Sub sender_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim dat As String, file As String, start As Integer, start2 As Integer, fileNum As Integer
' get request
    sender(Index).GetData dat
' take log
    fileNum = FreeFile
    Open App.Path & "\log.txt" For Append As #fileNum
    Print #fileNum, "Client " & sender(Index).RemoteHostIP & " Request:"
    Print #fileNum, dat
    Close (fileNum)
    Text2.Text = dat
' check request, if the first 3 are "GET" then it is a request for getting a file
    If Mid(dat, 1, 3) = "GET" Then
' check position for "GET "
        start = InStr(dat, "GET ")
' check position for the end of the file name
        start2 = InStr(start + 5, dat, " ")
' get the file name
        file = Mid(dat, start + 5, start2 - (start + 4))
' trim the file name for ending space
        file = RTrim(file)
' if name is empty, it means it is something like ".../" so it will call the default file
        If file = "" Or Right(file, 1) = "/" Then file = file & Text1.Text
' send file
        sendfile App.Path & "\wwwroot\" & file, Index
    End If
End Sub

Private Sub sender_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
' stop send if experience error
    stopSend(Index) = True
End Sub

Private Sub sender_SendComplete(Index As Integer)
' set the sending variable for the sock to false when sending is done
    sending(Index) = False
End Sub
