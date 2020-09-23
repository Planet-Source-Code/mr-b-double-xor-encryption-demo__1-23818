VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "XOR Encryption - Double XOR"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4680
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   735
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   609
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   180
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Text File (*.txt)|*.txt|All Files (*.*)|*.*"
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '¾Æ·¡ ¸ÂÃã
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   529
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdDecrypt 
      Caption         =   "Decrypt File"
      Height          =   495
      Left            =   2509
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdEncrypt 
      Caption         =   "Encrypt File"
      Height          =   495
      Left            =   957
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================
'Coded by: Mr.B
'If you have any comments or suggestions,
'Please e-mail me at: bleetcs@hotmail.com
'                    or blee@tcs.on.ca
'========================================

'description:
'just a program that demonstrates double XOR encryption.
'this only works for text files.
'*note: obvious procedures only noted ONCE.

Option Explicit

Dim FileName As String  'we have to know which file we want to encrypt or decrypt

Private Sub CmdDecrypt_Click()
    Dim temp As String  'just a temporary storage
    Dim PWD As String   'password
    Dim CurPos As Long  'where are we in the text file? for progress bar
    Dim Total As Long   'total number of lines on the text file
    
    CD.ShowOpen             'show common dialog box
    FileName = CD.FileName  'get file name
    CurPos = 0              'initialize the variable
        
    'only works when a file is chosen
    If Len(FileName) <> 0 Then
        PWD = InputBox("Enter Password")    'get password
        'only works when a password is chosen
        If Len(PWD) <> 0 Then
            
            'if "decrypted.txt" exists, 'kill' it!
            If Len(Dir(App.Path + "\decrypted.txt")) > 0 Then Kill App.Path + "\decrypted.txt"
            
            'get total number of lines in a text file
            Total = FindLineFile(FileName)
            
            'print status
            StatusBar.SimpleText = "Start Decoding..."
            
            Open FileName For Input As #1
                Do
                    CurPos = CurPos + 1
                    
                    'get encoded string
                    Line Input #1, temp
                    
                    'decrypt the string
                    temp = XORDecrypt(XORDecrypt(temp, PWD), PWD)
                    
                    'write into the output file
                    Open App.Path + "\decrypted.txt" For Append As #2
                        Print #2, temp
                    Close #2
                    
                    'show progress
                    ProgBar.Value = (CurPos / Total) * 100
                
                Loop Until EOF(1) = True
            Close #1
            
            'print status
            StatusBar.SimpleText = "Decoding Done..."
        End If
    End If
End Sub

Private Sub CmdEncrypt_Click()
    Dim temp As String
    Dim PWD As String
    Dim CurPos As Long
    Dim Total As Long
    
    CD.ShowOpen
    FileName = CD.FileName
    CurPos = 0
    
    If Len(FileName) = 0 Then Exit Sub
    
    PWD = InputBox("Enter Password")
    
    If Len(PWD) = 0 Then Exit Sub
    
    If Len(Dir(App.Path + "\encrypted.txt")) > 0 Then Kill App.Path + "\encrypted.txt"
    
    Total = FindLineFile(FileName)
    
    StatusBar.SimpleText = "Start Encoding..."
    Open FileName For Input As #1
        While Not EOF(1)
            CurPos = CurPos + 1
            Line Input #1, temp
            
            'encrypt string
            temp = XOREncrypt(XOREncrypt(temp, PWD), PWD)
            
            Open App.Path + "\encrypted.txt" For Append As #2
                Print #2, temp
            Close #2
            
            ProgBar.Value = (CurPos / Total) * 100
        Wend
    Close #1
    StatusBar.SimpleText = "Done Encoding..."
End Sub

Private Sub Form_Load()
    StatusBar.SimpleText = "Welcome!"
End Sub

Function FindLineFile(FileN As String) As Long
    Dim temp As String
    
    FindLineFile = 0
    
    'count the number of lines in a text file
    Open FileN For Input As #1
        While Not EOF(1)
            Line Input #1, temp
            FindLineFile = FindLineFile + 1
        Wend
    Close #1
End Function
