VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "FLASH.OCX"
Begin VB.Form Form1 
   Caption         =   "Visual with Flash"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton play2 
      Caption         =   "Play"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   5400
      Width           =   5055
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf2 
      Height          =   1935
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Width           =   7455
      _cx             =   13150
      _cy             =   3413
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.CommandButton play1 
      Caption         =   "Play"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   2880
      Width           =   4935
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash swf 
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      _cx             =   12515
      _cy             =   3201
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'swf1.Movie = vote.swf Select your *.swf file
'swf2.Movie = deneme.swf Select your *.swf file
End Sub

Private Sub play2_Click()
file2 = txt1.Text
swf2.Movie = file2
End Sub
Private Sub play1_Click()
Dim dosya As String
file1 = txt.Text
file2 = txt1.Text
swf.Movie = file1
End Sub

