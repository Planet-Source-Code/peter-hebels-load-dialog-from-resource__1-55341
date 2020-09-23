VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resource dialog loader"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Written in dialog text box:"
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3135
      Begin VB.TextBox Text1 
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load dialog from resource"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************
'Load Dialog From Resoucefile project written by Peter Hebels (http://www.phsoft.nl)*
'Please don't remove this message header when distributing this code in source form *
'                                                                                   *
'I've written this just for fun, don't expect anything special from it...           *
'                                                                                   *
'Don't forget you have to compile this project to an exe, otherwise it cannot read  *
'the resource file and thus cannot load the dialog from it.                         *
'                                                                                   *
'************************************************************************************

Private Sub Command1_Click()
    'Call the dialog loading function located in the module file, here it all begins :)
    LoadDialogRes
End Sub

Private Sub Form_Load()
    'Check if project is compiled to an exe and not running in the IDE
    If RunningInVbIDE <> 0 Then
        MsgBox "Please compile this project first before running it.", vbExclamation, "Error"
        Unload Me
    End If
End Sub
