VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   405
      Left            =   3600
      TabIndex        =   3
      Top             =   1875
      Width           =   1260
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   405
      Left            =   2220
      TabIndex        =   2
      Top             =   1875
      Width           =   1290
   End
   Begin VB.TextBox txtCnn 
      Height          =   1440
      Left            =   165
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   300
      Width           =   6945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection String"
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   105
      Width           =   1260
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdEdit_Click()
Dim myCnnstring As String
Dim oWizard As CnnWizard
        
        'new object cnnWizard
        Set oWizard = New CnnWizard
        
        'connection string to edit
        myCnnstring = txtCnn.Text
        
        'Edit Connection String
        If oWizard.EditCnn(Me.hWnd, myCnnstring) Then
            'ok, edited
            txtCnn.Text = myCnnstring
        Else
            'cancel
            MsgBox "Operation cancelled."
        End If
        'release object
        Set oWizard = Nothing

End Sub

Private Sub cmdNew_Click()
Dim myCnnstring As String
Dim oWizard As CnnWizard
        
        'new object cnnWizard
        Set oWizard = New CnnWizard
        
        'get new Connection String
        If oWizard.NewCnn(Me.hWnd, myCnnstring) Then
            'ok
            txtCnn.Text = myCnnstring
        Else
            'cancel
            MsgBox "Operation cancelled."
        End If
        'release object
        Set oWizard = Nothing


End Sub


