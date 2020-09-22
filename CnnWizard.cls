VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CnnWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function EditCnn(ByVal phWnd As Long, ByRef pCnnStringToModify As String) As Boolean
Dim myCnnstring As String
Dim myCnn As ADODB.Connection
Dim AdoWizard As DataLinks
                    
        Set myCnn = New ADODB.Connection
        myCnn.ConnectionString = pCnnStringToModify
        Set AdoWizard = New DataLinks
        AdoWizard.hWnd = phWnd
        
        'Edit cnnString
        If AdoWizard.PromptEdit(myCnn) Then
            'Ok edit
            EditCnn = True
            pCnnStringToModify = myCnn.ConnectionString
            Set myCnn = Nothing
            Set AdoWizard = Nothing
            Exit Function
        Else
            'Cancel
            EditCnn = False
            Set myCnn = Nothing
            Set AdoWizard = Nothing
            Exit Function
        End If

End Function
Public Function NewCnn(ByVal phWnd As Long, ByRef pCnnStringResult As String) As Boolean
Dim myCnnstring As String
Dim myCnn As ADODB.Connection
Dim AdoWizard As DataLinks
                    
        Set AdoWizard = New DataLinks
        AdoWizard.hWnd = phWnd
        'new cnn string
        Set myCnn = AdoWizard.PromptNew
        If Not myCnn Is Nothing Then
            'Ok new
            NewCnn = True
            pCnnStringResult = myCnn.ConnectionString
            Set myCnn = Nothing
            Set AdoWizard = Nothing
            Exit Function
        Else
            'Cancel
            NewCnn = False
            Set myCnn = Nothing
            Set AdoWizard = Nothing
            Exit Function
        End If

End Function

