VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SPTransfer"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Login"
      Height          =   1215
      Left            =   4560
      TabIndex        =   17
      Top             =   600
      Width           =   3855
      Begin VB.TextBox txtPass2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   25
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtUser2 
         Height          =   285
         Left            =   1440
         TabIndex        =   24
         Text            =   "sa"
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton optSQL2 
         Caption         =   "SQL security"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   3135
      End
      Begin VB.OptionButton optNT2 
         Caption         =   "NT security"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login"
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   3855
      Begin VB.TextBox txtPass1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtUser1 
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Text            =   "sa"
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton optSQL1 
         Caption         =   "SQL security"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   3135
      End
      Begin VB.OptionButton optNT1 
         Caption         =   "NT security"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdRen2 
      Caption         =   "Rename"
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdDel2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdRen1 
      Caption         =   "Rename"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdDel1 
      Caption         =   "Delete"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cmdDOWN 
      Caption         =   "<-"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton cmdUP 
      Caption         =   "->"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   2400
      Width           =   375
   End
   Begin VB.CommandButton cmdConnect2 
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton cmdConnect1 
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox txtDB2 
      Height          =   285
      Left            =   6240
      TabIndex        =   5
      Text            =   "northwind"
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtServer2 
      Height          =   285
      Left            =   4560
      TabIndex        =   4
      Text            =   "localhost"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtDB1 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "northwind"
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtServer1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "localhost"
      Top             =   240
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   4545
      Left            =   4560
      MultiSelect     =   2  'Extended
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   1920
      Width           =   3855
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   120
      MultiSelect     =   2  'Extended
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Server                           Database"
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Server                           Database"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst1 As ADODB.Recordset
Dim con1 As ADODB.Connection
Dim rst2 As ADODB.Recordset
Dim con2 As ADODB.Connection

Private Sub cmdConnect1_Click()

    On Error GoTo ErrHandler

    Set con1 = New ADODB.Connection
    con1.Provider = "SQLOLEDB"
    con1.ConnectionString = "Server=" & txtServer1.Text & ";Database=" & txtDB1.Text & ";"
    If optNT1.Value Then
        con1.ConnectionString = con1.ConnectionString & "Trusted_Connection=yes;"
    Else
        con1.ConnectionString = con1.ConnectionString & "UID=" & txtUser1.Text & ";PWD=" & txtPass1.Text & ";"
    End If
    
    con1.Open
    
    Set rst1 = con1.Execute("sp_help")

    rst1.Filter = "Object_type='stored procedure'"
    List1.Clear
    Do Until rst1.EOF
        List1.AddItem rst1("Name")
        rst1.MoveNext
    Loop

    Exit Sub

ErrHandler:

    MsgBox Err.Number & vbCrLf & Err.Description

End Sub

Private Sub cmdConnect2_Click()

    On Error GoTo ErrHandler
    
    Set con2 = New ADODB.Connection
    con2.Provider = "SQLOLEDB"
    con2.ConnectionString = "Server=" & txtServer2.Text & ";Database=" & txtDB2.Text & ";"
    
    If optNT2.Value Then
        con2.ConnectionString = con2.ConnectionString & "Trusted_Connection=yes;"
    Else
        con2.ConnectionString = con2.ConnectionString & "UID=" & txtUser2.Text & ";PWD=" & txtPass2.Text & ";"
    End If
    
    con2.Open
    
    Set rst2 = con2.Execute("sp_help")

    rst2.Filter = "Object_type='stored procedure'"
    List2.Clear
    Do Until rst2.EOF
        List2.AddItem rst2("Name")
        rst2.MoveNext
    Loop

    Exit Sub

ErrHandler:

    MsgBox Err.Number & vbCrLf & Err.Description

End Sub

Private Sub cmdDel1_Click()

    List1_KeyUp vbKeyDelete, 0

End Sub

Private Sub cmdDel2_Click()

    List2_KeyUp vbKeyDelete, 0

End Sub

Private Sub cmdDOWN_Click()

    On Error GoTo ErrHandler

    If con1 Is Nothing Or con2 Is Nothing Then
        MsgBox "You must be connected to both databases"
        Exit Sub
    End If
    
    If con1.State = adStateClosed Or con2.State = adStateClosed Then
        MsgBox "You must be connected to both databases"
        Exit Sub
    End If

    Dim T As Integer
    Dim R As Integer
    Dim strText As String
    Dim Processed As Integer
    Dim RC As Long
    Processed = 0
    
    If List2.SelCount <> 0 Then
        frmProcess.Show
        DoEvents
        For T = 0 To List2.ListCount - 1
            If List2.Selected(T) Then
                ' must transfer thisone
                frmProcess.lblCurrent.Caption = List2.List(T)
                
                ' check if exists on target server
                For R = 0 To List1.ListCount - 1
                    If List2.List(T) = List1.List(R) Then
                        ' drop it
                        con1.Execute "DROP PROC " & List1.List(R)
                    End If
                Next R
                
                Set rst2 = New ADODB.Recordset
                rst2.CursorLocation = adUseClient
                rst2.Open "sp_helptext " & List2.List(T), con2, adOpenDynamic, adLockOptimistic
                strText = ""
                RC = 0
                Do Until rst2.EOF
                    ' get text
                    RC = RC + 1
                    frmProcess.PC.Value = CInt(((RC / rst2.RecordCount) * 100) * (85 / 100)): DoEvents
                    strText = strText & rst2("Text")
                    rst2.MoveNext
                Loop
                If con1.State <> adStateClosed Then
                    ' create on other side
                    con1.Execute strText
                End If
                Processed = Processed + 1
                frmProcess.PO.Value = CInt((Processed / List2.SelCount) * 100): DoEvents
            End If
        Next T
        Unload frmProcess
    End If

    ' refresh target server
    cmdConnect1_Click

    Exit Sub

ErrHandler:

    MsgBox Err.Number & vbCrLf & Err.Description


End Sub

Private Sub cmdRen1_Click()

    List1_KeyUp vbKeyF2, 0

End Sub

Private Sub cmdRen2_Click()

    List2_KeyUp vbKeyF2, 0

End Sub

Private Sub cmdUP_Click()

    On Error GoTo ErrHandler

    If con1 Is Nothing Or con2 Is Nothing Then
        MsgBox "You must be connected to both databases"
        Exit Sub
    End If
    
    If con1.State = adStateClosed Or con2.State = adStateClosed Then
        MsgBox "You must be connected to both databases"
        Exit Sub
    End If

    Dim T As Integer
    Dim R As Integer
    Dim strText As String
    Dim Processed As Integer
    Dim RC As Long
    Processed = 0
    
    If List1.SelCount <> 0 Then
        frmProcess.Show
        DoEvents
        For T = 0 To List1.ListCount - 1
            If List1.Selected(T) Then
                ' must transfer thisone
                frmProcess.lblCurrent.Caption = List1.List(T)
                
                ' check if exists on target server
                For R = 0 To List2.ListCount - 1
                    If List1.List(T) = List2.List(R) Then
                        ' drop it
                        con2.Execute "DROP PROC " & List2.List(R)
                    End If
                Next R
                
                Set rst1 = New ADODB.Recordset
                rst1.CursorLocation = adUseClient
                rst1.Open "sp_helptext " & List1.List(T), con1, adOpenDynamic, adLockOptimistic
                strText = ""
                RC = 0
                Do Until rst1.EOF
                    ' get text
                    RC = RC + 1
                    frmProcess.PC.Value = CInt(((RC / rst1.RecordCount) * 100) * (85 / 100)): DoEvents
                    strText = strText & rst1("Text")
                    rst1.MoveNext
                Loop
                If con2.State <> adStateClosed Then
                    ' create on other side
                    con2.Execute strText
                End If
                Processed = Processed + 1
                frmProcess.PO.Value = CInt((Processed / List1.SelCount) * 100): DoEvents
            End If
        Next T
        Unload frmProcess
    End If

    ' refresh target server
    cmdConnect2_Click

    Exit Sub

ErrHandler:

    MsgBox Err.Number & vbCrLf & Err.Description


End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrHandler

    List1.Enabled = False
    If con1 Is Nothing Then
        MsgBox "You must be connected to a database"
        List1.Enabled = True
        Exit Sub
    End If
    
    If con1.State = adStateClosed Then
        MsgBox "You must be connected to a database"
        List1.Enabled = True
        Exit Sub
    End If
    List1.Enabled = True

    Dim T As Integer
    Dim R As Integer
    Dim Dup As Boolean
    
    Select Case KeyCode
    Case vbKeyDelete
        If List1.SelCount > 0 Then
            If vbYes = MsgBox("Delete all selected procedures?", vbYesNo + vbDefaultButton2, "Delete") Then
                For T = 0 To List1.ListCount - 1
                    If List1.Selected(T) Then
                        con1.Execute "DROP PROC " & List1.List(T)
                    End If
                Next T
            End If
        End If
        cmdConnect1_Click
    Case vbKeyF2
        Dim strNewName As String
        If List1.SelCount > 0 Then
            For T = 0 To List1.ListCount - 1
                If List1.Selected(T) Then
                    strNewName = InputBox("Please supply a new name", "Rename", List1.List(T))
                    If strNewName <> "" Then
                        Dup = False
                        For R = 0 To List1.ListCount - 1
                            If strNewName = List1.List(R) Then
                                MsgBox strNewName & " already exists"
                                Dup = True
                            End If
                        Next R
                            
                        If Dup Then Exit For
                        
                        ' this must be renamed
                        Set rst1 = New ADODB.Recordset
                        rst1.CursorLocation = adUseClient
                        rst1.Open "sp_helptext " & List1.List(T), con1, adOpenDynamic, adLockOptimistic
                        strText = ""
                        RC = 0
                        Do Until rst1.EOF
                            RC = RC + 1
                            frmProcess.PC.Value = CInt(((RC / rst1.RecordCount) * 100) * (85 / 100)): DoEvents
                            strText = strText & rst1("Text")
                            rst1.MoveNext
                        Loop
                        strText = Replace(strText, List1.List(T), strNewName)
                        If con1.State <> adStateClosed Then
                            con1.Execute strText
                            con1.Execute "DROP PROC " & List1.List(T)
                        End If
                    Else
                        Exit For
                    End If
                End If
            Next T
        End If
        cmdConnect1_Click
    End Select

    Exit Sub

ErrHandler:

    MsgBox Err.Number & vbCrLf & Err.Description


End Sub

Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrHandler
    
    List2.Enabled = False
    
    If con2 Is Nothing Then
        MsgBox "You must be connected to a database"
        List2.Enabled = True
        Exit Sub
    End If
    
    If con2.State = adStateClosed Then
        MsgBox "You must be connected to a database"
        List2.Enabled = True
        Exit Sub
    End If
    
    List2.Enabled = True
    
    Dim T As Integer
    Dim R As Integer
    Dim Dup As Boolean
    
    Select Case KeyCode
    Case vbKeyDelete
        If List2.SelCount > 0 Then
            If vbYes = MsgBox("Delete all selected procedures?", vbYesNo + vbDefaultButton2, "Delete") Then
                For T = 0 To List2.ListCount - 1
                    If List2.Selected(T) Then
                        con2.Execute "DROP PROC " & List2.List(T)
                    End If
                Next T
            End If
        End If
        cmdConnect2_Click
    Case vbKeyF2
        Dim strNewName As String
        If List2.SelCount > 0 Then
            For T = 0 To List2.ListCount - 1
                If List2.Selected(T) Then
                    strNewName = InputBox("Please supply a new name", "Rename", List2.List(T))
                    If strNewName <> "" Then
                        Dup = False
                        For R = 0 To List2.ListCount - 1
                            If strNewName = List2.List(R) Then
                                MsgBox strNewName & " already exists"
                                Dup = True
                            End If
                        Next R
                            
                        If Dup Then Exit For
                        
                        ' this must be renamed
                        Set rst2 = New ADODB.Recordset
                        rst2.CursorLocation = adUseClient
                        rst2.Open "sp_helptext " & List2.List(T), con2, adOpenDynamic, adLockOptimistic
                        strText = ""
                        RC = 0
                        Do Until rst2.EOF
                            RC = RC + 1
                            frmProcess.PC.Value = CInt(((RC / rst2.RecordCount) * 100) * (85 / 100)): DoEvents
                            strText = strText & rst2("Text")
                            rst2.MoveNext
                        Loop
                        strText = Replace(strText, List2.List(T), strNewName)
                        If con2.State <> adStateClosed Then
                            con2.Execute strText
                            con2.Execute "DROP PROC " & List2.List(T)
                        End If
                    Else
                        Exit For
                    End If
                End If
            Next T
        End If
        cmdConnect2_Click
    End Select

    Exit Sub

ErrHandler:

    MsgBox Err.Number & vbCrLf & Err.Description


End Sub
