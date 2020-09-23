VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Deleting Multiple Lines from MSFlexGrid"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3836
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Call MSFlexGrid1_KeyUp(vbKeyDelete, 0)
End Sub

Private Sub cmdRefresh_Click()
Dim db As Database
Dim SQL As String
Dim rs As Recordset

    MSFlexGrid1.Visible = False
    
    Form1.MSFlexGrid1.Col = 1
    
    Set db = OpenDatabase(App.Path & "\db1.mdb")
    SQL = ("SELECT * From Table1")
    Set rs = db.OpenRecordset(SQL, dbOpenDynaset)
        
    With MSFlexGrid1
        .Cols = 3
        .Rows = 2
        .Visible = True
    
        .ColWidth(0) = 300
        .ColWidth(1) = 1000
        .Col = 1
        .Row = 0
        .Text = "AutoNumColumn"
        
        .ColWidth(2) = 1000
        .Col = 2
        .Row = 0
        .Text = "TextColumn"
        
        Do While Not rs.EOF
            .AddItem "" & vbTab & rs![AutoNumColumn] & _
                          vbTab & rs![TextColumn]
            rs.MoveNext
        Loop
            
        .RemoveItem 1
        
        .Visible = True
    
    End With
End Sub

Private Sub Form_Load()
    cmdRefresh_Click
End Sub

Private Sub MSFlexGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim db As Database
Dim SQL As String
Dim rs As Recordset
Dim i As Long
Dim iStep As Integer
Dim strWhere As String
Dim RowCount As Integer
Dim Ans As String

Set db = OpenDatabase(App.Path & "\db1.mdb")

    If KeyCode = vbKeyDelete Then
        Me.MousePointer = vbHourglass
        With MSFlexGrid1
            If .Row > .RowSel Then
                RowCount = (.Row - .RowSel) + 1
                iStep = -1
            Else
                iStep = 1
                RowCount = (.RowSel - .Row) + 1
            End If
            
            Ans = MsgBox("Are you sure you want to deleted the selected " & RowCount & " record(s)?", vbYesNo + vbCritical + vbDefaultButton2)
            
            Select Case Ans
                Case vbYes
                    MSFlexGrid1.Visible = False
                        For i = .Row To .RowSel Step iStep
                            If Len(strWhere) > 0 Then
                                strWhere = strWhere & " or "
                            End If
                            
                            strWhere = strWhere & "AutoNumColumn = " & .TextMatrix(i, 1)
                            SQL = "DELETE From Table1 Where " & strWhere
                            db.Execute SQL
                        Next
                        db.Close
                        Me.MousePointer = vbNormal
                Case vbNo
                    db.Close
                    Me.MousePointer = vbNormal
                    Exit Sub
            End Select
        End With
        
        cmdRefresh_Click
        
        MsgBox "Delete completed."
    End If
End Sub
