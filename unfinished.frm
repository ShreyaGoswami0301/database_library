VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13020
   FillColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   Picture         =   "unfinished.frx":0000
   ScaleHeight     =   12930
   ScaleWidth      =   23760
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1095
      Left            =   1080
      Top             =   6720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1931
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Suvajit\Documents\LIBRARY1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Suvajit\Documents\LIBRARY1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "DETAILS"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PREVIOUS"
      Height          =   735
      Left            =   7320
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "NEXT"
      Height          =   735
      Left            =   4080
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CANCEL"
      Height          =   735
      Left            =   1080
      TabIndex        =   9
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DELETE"
      Height          =   735
      Left            =   7320
      TabIndex        =   8
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE"
      Height          =   735
      Left            =   4080
      TabIndex        =   7
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEW"
      Height          =   735
      Left            =   1080
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   4200
      TabIndex        =   5
      Text            =   "ENTER"
      Top             =   2400
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   4200
      TabIndex        =   4
      Text            =   "ENTER"
      Top             =   1320
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   4200
      TabIndex        =   3
      Text            =   "ENTER"
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "AUTHOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TITLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.Fields("Title") = Text1.Text



Adodc1.Recordset.Fields("Author") = Text2.Text
Adodc1.Recordset.Fields("Year") = Text3.Text

Adodc1.Recordset.Update







End Sub

Private Sub Command2_Click()
Adodc1.Recordset.AddNew


End Sub

Private Sub Command3_Click()
confirm = MsgBox("are you sure you want to delete?", vbYesNo, "Deletion Confirmation")
If confirm = vbYes Then
Adodc1.Recordset.Delete
MsgBox "record deleted", , "message"
Else
MsgBox "not deleted", , "Mesage"
End If

End Sub

Private Sub Command4_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "

End Sub

Private Sub Command5_Click()
If Not Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveNext
If Adodc1.recorset.EOF Then
Adoc1.Recordset.MovePrevious
End If
End If

End Sub

Private Sub Command6_Click()
End

End Sub
