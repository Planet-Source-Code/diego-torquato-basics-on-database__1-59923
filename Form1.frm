VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "DataSample"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   4935
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtClass 
         Height          =   285
         Left            =   3600
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtNumber 
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblClass 
         Caption         =   "Class:"
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblNumber 
         Caption         =   "Number:"
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Students"
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Name"
         Caption         =   "Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Number"
         Caption         =   "Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Class"
         Caption         =   "Class"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is a tutorial on Database programming
' It explains just the basic stuff on it

' In the beginning, we need to put all our referencies and
' components
' They are:
' Referencies: Microsoft ActiveX Data Objects 2.0 Library
' Components: Microsoft DataGrid Control 6.0
' You're gonna find them ready to be added in Project>Referencies
' and Project>Components

' Basically these tools are used to communicate with the database

' From now, I'm gonna tell step by step what is needed to be
' done to Open, Edit and Delete data in a database(db)
' These steps will be ordered by numbers, it may be out of order,
' in this case, fallow number to number, without pass out. Instead
' you're not gonna get it.

' 1. Declarations
Public CN As New ADODB.Connection ' Conection tool
Public RS As New ADODB.Recordset ' Recordset tool


' 2. Utilities
' The Basic declarations have been done, now we just add a DataGrid
' (DataGrid1) and we change it the way to contain the fields needed
' (correponding to the db)

' The database was created like this:
' |Table: Students
' |Fields: Name, Number, Class
' The database was previouslly created with Microsoft Access 2000

' We must add these fields to the DataGrid.
' Click in it with the right mouse button then click 'Edit'
' Give a right click again, you're gonna see the following options:
' Cut, Copy, Paste, Delete, Insert and Append
' I don't think I need to explain it.
' After adding the right number of fields click with the right
' mouse button again and go to Properties. Click in Columns,
' so in each column put the correspondings Caption and Datafield
' Where Caption is the 'Title' that will be shown on the Datagrid
' and Datafield is the corresponding name of the fields on db

' Well, we already have the basic, the db and the DataGrid
' I'm not gonna specify the TextBoxes e Commands buttons added in the
' project, pay attention to the names.
' I've just added 1 Frame, 3 Labels, 3 TextBoxes e 3 CommandsButtons.

' Now we have the main desing ready, now we are going to the best
' part: coding

' 3. Getting Started
Private Sub Form_Load()
    ' Opening database conection
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\datasample.mdb"
    ' The following instruction must be executed before opening our recordset
    RS.CursorLocation = adUseClient
    ' Opening the recordset
    RS.Open "SELECT * FROM Students ORDER BY Name ASC", CN, adOpenStatic, adLockOptimistic
    ' Now we set our DataGrid to have the content of the db
    Set DataGrid1.DataSource = RS
    ' Ready, we already learnt how to read the db.
    ' This datagrid is configurated the way that could be possible change
    ' the content of db through it. It can be easylly configurated
    
    ' We need to configurate our soft to fill the textboxes
    ' so we will be able to change the content of the db.
    
    With RS
        ' If not empty...
        If Not .BOF Then
            ' Put the content of our current recordset in the
            ' textboxes
            txtName.Text = .Fields("Name")
            txtNumber.Text = .Fields("Number")
            txtClass.Text = .Fields("Class")
        End If
    End With
End Sub

' 4. Let's configurate the buttons
Private Sub cmdSave_Click()
    ' This button will save the changes and additions
    With RS
        ' If our Textboxes not empty
        If txtName.Text <> "" And txtNumber.Text <> "" And txtClass.Text <> "" Then
            ' Put the content of our textboxes in the current
            ' recordset
            .Fields("Name") = txtName.Text
            .Fields("Number") = txtNumber.Text
            .Fields("Class") = txtClass.Text
            .Update
        End If
    End With
    ' Ready, now we're able to change our data
End Sub

' We need now to add the fields:
Private Sub cmdNew_Click()
    ' Set defaults values to our txts
    txtName.Text = "Name"
    txtNumber.Text = "Number"
    txtClass.Text = "Class"
    ' Create a new registry in the db and wait for data
    RS.AddNew
    ' After clicking this button, you change the data
    ' and click in Save
End Sub

' How to delete?
Private Sub cmdDelete_Click()
    With RS
        ' Conect to the db sending a SQL query asking to delete
        CN.Execute "DELETE * FROM Students WHERE Name ='" & .Fields("Name") & "'"
        .Requery ' update our datagrid
    End With
End Sub

' Well, I think I've finished. This is the basic of the basic
' I hope have helped someone who would like to learn this...

' Well, that's my first usefull (ok, not that much) tutorial,
' do not criticize me for it.

' Diego Torquato (binary)
' torquato@totalsecurity.com.br
