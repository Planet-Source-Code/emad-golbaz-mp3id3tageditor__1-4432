VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmId3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Id3Tagger"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.mp3"
      DialogTitle     =   "Mp3 Filez"
      Filter          =   "Mp3 filez (*.mp3)|*.mp3"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   495
      Left            =   600
      TabIndex        =   12
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmId3.frx":0000
      Left            =   2040
      List            =   "frmId3.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   840
      MaxLength       =   30
      TabIndex        =   4
      Top             =   3240
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   840
      MaxLength       =   4
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   840
      MaxLength       =   30
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   840
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      MaxLength       =   30
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Genre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2580
      TabIndex        =   11
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   10
      Top             =   3000
      Width           =   885
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   9
      Top             =   2280
      Width           =   435
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Album"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1725
      TabIndex        =   8
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Artist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1770
      TabIndex        =   7
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   120
      Width           =   405
   End
End
Attribute VB_Name = "frmId3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.ShowOpen
GetId3 CommonDialog1.Filename           ' Get the filename
Text1 = RTrim(id3Info.Title)            ' since the fields in the type are
Text2 = RTrim(id3Info.Artist)                  ' fixed lenght, we use Rtrim to cut the
Text3 = RTrim(id3Info.Album)                   ' trailing bytes
Text4 = RTrim(id3Info.sYear)
Text5 = RTrim(id3Info.Comments)
Combo1.ListIndex = id3Info.Genre        ' fill in all the correct info.
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
        
id3Info.Title = Text1           ' just filling in the information into the type
id3Info.Artist = Text2
id3Info.Album = Text3
id3Info.sYear = Text4
id3Info.Comments = Text5
id3Info.Genre = Combo1.ListIndex
On Error GoTo ErrHandle             ' If the file is writeprotected
SaveId3 CommonDialog1.Filename, id3Info     ' Calling the Saveid3 function
Exit Sub


ErrHandle:
If Err.Number = 75 Then
MsgBox "File is Write Protected"
Else
MsgBox Err.Description
End If
End Sub

Private Sub Form_Load()
GenreArray = Split(sGenreMatrix, "|")   ' we fill the array with the Genre's
For i = LBound(GenreArray) To UBound(GenreArray)
Combo1.AddItem GenreArray(i)        ' now fill the Combobox with the array, and voila, the code you
                                    ' you recieve form the Genre part of the Type, represents the combobox Listindex =)
Next


End Sub
