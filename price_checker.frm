VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form price_checker 
   AutoRedraw      =   -1  'True
   Caption         =   "Stock taking"
   ClientHeight    =   9105
   ClientLeft      =   225
   ClientTop       =   120
   ClientWidth     =   15120
   FillStyle       =   3  'Vertical Line
   Icon            =   "price_checker.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   607
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1008
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "EXPORT"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12600
      MaskColor       =   &H0000FFFF&
      TabIndex        =   13
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00808080&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      MaskColor       =   &H0000FFFF&
      TabIndex        =   3
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   " "
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   18735
      Begin VB.CommandButton CMDBACKUP 
         Caption         =   "BACKUP"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2880
         TabIndex        =   37
         Top             =   8040
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000A&
         Height          =   495
         Left            =   10800
         TabIndex        =   33
         Top             =   7680
         Width           =   5025
         Begin VB.OptionButton optAll 
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   36
            Top             =   180
            Value           =   -1  'True
            Width           =   1245
         End
         Begin VB.OptionButton optCash 
            Caption         =   "Cash"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   35
            Top             =   180
            Width           =   1245
         End
         Begin VB.OptionButton OptCredit 
            Caption         =   "Credit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1980
            TabIndex        =   34
            Top             =   180
            Width           =   1245
         End
      End
      Begin VB.CommandButton Command_CLEAR 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   480
         TabIndex        =   32
         Top             =   8040
         Width           =   1695
      End
      Begin VB.OptionButton Option_CASH 
         BackColor       =   &H00FFC0FF&
         Caption         =   "CASH"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   31
         Top             =   7440
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option_GPAY 
         BackColor       =   &H00FFC0FF&
         Caption         =   "GPAY/PHONEPAY"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3600
         TabIndex        =   30
         Top             =   7440
         Width           =   2175
      End
      Begin VB.CheckBox ChecKSENDINGMSGDATA 
         Caption         =   "SENDING MSG DATA"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12600
         TabIndex        =   29
         Top             =   240
         Width           =   2415
      End
      Begin VB.CheckBox Check_whatsappmsg 
         Caption         =   "send whatsapp msg"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         TabIndex        =   28
         Top             =   7440
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Check_selectall 
         Caption         =   "SELECT ALL "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10920
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command_SENDMSG 
         Caption         =   "SEND MSG"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8880
         TabIndex        =   26
         Top             =   8040
         Width           =   1695
      End
      Begin VB.TextBox TXTID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   2400
         TabIndex        =   22
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtamount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   2400
         TabIndex        =   20
         Top             =   6600
         Width           =   8175
      End
      Begin VB.TextBox txtmobile 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   2400
         TabIndex        =   18
         Top             =   5760
         Width           =   8175
      End
      Begin VB.TextBox txtname 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   2400
         TabIndex        =   15
         Top             =   4920
         Width           =   8175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00808080&
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10800
         MaskColor       =   &H0000FFFF&
         TabIndex        =   12
         Top             =   8280
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   12690
         Top             =   7680
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00808080&
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6720
         MaskColor       =   &H0000FFFF&
         TabIndex        =   4
         Top             =   8040
         Width           =   1695
      End
      Begin VB.TextBox txtplace 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Left            =   2400
         TabIndex        =   1
         Top             =   1920
         Width           =   8175
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   6855
         Left            =   10800
         TabIndex        =   5
         Top             =   720
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   12091
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "id"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "place"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   4762
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "mobile"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "amount"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "MONEY MODE"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2235
         Left            =   2400
         TabIndex        =   24
         Top             =   2520
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   3942
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   4210688
         BackColor       =   65535
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "name"
            Object.Width           =   8819
         EndProperty
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label lbltot 
         BackColor       =   &H00808080&
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   14640
         TabIndex        =   25
         Top             =   8280
         Width           =   3015
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   23
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "AMOUNT"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   6600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "MOBILE "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "PLACE"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "19 -10 - 2004"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   9120
         TabIndex        =   11
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label LBLUSER 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..  "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   330
         Left            =   2040
         TabIndex        =   10
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME:  "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   330
         Left            =   240
         TabIndex        =   9
         Top             =   270
         Width           =   1650
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   330
         Left            =   8400
         TabIndex        =   8
         Top             =   270
         Width           =   645
      End
      Begin VB.Label lbldate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "19 -10 - 2004"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   330
         Left            =   6120
         TabIndex        =   7
         Top             =   270
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   330
         Left            =   5340
         TabIndex        =   6
         Top             =   270
         Width           =   525
      End
      Begin VB.Label lblcompany 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Y-Not Software MARRIAGE FUNCTION "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Top             =   720
         Width           =   8175
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "PLACE"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   2040
      Width           =   1935
   End
End
Attribute VB_Name = "price_checker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public YNOTWAurl As String
Public YNOTWAtokan As String
Public paymentMethod As String
Public EditMode As Boolean

Private Sub InitializeWASettings()
    ' Read from POS settings stored in Registry
    YNOTWAurl = GetSetting("POS", "Settings", "YNOTWAurl", "")
    YNOTWAtokan = GetSetting("POS", "Settings", "YNOTWAtokan", "")

    ' Validate
    If Trim(YNOTWAurl) = "" Or Trim(YNOTWAtokan) = "" Then
        MsgBox "YNOTWAurl or YNOTWAtokan is not configured! Please check POS settings.", vbCritical
        Exit Sub
    End If

    ' Ensure URL starts with http:// or https://
    If Left$(YNOTWAurl, 7) <> "http://" And Left$(YNOTWAurl, 8) <> "https://" Then
        MsgBox "YNOTWAurl must start with http:// or https://", vbCritical
        Exit Sub
    End If
End Sub

Private Sub Check_selectall_Click()
    Dim i As Long
    Dim allChecked As Boolean

    ' Determine whether the "Select All" checkbox is checked or unchecked
    allChecked = Check_selectall.Value ' True = ticked, False = unticked

    ' Loop through all items in the ListView
    For i = 1 To lvwItems.ListItems.Count
        lvwItems.ListItems(i).Checked = allChecked
    Next i
End Sub

Private Sub ChecKSENDINGMSGDATA_Click()
    Dim str As String
    Dim i As Long

    
    lvwItems.ListItems.clear

    If ChecKSENDINGMSGDATA.Value = 1 Then
       
        str = "SELECT * FROM marriagefn WHERE MsgSent = 0 ORDER BY id DESC"
        showListView lvwItems, str, 7

       
        For i = 1 To lvwItems.ListItems.Count
            lvwItems.ListItems(i).Checked = True
        Next i
    Else
       
        str = "SELECT * FROM marriagefn ORDER BY id DESC"
        showListView lvwItems, str, 7

        
        For i = 1 To lvwItems.ListItems.Count
            lvwItems.ListItems(i).Checked = False
        Next i
    End If
End Sub

Private Sub CMDBACKUP_Click()
frmNewBackup.Show
End Sub

Private Sub cmdok_Click()

Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset

If Validations Then Exit Sub

'-------------------------
' UPDATE MODE
'-------------------------
If EditMode = True Then

    With cmd
        .ActiveConnection = con
        .CommandText = "UPDATE marriagefn SET " & _
                       "PLACE='" & UCase(NzString(txtplace.Text)) & "'," & _
                       "Name='" & UCase(NzString(txtname.Text)) & "'," & _
                       "mobile='" & NzString(txtmobile.Text) & "'," & _
                       "AMOUNT=" & val(txtamount.Text) & "," & _
                       "Money='" & paymentMethod & "' " & _
                       "WHERE id=" & val(TXTID.Text)
        .CommandType = adCmdText
        .Execute
    End With

    MsgBox "Record Updated Successfully", vbInformation

    EditMode = False

Else

'-------------------------
' YOUR ORIGINAL INSERT CODE
'-------------------------
    rs.Open "select * from marriagefn", con, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!id = val(TXTID.Text)
    rs!PLACE = UCase(NzString(txtplace.Text))
    rs!Name = UCase(NzString(txtname.Text))
    rs!mobile = NzString(txtmobile.Text)
    rs!AMOUNT = val(txtamount.Text)

    ' Add the payment method to the record
    rs!Money = paymentMethod

    rs.Update
    rs.Close

End If

'-------------------------
' YOUR EXISTING CODE
'-------------------------
loaditems
ArabicInvPrint val(TXTID.Text)
billingTime = Format(Now, "hh:mm AM/PM")

If Check_whatsappmsg.Value = 1 Then
    sendwhatsapp
End If

clear
loadtot

End Sub
Public Sub loadtot()
Dim totamt As Double
totamt = 0
    For i = 1 To lvwItems.ListItems.Count
        totamt = totamt + val(Trim(lvwItems.ListItems(i).SubItems(4)))
      
    Next i
    
    lbltot.Caption = Format(totamt, Currencymask)
End Sub
Private Function Validations() As Boolean
On Error GoTo ErrHnd
    If (txtname.Text = "") Then
        MsgBox "Please Enter Name", vbInformation, "Message"
        Validations = True
    ElseIf (txtamount.Text = "") Then
        MsgBox "Please Enter Amount", vbInformation, "Message"
        Validations = True
    ElseIf (TXTID.Text = "") Then
        MsgBox "Please Enter id", vbInformation, "Message"
        Validations = True
    Else
        Validations = False
    End If
        
    Exit Function
ErrHnd:
'ErrorTrap Err.Number, Err.Description, "frmcards", "Validations"
End Function
Public Sub ArabicInvPrint(p_nInvNo As Long, Optional p_nFlag As Integer, Optional p_nDupBill As Integer)
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim rep As New INVBILL
Dim frm As New Form1

PrintTimes = PrintTimes + 1
    With cmd
        Set .ActiveConnection = con
        .CommandText = "INVBILL " & p_nInvNo
        .CommandType = adCmdText
        Set rs = .Execute
    If rs.EOF Then
        MsgBox "No records to print for the Inv  " & p_nInvNo, vbInformation, "Message"
        Exit Sub
    End If
        Set .ActiveConnection = Nothing
    End With
    rep.DiscardSavedData
    rep.database.SetDataSource rs
    rep.txtPicLogoPath.SetText App.Path & "\\InvLogo.JPG"
    ''rep.SelectPrinter "HP LaserJet M1530 MFP Series PCL 6", "HP LaserJet M1530 MFP Series PCL 6", "HPLAserJetM1536dnfMFP"
    ''rep.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.port
  '  rep.txtCashier.SetText "Cashier: " & Trim(lblcashier)
  '  If DefaultPrinter = 0 Then
        rep.SelectPrinter InvPrinterName, InvPrinterName, InvPrinterPort
  '  End If
  '  rep.Field3.Suppress = True
   ' rep.txtFooter1.SetText GFooter1
  '  rep.txtFooter2.SetText GFooter2
    'rep.txtFooter3.SetText GFooter3
  '  rep.txtAddress.SetText GAddress1

        rep.DisplayProgressDialog = False
        rep.PrintOut False, 1
   ' If ExtcnInvShow = 1 Then
'        frm.CRViewer1.ReportSource = rep
'        frm.CRViewer1.ViewReport
'        frm.Show
    'End If
    On Error GoTo 0
Exit Sub

ArabicInvPrint_Error:

MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ArabicInvPrint of Form frmSales"
End Sub



Private Sub loaditems()
    Dim str As String
    
    
    str = "select * from marriagefn order by id desc"
    showListView lvwItems, str, 7
    
End Sub
Public Function nextnumber(tablename As String, fieldname As String) As Long
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
    With cmd
        .ActiveConnection = con
        .CommandText = "select isnull(max(" & fieldname & "),0) +1  from " & tablename
        .CommandType = adCmdText
        Set rs = .Execute
        
    End With
    nextnumber = rs(0)
End Function
Public Sub clear()
txtplace.Text = ""
txtmobile.Text = ""
txtname.Text = ""
txtamount.Text = ""
txtplace.SetFocus
loaditems
SetListViewColor
 TXTID = nextnumber("marriagefn ", "id")
End Sub
Private Function NzString(val As Variant) As String
    If IsNull(val) Then
        NzString = ""
    Else
        NzString = CStr(val)
    End If
End Function
Private Function GetListViewSubItem(lv As ListView, iItem As Long, iSub As Long) As String
    ' Safely get a ListView subitem
    On Error Resume Next
    GetListViewSubItem = NzString(lv.ListItems(iItem).SubItems(iSub))
    If Err.Number <> 0 Then
        Err.clear
        GetListViewSubItem = ""
    End If
End Function

Private Sub Command_CLEAR_Click()
clear
End Sub

Private Sub Command_SENDMSG_Click()
    On Error GoTo ErrHandler

    Dim i As Long
    Dim http As Object
    Dim fullURL As String
    Dim msg As String
    Dim mobile As String
    Dim cmd As New ADODB.Command

    If lvwItems.ListItems.Count = 0 Then
        MsgBox "No Records Found", vbInformation
        Exit Sub
    End If
    
    ' Use current time if not provided
    If sendTime = "" Then sendTime = Format(Now, "hh:mm AM/PM")

    ' Loop through ListView items
    For i = 1 To lvwItems.ListItems.Count
        ' Only process checked items
        If lvwItems.ListItems(i).Checked = True Then
            mobile = Trim(GetListViewSubItem(lvwItems, i, 3)) ' Mobile
            Dim customerName As String, PLACE As String, AMOUNT As String, MsgSent As String
            customerName = GetListViewSubItem(lvwItems, i, 2) ' Name
            PLACE = GetListViewSubItem(lvwItems, i, 1)       ' Place
            AMOUNT = GetListViewSubItem(lvwItems, i, 4)      ' Amount
            MsgSent = GetListViewSubItem(lvwItems, i, 5)     ' MsgSent

            ' Skip if already sent
            If MsgSent = "1" Then GoTo SkipItem

            ' Validate mobile
            If mobile <> "" And IsNumeric(mobile) And Len(mobile) = 10 Then
                mobile = "+91" & mobile

                
msg = "*Dear " & customerName & ",*" & _
      "%0A%0A" & WHATSAPPMSGSEND1 & _
      "%0A%0AYour presence and blessings made our day unforgettable." & _
      "%0A%0A------------------------" & _
      "%0A*Venue: " & WHATSAPPVENUE1 & "* " & _
      "%0A*Date:* " & Format(Date, "dd/MM/yyyy") & _
      "%0A*Time:* " & sendTime & _
      "%0A------------------------" & _
      "%0A*Name:* " & customerName & _
      "%0A*Place:* " & PLACE & _
      "%0A*Amount:* " & AMOUNT & " Successfully Received" & _
      "%0A%0A_Technology Partner_" & _
      "%0A*YNOT-SOFTWARE SOLUTION'S*" & _
      "%0A*04567355864*"

                ' Build full API URL
                fullURL = YNOTWAurl & "token=" & YNOTWAtokan & "&to=" & mobile & "&body=" & msg

                ' Send HTTP request
                Set http = CreateObject("MSXML2.ServerXMLHTTP")
                http.Open "GET", fullURL, False
                http.Send
                Set http = Nothing

                ' Mark as sent in database
                With cmd
                    .ActiveConnection = con
                    .CommandText = "UPDATE marriagefn SET MsgSent = 1 WHERE id = " & lvwItems.ListItems(i).Text
                    .CommandType = adCmdText
                    .Execute
                End With
            End If
SkipItem:
        End If
    Next i

    
    loaditems
    Exit Sub

ErrHandler:
    MsgBox "Error sending message: " & Err.Description, vbCritical
End Sub
Private Function whatsappforcreditcust() As Boolean

On Error GoTo ErrHandler

    Dim url As String
    Dim postdata As String
    Dim strurl As String
    Dim mobileno As String
    Dim filename As String
    Dim WinHttpReq As Object
    Dim responseText As String
    
    url = YNOTWAurl
    
    '-------------------------------
    ' Format Mobile Number
    '-------------------------------
    mobileno = "+91" & Trim(wacustomermobile)
    
    '-------------------------------
    ' Create Message
    '-------------------------------
    filename = "THANK YOU FOR VISITING" & _
               " %0A Name : " & wacustomername & _
               " %0A Invoice No : " & lblInvoice & _
               " %0A Amount : " & txtTotAmt.Text & _
               " %0A Have a Great Day!"
    
    '-------------------------------
    ' Build URL
    '-------------------------------
    postdata = url & "token=" & YNOTWAtokan
    postdata = postdata & "&to=" & mobileno
    postdata = postdata & "&body=" & filename
    
    strurl = postdata
    
    '-------------------------------
    ' Send Request
    '-------------------------------
    Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
    
    With WinHttpReq
        .Open "GET", strurl, False
        .Send
        responseText = .responseText
    End With
    
    Set WinHttpReq = Nothing
    
    '-------------------------------
    ' Check Response
    '-------------------------------
    If InStr(1, responseText, "success", vbTextCompare) > 0 Then
        whatsappforcreditcust = True
    Else
        whatsappforcreditcust = False
        Debug.Print responseText
    End If
    
    Exit Function

ErrHandler:
    whatsappforcreditcust = False

End Function

Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Command2_Click()

    Dim cmd As New ADODB.Command
    Dim str As String
    Dim adminPwd As String
    Dim enteredPwd As String
    Dim i As Long
    
    On Error GoTo cmdeditupdate_Click_Error

    '-------------------------
    ' ADMIN PASSWORD CHECK
    '-------------------------
    adminPwd = "1234"   ' <<< SET YOUR ADMIN PASSWORD HERE
    
    enteredPwd = InputBox("Enter Admin Password to Delete:", "Admin Authorization")
    
    If enteredPwd = "" Then Exit Sub
    
    If enteredPwd <> adminPwd Then
        MsgBox "Wrong Password! Delete Not Allowed.", vbCritical, "Access Denied"
        Exit Sub
    End If

    '-------------------------
    ' BUILD ID LIST
    '-------------------------
    str = "("

    For i = 1 To lvwItems.ListItems.Count
        If lvwItems.ListItems(i).Checked = True Then
            ' Use ID (Text property) – NOT subitems
            str = str & CLng(lvwItems.ListItems(i).Text) & ","
        End If
    Next i

    If Right(str, 1) = "," Then
        str = Left(str, Len(str) - 1) & ")"
    Else
        MsgBox "Please Select Item to DELETE", vbInformation, "Message"
        Exit Sub
    End If

    '-------------------------
    ' DELETE USING ID
    '-------------------------
    With cmd
        .ActiveConnection = con
        .CommandText = "DELETE FROM marriagefn WHERE id IN " & str
        .CommandType = adCmdText
        .Execute
    End With

    MsgBox "Item Deleted Successfully", vbInformation, "Message"

    loaditems

    Exit Sub

cmdeditupdate_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Delete Procedure"

End Sub

Private Sub Command3_Click()
    Dim frm As New Form1
     Dim rs As New ADODB.Recordset
     Dim cmd As New ADODB.Command
    Dim str As String
       If optCash.Value = True Then
        PTYPESTR = " WHERE MONEY ='CASH' "
    ElseIf OptCredit.Value = True Then
        PTYPESTR = " WHERE MONEY ='GPAY' "
    Else
        PTYPESTR = ""
    End If
    
str = "alter  view marriagefnview as select * from marriagefn"
str = str & PTYPESTR

    With cmd
        .ActiveConnection = con
        .CommandText = str
        .CommandType = adCmdText
        .Execute
    End With
    
          CrystalReport1.ReportFileName = App.Path & "\marriage.rpt"
   
    CrystalReport1.DiscardSavedData = True
    CrystalReport1.SQLQuery = "select * from marriagefnview "
    Dim ir As Integer
    ir = CrystalReport1.LogOnServer("pdsodbc.dll", "Accounts", gCstrDatabase, "sa", pwd1)
    CrystalReport1.RetrieveSQLQuery
'If mangalamtiles = 0 Then
     
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.Action = 1
'Else

' CrystalReport1.Destination = crptToPrinter
' CrystalReport1.CopiesToPrinter = printcopies
' CrystalReport1.PrintReport
 'End If


End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()

Centerform Me
  paymentMethod = "CASH"
 TXTID = nextnumber("marriagefn ", "id")
loaditems
    PrintTimes = 0
    loadtot
InitializeWASettings
End Sub
Private Sub Form_Activate()
txtplace.SetFocus
    lbldate = Date
    lblTime = Time
End Sub


Public Sub LoadData()
    Dim sqls As String
    sqls = "select distinct PLACE from marriagefn where   PLACE like '" & txtplace.Text & "%' order by PLACE "
    showListView ListView1, sqls, 1
   
    
End Sub





Private Sub lblTime_Click()
lblTime = Time
End Sub

Private Sub lbltot_Click()

End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

txtplace.Text = Trim(ListView1.SelectedItem.Text)
End Sub

Private Sub ListView1_DblClick()

txtplace.Text = Trim(ListView1.SelectedItem.Text)
txtplace.SetFocus
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtplace.Text = (ListView1.SelectedItem.Text)
txtplace.SetFocus
End If
End Sub










Private Sub lvwItems_DblClick()


If lvwItems.SelectedItem Is Nothing Then Exit Sub

TXTID.Text = lvwItems.SelectedItem.Text
txtplace.Text = lvwItems.SelectedItem.SubItems(1)
txtname.Text = lvwItems.SelectedItem.SubItems(2)
txtmobile.Text = lvwItems.SelectedItem.SubItems(3)
txtamount.Text = lvwItems.SelectedItem.SubItems(4)

paymentMethod = Trim(lvwItems.SelectedItem.SubItems(6))

If UCase(paymentMethod) = "GPAY" Then
    Option_GPAY.Value = True
Else
    Option_CASH.Value = True
End If

EditMode = True



End Sub

' This sub is triggered when the CASH option is clicked

    


' This sub is triggered when the GPAY option is clicked
Private Sub Option_CASH_Click()
    paymentMethod = "CASH"
End Sub

Private Sub Option_GPAY_Click()
    paymentMethod = "GPAY"
End Sub






Private Sub Timer1_Timer()
'lblTime = Time
End Sub

Private Sub txtbarcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    frmsearch.callingform = "PRICECHK"
    frmsearch.Show
    frmsearch.txtname.SetFocus
    lvwItems.Refresh
ElseIf KeyCode = vbKeyDown Then
    lvwItems.SetFocus
End If
End Sub

Private Sub txtbarcode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtname.SetFocus
        
    End If
End Sub



Private Sub Txtstock_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdok.SetFocus
End If
End Sub


Private Sub txtamount_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then
       Option_GPAY.Value = True
 End If
End Sub

Private Sub txtamount_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
    cmdok.SetFocus
 Else
   KeyAscii = CheckForFloat(KeyAscii)
End If
  
End Sub


Private Sub txtmobile_KeyPress(KeyAscii As Integer)
    ' Allow only numbers (0-9) and Backspace
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> vbKeyBack And KeyAscii <> 13 Then
        KeyAscii = 0  ' Ignore the key
        Exit Sub
    End If

    ' Limit to 10 digits
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        If Len(txtmobile.Text) >= 10 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If

    ' If Enter key is pressed, move focus to txtamount
    If KeyAscii = 13 Then
        txtamount.SetFocus
    End If
End Sub


Private Sub txtname_KeyPress(KeyAscii As Integer)

 
    If KeyAscii = 13 Then
        If txtname.Text = "" Then Exit Sub
        txtmobile.SetFocus
    End If

End Sub


Private Sub txtplace_Change()
LoadData
End Sub

Private Sub txtplace_DblClick()

txtplace.Text = Trim(ListView1.SelectedItem.Text)
End Sub

Private Sub txtplace_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
       ListView1.SetFocus
   End If
End Sub

Private Sub txtplace_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtplace.Text = "" Then Exit Sub
txtname.SetFocus
End If
End Sub


Public Sub sendwhatsapp()


Dim i As Long
Dim http As Object
Dim fullURL As String
Dim msg As String
Dim mobile As String
Dim cmd As New ADODB.Command

mobile = Trim(txtmobile.Text)

Dim customerName As String
Dim PLACE As String
Dim AMOUNT As String
Dim MsgSent As String
Dim msgsentnew As String
customerName = txtname
PLACE = txtplace
AMOUNT = txtamount    ' Amount
' MsgSent = GetListViewSubItem(lvwItems, i, 5)     ' MsgSent column index (add this to ListView)
    If sendTime = "" Then
        sendTime = Format(Now, "hh:mm AM/PM")
    End If
    If mobile <> "" And IsNumeric(mobile) And Len(mobile) = 10 Then
        mobile = "+91" & mobile
' Build message
 ' msgsentnew = WHATSAPPMSGSEND
'WHATSAPPMSGSEND = GetSetting("POS", "Settings", "WHATSAPPMSGSEND", "")
msg = "*Dear " & customerName & ",*" & _
      "%0A%0A" & WHATSAPPMSGSEND1 & _
      "%0A%0AYour presence and blessings made our day unforgettable." & _
      "%0A%0A------------------------" & _
      "%0A*Venue: " & WHATSAPPMSGvenue1 & "* " & _
      "%0A*Date:* " & Format(Date, "dd/MM/yyyy") & _
      "%0A*Time:* " & sendTime & _
      "%0A------------------------" & _
      "%0A*Name:* " & customerName & _
      "%0A*Place:* " & PLACE & _
      "%0A*Amount:* " & AMOUNT & " Successfully Received" & _
      "%0A%0A_Technology Partner_" & _
      "%0A*YNOT-SOFTWARE SOLUTION'S*" & _
      "%0A*04567355864*"

        ' Build full API URL
        fullURL = YNOTWAurl & "token=" & YNOTWAtokan & "&to=" & mobile & "&body=" & msg
        ' Send HTTP request
        Set http = CreateObject("MSXML2.ServerXMLHTTP")
        http.Open "GET", fullURL, False
        http.Send
        Set http = Nothing

        ' Mark as sent in database
        With cmd
            .ActiveConnection = con
            .CommandText = "UPDATE marriagefn SET MsgSent = 1 WHERE id = " & val(TXTID)
            .CommandType = adCmdText
            .Execute
        End With
    End If
End Sub




Private Function SetListViewColor()
Dim nCount As Long
For nCount = 1 To lvwItems.ListItems.Count
    If lvwItems.ListItems.Item(nCount).SubItems(5) = 1 Then
        ColorListviewRow lvwItems, nCount, vbRed
   
    Else
        ColorListviewRow lvwItems, nCount, vbBlack
    End If
Next
End Function
Public Sub ColorListviewRow(lv As ListView, RowNbr As Long, RowColor As OLE_COLOR)
'***************************************************************************
'Purpose: Color a ListView Row
'Inputs : lv - The ListView
'         RowNbr - The index of the row to be colored
'         RowColor - The color to color it
'Outputs: None
'***************************************************************************

Dim itmX As ListItem
Dim lvSI As ListSubItem
Dim intIndex As Integer

On Error GoTo ErrorRoutine

Set itmX = lv.ListItems(RowNbr)
itmX.ForeColor = RowColor
For intIndex = 1 To lv.ColumnHeaders.Count - 2
Set lvSI = itmX.ListSubItems(intIndex)
lvSI.ForeColor = RowColor
Next

Set itmX = Nothing
Set lvSI = Nothing

Exit Sub

ErrorRoutine:

MsgBox Err.Description

End Sub

