VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MAIIN 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   ClientHeight    =   8910
   ClientLeft      =   1380
   ClientTop       =   1635
   ClientWidth     =   15105
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "MAIIN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   594
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1007
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox STATUS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   825
      Left            =   6360
      LinkTimeout     =   5000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Text            =   "MAIIN.frx":0D4A
      Top             =   7320
      Width           =   3975
   End
   Begin VB.TextBox LWait 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   525
      Left            =   6360
      LinkTimeout     =   5000
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Text            =   "MAIIN.frx":0D56
      Top             =   6840
      Width           =   3975
   End
   Begin VB.PictureBox PicMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   13080
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   38
      Top             =   6960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox PicSRC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   14400
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   37
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame frMosaic 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Mosaic Creation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   6855
      Left            =   10320
      TabIndex        =   18
      Top             =   0
      Width           =   3975
      Begin MSComctlLib.Slider scrollJPGQ 
         Height          =   255
         Left            =   2400
         TabIndex        =   56
         ToolTipText     =   "JPG Quality"
         Top             =   4800
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   450
         _Version        =   393216
         Min             =   80
         Max             =   100
         SelStart        =   95
         TickFrequency   =   5
         Value           =   95
      End
      Begin VB.CommandButton RE_BUILD 
         Caption         =   "Open and reBuild"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   22
         ToolTipText     =   "Open PreBuild Photomosaic and Rebuild"
         Top             =   5160
         Width           =   1455
      End
      Begin VB.OptionButton asJPG 
         BackColor       =   &H00808000&
         Caption         =   "JPG"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   2400
         TabIndex        =   53
         Top             =   4560
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton asBMP 
         BackColor       =   &H00808000&
         Caption         =   "BMP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   3120
         TabIndex        =   52
         Top             =   4560
         Width           =   680
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Draw and -OpenAndRebuild- Options"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   2895
         Left            =   0
         TabIndex        =   45
         Top             =   3600
         Width           =   2295
         Begin VB.CheckBox ADJ 
            BackColor       =   &H00C0C000&
            Caption         =   "Adjust Colors"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   840
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin MSComctlLib.Slider AdjBLEND 
            Height          =   330
            Left            =   120
            TabIndex        =   46
            ToolTipText     =   "% of Blend with Subject"
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            Max             =   20
            TickFrequency   =   10
         End
         Begin MSComctlLib.Slider adjPERC 
            Height          =   330
            Left            =   120
            TabIndex        =   49
            ToolTipText     =   "% of Adjust Color"
            Top             =   1080
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            Max             =   100
            SelStart        =   66
            TickFrequency   =   10
            Value           =   66
         End
         Begin MSComctlLib.Slider OUTS 
            Height          =   495
            Left            =   120
            TabIndex        =   50
            ToolTipText     =   "Output SIZE (useful for Open and Rebuild)"
            Top             =   1440
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            _Version        =   393216
            Min             =   5
            Max             =   50
            SelStart        =   10
            TickFrequency   =   5
            Value           =   10
         End
         Begin VB.Label OUTsLabel 
            BackColor       =   &H00C0C000&
            Caption         =   "output size"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   51
            Top             =   1920
            Width           =   2055
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C000&
            Caption         =   "BLEND"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.CheckBox chMIRROR 
         BackColor       =   &H00C0C000&
         Caption         =   "Mirrored Tiles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   36
         ToolTipText     =   "Use Mirrored Tiles?"
         Top             =   3240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton LoadSUB 
         Caption         =   "(3) Get Random PIC as subject"
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
         Left            =   2280
         TabIndex        =   33
         ToolTipText     =   " Get Random PIC as subject"
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox FmTYPE 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Photomosaic TYPE"
         Top             =   1200
         Width           =   2145
      End
      Begin VB.CheckBox NCENR 
         BackColor       =   &H00C0C000&
         Caption         =   "Same"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "Number of Columns = Number of Rows"
         Top             =   2280
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.TextBox tNC 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   28
         Text            =   "10"
         ToolTipText     =   "Number of Columns"
         Top             =   1560
         Width           =   570
      End
      Begin VB.TextBox tNR 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   27
         Text            =   "10"
         ToolTipText     =   "Number of Rows"
         Top             =   1950
         Width           =   570
      End
      Begin VB.TextBox Ncelle 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "7"
         ToolTipText     =   "Number of Tiles in Photomosaic"
         Top             =   2280
         Width           =   570
      End
      Begin VB.CheckBox FAST 
         BackColor       =   &H00C0C000&
         Caption         =   "FAST (less accurate)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   6480
         Width           =   2055
      End
      Begin VB.CommandButton LoadSubDlg 
         Caption         =   "(3) Load SUBJECT pic"
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
         Left            =   240
         TabIndex        =   23
         ToolTipText     =   "Load Subject Picture"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton CREA 
         Caption         =   "(4) C R E A TE  Photomosaic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2400
         TabIndex        =   21
         ToolTipText     =   "Start Creation Process"
         Top             =   5760
         Width           =   1455
      End
      Begin VB.CheckBox chAllowDuplicates 
         BackColor       =   &H00C0C000&
         Caption         =   "Allow Duplicates"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Allow or Deny Same photo to Appear multiple times."
         Top             =   2640
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox tMINDIST 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   19
         Text            =   "5"
         ToolTipText     =   "Minimal Distance Between Identical Tiles"
         Top             =   2880
         Width           =   570
      End
      Begin REALpm.MINI MINI 
         Height          =   3255
         Left            =   2400
         TabIndex        =   24
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   5741
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "R"
         Height          =   375
         Left            =   2040
         TabIndex        =   55
         ToolTipText     =   "Refresh Subject Pic"
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Save as"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   2400
         TabIndex        =   54
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         Caption         =   "PM TYPE:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label numC 
         BackStyle       =   0  'Transparent
         Caption         =   "N Cols:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label numR 
         BackStyle       =   0  'Transparent
         Caption         =   "N Rows:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1950
         Width           =   735
      End
   End
   Begin VB.Frame frCollections 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Collection(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   6855
      Left            =   6360
      TabIndex        =   8
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton chiNVERT 
         Caption         =   "Invert"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   43
         ToolTipText     =   "Invert Checked Collection(s)"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton UncheckALL 
         Caption         =   "UnCheck ALL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   42
         ToolTipText     =   "UnCheck All Collections"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdPF 
         Caption         =   "Click To Select Collection Folder From Root"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         TabIndex        =   35
         Top             =   5760
         Width           =   1215
      End
      Begin VB.CheckBox chHELP 
         BackColor       =   &H00808000&
         Caption         =   "Show Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   2640
         TabIndex        =   34
         Top             =   6480
         Width           =   1215
      End
      Begin VB.CommandButton CreateCollection 
         Caption         =   "(1) Create Collection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         TabIndex        =   17
         ToolTipText     =   "Select folder and Create New Collection"
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton LoadCollection 
         Caption         =   "(2) Load (checked) Collection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2640
         TabIndex        =   16
         ToolTipText     =   "Load Checked Collection(s) to use"
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton LoadCS 
         Caption         =   "LOAD SET"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   5760
         Width           =   735
      End
      Begin VB.CommandButton SaveCS 
         Caption         =   "SAVE SET"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   6240
         Width           =   735
      End
      Begin VB.CommandButton UpdateCollecions 
         Caption         =   "Update Collections"
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
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Update Checked Collection(s)"
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton checkALL 
         Caption         =   "Check ALL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         ToolTipText     =   "Check All Collections"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton SORTbutton 
         Caption         =   "Sort By Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Sort Collections"
         Top             =   240
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListCol 
         Height          =   3615
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Collection(s) List"
         Top             =   840
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   6376
         View            =   3
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   12632256
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Collection Name"
            Object.Width           =   4762
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Photos"
            Object.Width           =   1323
         EndProperty
      End
      Begin VB.CheckBox chShowFolder 
         BackColor       =   &H00808000&
         Caption         =   "Show Mosaic Folder"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   855
         Left            =   2640
         TabIndex        =   40
         ToolTipText     =   "Show Mosaic Folder after Creation"
         Top             =   5640
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Label LClabel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Collection Loaded"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   4560
         Width           =   3735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Speed Tests"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton putd 
      Caption         =   "conversione vecchie coll"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox RotaPIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   840
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   5
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox PBlabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   465
      Left            =   6240
      LinkTimeout     =   5000
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "t i m e s"
      Top             =   8160
      Width           =   4215
   End
   Begin VB.Timer pbTimer 
      Interval        =   1000
      Left            =   14280
      Top             =   3720
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   210
      Left            =   6360
      TabIndex        =   2
      Top             =   8640
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox picLoad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   1680
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox PicR 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1680
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   1
      Top             =   120
      Width           =   855
      Begin VB.Shape ShapeCenter 
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   240
         Shape           =   2  'Oval
         Top             =   120
         Width           =   165
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
End
Attribute VB_Name = "MAIIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author :Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact
'
'
'
'
'--------------------------------------------------------------------------------

Option Explicit

Private fs As New FileSystemObject

Private picFILE As File
Private FO As Folder
Private foS As Folders
Private CONTA As Long
Private ContaFOR As Long
Private i As Long
Private I2 As Long
Private Z As Long
Private ErrorI As Single

Private Real_I As Long

Private indexLOAD As Long

Private Pic2Read() As String
Private Pic2ReadDATE() As Double

Private WW As Long
Private HH As Long

Private Cr As Byte
Private cG As Byte
Private cb As Byte
Private ccc As Long

Private Coll() As tColl
Private PBmax As Long
Private PBvalue As Long

Private ColName As String

Private toSEE() As tPMzone

Private FM As tFM

Private SubjFileName As String

Private IsBuilding As Boolean

Private SearchPath As String

Private mySHELL As New shell32.Shell

Private TotalSourcePhotos As Long
Private ListaNOMI() As String
Private NumberOfCollections As Long


Private SWAPtoSee As tPMzone



'Private FASTColorDistanceR(-255 To 255, 0 To 255) As Single
'Private FASTColorDistanceG(-255 To 255) As Single
'Private FASTColorDistanceB(-255 To 255, 0 To 255) As Single
Private FASTColorDistanceR(0 To 255, 0 To 255) As Single
Private FASTColorDistanceG(0 To 255) As Single
Private FASTColorDistanceB(0 To 255, 0 To 255) As Single
'768
'1024
'768
'so every function returns between 0 and 255
'(((Then divide by 2 so function returns between 0 and 510)))

Private CollSORT As Boolean

Private MINDIST As Long



Dim H2 As Long
Dim W2 As Long
Dim P As Long
Dim Tempo As Single
Dim Tempo1 As Single
Dim Tempo2 As Single

Private Const ComputationalComplexity = 18

Private StatusFile As Integer

Private jpgQuality As Byte


Private Const PI = 3.1415926535898
Private Const PI2 = 6.2831853071796

Private cH1 As Long
Private cS1 As Long
Private cP1 As Long
Private cH2 As Long
Private cS2 As Long
Private cP2 As Long



Dim FX As clsFX


Private Sub CollSORT_Click()
'If CollSORT.Value = Checked Then
'    ListCol.SortKey = 0
'    ListCol.SortOrder = lvwAscending
'    ListCol.Sorted = True
'Else'
'
'    ListCol.SortKey = 1
'    ListCol.SortOrder = lvwDescending
'    ListCol.Sorted = True


'End If


End Sub



Private Sub ADJ_Click()
If ADJ.Value = Checked Then
    adjPERC.Visible = True
Else
    adjPERC.Visible = False
End If

End Sub

Private Sub ADJ_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 8

End Sub

Private Sub AdjBLEND_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 15
End Sub

Private Sub adjPERC_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 8

End Sub

Private Sub asBMP_Click()
scrollJPGQ.Visible = False

End Sub

Private Sub asJPG_Click()
scrollJPGQ.Visible = True
End Sub

Private Sub chAllowDuplicates_Click()
If chAllowDuplicates.Value = Checked Then
    tMINDIST.Visible = True
    
Else
    tMINDIST.Visible = False
End If


End Sub

Private Sub chAllowDuplicates_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 7


End Sub

Private Sub checkALL_Click()
For i = 1 To ListCol.ListItems.Count
    
    ListCol.ListItems.Item(i).Checked = True
    
Next

End Sub

Private Sub chiNVERT_Click()
For i = 1 To ListCol.ListItems.Count
    
    ListCol.ListItems.Item(i).Checked = Not (ListCol.ListItems.Item(i).Checked)
    
Next
End Sub

Private Sub cmdRefresh_Click()
If SubjFileName <> "" Then loads (SubjFileName)

End Sub

Private Sub OUTsLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 10
End Sub

Private Sub PicR_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ShapeCenter.left = x - ShapeCenter.Width \ 2
ShapeCenter.top = y - ShapeCenter.Height \ 2

End Sub



Private Sub scrollJPGQ_Change()
jpgQuality = scrollJPGQ
End Sub

Private Sub uncheckALL_Click()
For i = 1 To ListCol.ListItems.Count
    
    ListCol.ListItems.Item(i).Checked = False
    
    
Next

End Sub
Private Sub chMIRROR_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 13

End Sub

Private Sub cmdPF_Click()
If fs.FileExists(App.Path & "\lastopen.txt") Then fs.DeleteFile (App.Path & "\lastopen.txt")



End Sub

Private Sub Command1_Click()
Dim R1 As Byte
Dim R2 As Byte
Dim tmpMR As Integer
Dim LongDiv As Long
Dim IntDiv As Integer
Dim L1 As Long
Dim L2 As Long



Tempo1 = Timer
For i = 0 To 9100000
    LongDiv = Rnd * 100 + 1
    L2 = Int(Rnd * 1000)
    L1 = L2 / LongDiv
Next
Tempo1 = Timer - Tempo1


Tempo2 = Timer
For i = 0 To 9100000
    IntDiv = Rnd * 100 + 1
    L2 = Int(Rnd * 1000)
    L1 = L2 / IntDiv
Next
Tempo2 = Timer - Tempo2

'----------------------------------------------------------------
MsgBox "T1LONG: " & Tempo1 & "   T2INT: " & Tempo2


Tempo1 = Timer
For i = 0 To 9100000
    R1 = Rnd * 255
    R2 = Rnd * 255
    tmpMR = (R1 \ 1 + R2 \ 1) \ 2
Next
Tempo1 = Timer - Tempo1

Tempo2 = Timer
For i = 0 To 9100000
    R1 = Rnd * 255
    R2 = Rnd * 255
    tmpMR = R1 \ 2 + R2 \ 2
Next
Tempo2 = Timer - Tempo2

'----------------------------------------------------------------
MsgBox "T1: " & Tempo1 & "   T2: " & Tempo2


Tempo2 = Timer
For i = 0 To 9100000
    Z = Sqr((Rnd * 1530))
Next i
Tempo2 = Timer - Tempo2


Tempo1 = Timer
For i = 0 To 9100000
    Z = fastSQR(((Rnd * 1530)))
Next i
Tempo1 = Timer - Tempo1


MsgBox "normal: " & Tempo2 & "   fast: " & Tempo1


End Sub



Private Sub CREA_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 4

End Sub

Private Sub CreateCollection_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 1


End Sub

Private Sub FAST_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 9

End Sub

Private Sub FmTYPE_Click()
If (FmTYPE = "STND_MASK") Or (FmTYPE = "OVERLAP_MASK") Then MINI.Visible = True Else: MINI.Visible = False

If (FmTYPE <> "STND_MASK") And (FmTYPE <> "OVERLAP_MASK") Then FM.MaskName = "NOTHING"

If InStr(1, FmTYPE, "CIRCLED") Then
    ShapeCenter.Visible = True
    ShapeCenter.left = PicR.Width / 2 - ShapeCenter.Width / 2
    ShapeCenter.top = PicR.Height / 2 - ShapeCenter.Height / 2
    
Else
     ShapeCenter.Visible = False
End If



End Sub

Private Sub Form_Activate()

If fs.FolderExists(App.Path & "\MOSAIC") = False Then fs.CreateFolder (App.Path & "\MOSAIC")
If fs.FolderExists(App.Path & "\COLL") = False Then fs.CreateFolder (App.Path & "\COLL")

RefreshListCOl
SORTbutton_Click
SORTbutton_Click
'If ListCol.ListItems.Count > 0 Then ListCol.ListItems.Item(ListCol.ListItems.Count).Checked = True
'LoadCollection_Click
''''''''''''''''''LoadSUB_Click

MINI.GeneraMiniaturas App.Path & "\masks\"




End Sub

Private Sub Form_Initialize()
'WindowsXPC1.ColorScheme = XP_Blue
'WindowsXPC1.EndWinXPCSubClassing
'WindowsXPC1.InitSubClassing

InitFastSQR
InitFASTColorDistanceR
InitFASTColorDistanceG
InitFASTColorDistanceB
'Me.ScaleHeight = Me.ScaleWidth / 1.61803

End Sub

Private Sub Form_Load()




'RGBtoHSP 200, 255, 255, cH1, cS1, cP1
'MsgBox cH1 & " " & cS1 & " " & cP1



Set FX = New clsFX

jpgQuality = 95

scrollJPGQ = jpgQuality
 
Randomize Timer

PBmax = 1
PBvalue = 0

FmTYPE.AddItem "STANDARD"
FmTYPE.AddItem "STND_MASK"

FmTYPE.AddItem "OVERLAP"
FmTYPE.AddItem "OVERLAP_MASK"

'FmTYPE.AddItem "OVERLAP"
FmTYPE.AddItem "ART_1"
FmTYPE.AddItem "ART_brain"

FmTYPE.AddItem "CIRCLED_LR"
FmTYPE.AddItem "CIRCLED_UD"

FmTYPE.AddItem "ANG_OVERLAP_RND"
FmTYPE.AddItem "ANG_OVERLAP_COL"



FmTYPE.ListIndex = 0





Ncelle = Val(tNR) * Val(tNC)

ProcessPrioritySet , , ppidle 'ppbelownormal  So While is Computing You Can to Other

Load frmHELP

Randomize Timer


StatusFile = 100
Open App.Path & "\LOG.txt" For Output As StatusFile

Me.Caption = "Real PhotoMosaic " & App.Major & "." & App.Minor & "." & App.Revision
STATUS = "Status  (" & Me.Caption & ")"



End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
frmHELP.Visible = False
End Sub

Private Sub Form_Resize()
'PBlabel.left = MAIIN.ScaleWidth / 2 - PBlabel.Width / 2
'LWait.left = MAIIN.ScaleWidth / 2 - LWait.Width / 2
'STATUS.left = MAIIN.ScaleWidth / 2 - STATUS.Width / 2
'PB.Width = MAIIN.ScaleWidth - 20
'PB.Width = PBlabel.Width
'
'PB.left = MAIIN.ScaleWidth / 2 - PB.Width / 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
'  WindowsXPC1.EndWinXPCSubClassing
Unload frmHELP

Close StatusFile

End

End Sub

Sub Long2RGB(RGBcol As Long, ByRef R As Byte, ByRef G As Byte, ByRef B As Byte)

If RGBcol > 0 Then
    R = RGBcol And &HFF ' set red
    G = (RGBcol And &H100FF00) / &H100 ' set green
    B = (RGBcol And &HFF0000) / &H10000 ' set blue
Else
    R = 0: G = 0: B = 0
End If

End Sub

Sub WRITEbin(S As String, Optional FileNumber = 1)
Dim wbI As Long

For wbI = 1 To Len(S)
    Put #FileNumber, , CByte(Asc(Mid$(S, wbI, 1)))
Next wbI

End Sub

Function ReadBin(Optional FileNumber = 1) As String
Dim B As Byte
Dim S As String

'B = 0
''s = ""
S = vbNullString

'While B <> 124
'    Get #FileNumber, , B
'    S = S & Chr$(B)
'
'Wend
Do
    Get #FileNumber, , B
    S = S & Chr$(B)
Loop While B <> 124

ReadBin = left$(S, Len(S) - 1)

End Function

Private Sub CreateCollection_Click()
Dim ExitFor As Boolean
Dim XX As Integer
Dim YY As Integer
Dim SPATH As String
Dim spl() As String
Dim S As String

Dim Cancelled As Boolean
Dim PICFOLDER As String


Dim RealFolder As shell32.Folder2
Dim MsgRet
picLoad.Visible = False

Dim i As Long



If Dir(App.Path & "\lastopen.txt") <> vbNullString Then
    Open App.Path & "\lastopen.txt" For Input As 77
    Input #77, SPATH
    Close 77
End If


'GoTo SECONDWay
FirstWAY:
''' SHELLL'Browse
'Set RealFolder = mySHELL.BrowseForFolder(Me.hwnd, "Select Folder to Scan for Pictures", 1) ', App.Path & "\..\images")
Set RealFolder = mySHELL.BrowseForFolder(Me.hWnd, "Select Folder to Scan for Pictures", 1, SPATH)


If RealFolder Is Nothing Then Exit Sub

''''
spl = Split(CStr(RealFolder.Self.Path), "\")
If UBound(spl) > 0 Then
    SPATH = ""
    For i = 0 To UBound(spl) - 1
        SPATH = SPATH & spl(i) & "\"
    Next
End If


Open App.Path & "\lastopen.txt" For Output As 77
Print #77, SPATH
Close 77
'''''''''''''''''''''''''''
'FirstWAy
PICFOLDER = RealFolder.Self.Path


GoTo CONTINUA

'''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SECONDWay:
'Stop

'Load frmBrowse
'frmBrowse.SelFolder = SPATH
'frmBrowse.Show vbModal
'Cancelled = frmBrowse.Cancelled
'If Not Cancelled Then
''MsgBox frmBrowse.SelFolder
'PICFOLDER = frmBrowse.SelFolder
'PICFOLDER = left$(PICFOLDER, Len(PICFOLDER) - 1)
'Open App.Path & "\lastopen.txt" For Output As 77
'Print #77, PICFOLDER
'Close 77
'   '        PrevShowFiles = frmBrowse.ShowFiles
'   '        PrevShowNetwork = frmBrowse.ShowNetwork
'   '        PrevSubfolders = frmBrowse.AddSubfolders
'   '        PrevFolder = frmBrowse.SelFolder
'   '        Recurse = frmBrowse.AddSubfolders
'End If
'
'    Unload frmBrowse
''Stop
'
'If Cancelled Then Exit Sub
'
'
CONTINUA:
'

LWait.Visible = True


On Error GoTo ManageERROR

STATUS = "Looking... in " & PICFOLDER

ColName = DopoUltimaBarra(PICFOLDER)
'Stop

FindPICFiles (PICFOLDER)

'Stop



Open App.Path & "\COLL\" & ColName & ".COL" For Binary Access Write As 1

WRITEbin (CStr(CONTA) & "|" & vbCrLf)
WRITEbin PICFOLDER & "|" & vbCrLf

PBmax = CONTA
Tempo = Timer

'Stop

For i = 1 To CONTA
    PBvalue = i
    STATUS = Right$(Pic2Read(i), Len(Pic2Read(i)) \ 2) & " " & i & "/" & CONTA & " (" & CONTA - i & ")"
    
    
    Err.Clear
    
    picLoad = LoadPicture(Pic2Read(i))
    WW = picLoad.Width
    HH = picLoad.Height
    If WW > HH Then
        H2 = ComputationalComplexity '6 '12
        W2 = Round(WW * H2 / HH)
        Debug.Print WW & "x" & HH & " " & W2 & "x" & H2 & "   " & WW / HH & " " & W2 / H2
    Else
        W2 = ComputationalComplexity '6 '12
        H2 = Round(HH * W2 / WW)
        Debug.Print WW & "x" & HH & " " & W2 & "x" & H2 & "   " & WW / HH & " " & W2 / H2
    End If
    
    PicR.Width = W2
    PicR.Height = H2
    
    Call SetStretchBltMode(PicR.hdc, STRETCHMODE)
    Call StretchBlt(PicR.hdc, 0, 0, W2, H2, _
            picLoad.hdc, 0, 0, picLoad.Width - 1, picLoad.Height - 1, vbSrcCopy)
    
    PicR.Refresh
    '    Stop
    
    WRITEbin i & "|" & Replace(Pic2ReadDATE(i), ",", ".") & "|" & Pic2Read(i) & "|"
    WRITEbin WW & "|"
    WRITEbin HH & "|"
    WRITEbin W2 & "|"
    WRITEbin H2 & "|"
    Me.Cls
    For YY = 0 To PicR.Height - 1
        For XX = 0 To PicR.Width - 1
            
            ccc = GetPixel(PicR.hdc, XX, YY)
            Long2RGB ccc, Cr, cG, cb
            
            Put #1, , Cr
            Put #1, , cG
            Put #1, , cb
            
            Me.Line (5 + XX * 5, 5 + YY * 5)-(5 + (XX + 1) * 5, 5 + (YY + 1) * 5), RGB(Cr, cG, cb), BF
            
        Next XX
        WRITEbin (vbCrLf)
    Next YY
    
    'WRITEbin "|"
    'Me.Refresh
    DoEvents
Next i
Close 1

STATUS = "Collection " & ColName & " Created!"

RefreshListCOl
LWait.Visible = False
DoEvents
Exit Sub

'-------------------------------------------------------------------------
ManageERROR:
S = "Error: " & Err.Number & ".  Description: " & Err.Description & "." & vbCrLf & vbCrLf
S = S + "Picture " & Pic2Read(i) & vbCrLf & "is NOT VALID and can't be loaded!" & vbCrLf & vbCrLf
S = S + "The creation of collection " & Chr$(34) & ColName & Chr$(34) & " will be aborted." & vbCrLf & vbCrLf
S = S + "Do you want to DELETE  " & Pic2Read(i) & "  ?"

STATUS = S
MsgRet = MsgBox(S, vbYesNo, "Cant Load Picture!")
Close 1
STATUS = "Creation of  " & ColName & ".COL  Aborted."
Kill App.Path & "\COLL\" & ColName & ".COL"
If MsgRet = vbYes Then STATUS = STATUS & vbCrLf & Pic2Read(i) & " DELETED!": Kill Pic2Read(i)

RefreshListCOl
LWait.Visible = False
Err.Clear
'--------------------------------------------------------------------------

End Sub


Sub FindPICFiles(StartFolder)
Dim ExitFor As Boolean

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''' FIND PICTURES FILES
STATUS = "Looking folder: " & StartFolder
DoEvents

'Set FO = FS.GetFolder(App.Path & "\..\images")
Set FO = fs.GetFolder(StartFolder)
Set foS = FO.SubFolders

CONTA = 0
For Each picFILE In FO.Files
    'Or LCase(Right$(CStr(picFILE), 3)) = "bmp"
    If LCase(Right$(CStr(picFILE), 3)) = "jpg" Then
        CONTA = CONTA + 1
        ReDim Preserve Pic2Read(CONTA)
        ReDim Preserve Pic2ReadDATE(CONTA)
        Pic2Read(CONTA) = picFILE
        Pic2ReadDATE(CONTA) = CDbl(picFILE.DateLastModified)
        If CONTA >= 20000 Then ExitFor = True: Exit For
    End If
    DoEvents
Next
For Each FO In foS
    STATUS = FO & " " & CONTA
    DoEvents
    For Each picFILE In FO.Files
        'Or LCase(Right$(CStr(picFILE), 3)) = "bmp"
        If LCase(Right$(CStr(picFILE), 3)) = "jpg" Then
            CONTA = CONTA + 1
            ReDim Preserve Pic2Read(CONTA)
            ReDim Preserve Pic2ReadDATE(CONTA)
            Pic2Read(CONTA) = picFILE
            Pic2ReadDATE(CONTA) = CDbl(picFILE.DateLastModified)
            If CONTA >= 20000 Then ExitFor = True: Exit For
        End If
        DoEvents
    Next
    If ExitFor Then Exit For
Next
End Sub

Private Sub frCollections_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
frmHELP.Visible = False
End Sub

Private Sub frMosaic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
frmHELP.Visible = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 14

End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 15
End Sub

Private Sub ListCol_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 6

End Sub

Private Sub LoadCollection_Click()

Dim bR As Byte
Dim bG As Byte
Dim BB As Byte
Dim aCapo As Byte
Dim CC As Integer
Dim DateOrName
Dim GlobI As Long
Dim x As Long
Dim y As Long

Dim BYTES() As Byte

LWait.Visible = True

NumberOfCollections = 0
TotalSourcePhotos = 0
For i = 1 To ListCol.ListItems.Count
    If ListCol.ListItems.Item(i).Checked Then
        NumberOfCollections = NumberOfCollections + 1
        ReDim Preserve Coll(NumberOfCollections)
        Coll(NumberOfCollections).NAME = ListCol.ListItems.Item(i)
        '    MsgBox ListCol.ListItems.Item(I)
        
    End If
Next i

For CC = 1 To NumberOfCollections
    
    Open App.Path & "\Coll\" & Coll(CC).NAME For Binary Access Read As 1
    CONTA = CInt(ReadBin)
    'CONTA = readbin \ 1
    
    Get #1, , aCapo
    Get #1, , aCapo
    
    Coll(CC).STARTdir = ReadBin
    
    ContaFOR = CONTA
    '    Stop
    
    If chMIRROR.Value = Checked Then CONTA = CONTA * 2
    
    Coll(CC).NofPhotos = CONTA
    TotalSourcePhotos = TotalSourcePhotos + CONTA
    
    ReDim Coll(CC).Z(CONTA)
    'Stop
    PBvalue = 0
    PBmax = ContaFOR
    Tempo = Timer
    
    STATUS = "Loading Collection " & Coll(CC).NAME & "(" & CONTA & " photos)   " & CC & " of " & NumberOfCollections
    
    For i = 1 To ContaFOR 'CONTA
        
        PBvalue = i
        'STATUS = I & "/" & CONTA & " " & CONTA - I
        
        ReadBin
        
        DateOrName = ReadBin
        '        Stop
        
        '      Stop
        
        If Val(DateOrName) = 0 Then
            Coll(CC).Z(i).filename = DateOrName
        Else
            '    Stop
            ' DateOrName = Replace(DateOrName, ",", ".", , , vbBinaryCompare)
            
            Coll(CC).Z(i).FileDate = Val(DateOrName) '(Val(DateOrName))
            Coll(CC).Z(i).filename = ReadBin
        End If
        
        Coll(CC).Z(i).oW = CInt(ReadBin)
        Coll(CC).Z(i).oH = CInt(ReadBin)
        Coll(CC).Z(i).WW = CInt(ReadBin)
        Coll(CC).Z(i).HH = CInt(ReadBin)
        
        'seems useless If Coll(CC).Z(I).oW > Coll(CC).Z(I).oH Then O = ORR Else: O = VER
        
        ReDim Preserve Coll(CC).Z(i).R(Coll(CC).Z(i).WW, Coll(CC).Z(i).HH)
        ReDim Preserve Coll(CC).Z(i).G(Coll(CC).Z(i).WW, Coll(CC).Z(i).HH)
        ReDim Preserve Coll(CC).Z(i).B(Coll(CC).Z(i).WW, Coll(CC).Z(i).HH)
        
        '        Stop
        
        
        '-------------------------------------------------------
        '-------------------------------------------------------
        ' For y = 1 To Coll(CC).Z(i).HH
        '     For x = 1 To Coll(CC).Z(i).WW
        '     '    Get #1, , bR
        '     '    Get #1, , bG
        '     '    Get #1, , BB
        '     '    Coll(CC).Z(i).R(x, y) = bR
        '     '    Coll(CC).Z(i).G(x, y) = bG
        '     '    Coll(CC).Z(i).b(x, y) = BB
        '         Coll(CC).Z(i).G(x, y) = BYTES(x + 1)
        '         Coll(CC).Z(i).b(x, y) = BYTES(x + 2)
        '     Next x
        '     Get #1, , aCapo
        '     Get #1, , aCapo
        ' Next y
        '-------------------------------------------------------
        '-------------------------------------------------------
        ReDim BYTES(0 To Coll(CC).Z(i).WW * 3 + 2 - 1, 1 To Coll(CC).Z(i).HH)
        
        Get #1, , BYTES
        For y = 1 To Coll(CC).Z(i).HH
            For x = 1 To Coll(CC).Z(i).WW
                Coll(CC).Z(i).R(x, y) = BYTES((x - 1) * 3, y) ' + (y - 1) * Coll(CC).Z(i).HH)
                Coll(CC).Z(i).G(x, y) = BYTES((x - 1) * 3 + 1, y) '+ (y - 1) * Coll(CC).Z(i).HH)
                Coll(CC).Z(i).B(x, y) = BYTES((x - 1) * 3 + 2, y) '+ (y - 1) * Coll(CC).Z(i).HH)
            Next x
        Next y
        '------------------------------------------------------
        '-------------------------------------------------------
        
        
        Coll(CC).Z(i).Mirrored = False
        
        If chMIRROR.Value = Checked Then
            
            I2 = i + ContaFOR
            Coll(CC).Z(I2) = Coll(CC).Z(i)
            Coll(CC).Z(I2).Mirrored = True
            ReDim Preserve Coll(CC).Z(I2).R(Coll(CC).Z(I2).WW, Coll(CC).Z(I2).HH)
            ReDim Preserve Coll(CC).Z(I2).G(Coll(CC).Z(I2).WW, Coll(CC).Z(I2).HH)
            ReDim Preserve Coll(CC).Z(I2).B(Coll(CC).Z(I2).WW, Coll(CC).Z(I2).HH)
            
            For y = 1 To Coll(CC).Z(i).HH
                For x = 1 To Coll(CC).Z(i).WW
                    Coll(CC).Z(I2).R(x, y) = Coll(CC).Z(i).R(Coll(CC).Z(i).WW - x + 1, y)
                    Coll(CC).Z(I2).G(x, y) = Coll(CC).Z(i).G(Coll(CC).Z(i).WW - x + 1, y)
                    Coll(CC).Z(I2).B(x, y) = Coll(CC).Z(i).B(Coll(CC).Z(i).WW - x + 1, y)
                Next x
            Next y
        End If
        
        DoEvents
        
    Next i
    
    GlobI = GlobI + ContaFOR
    
    Close 1
    STATUS = "Collection:" & ColName & " Loaded. " & CONTA & " photos " & IIf(chMIRROR.Value = Checked, "(With Mirrored)", "") & "  TOTAL = * " & TotalSourcePhotos & " *"
    ' Version " & App.Major & "." & App.Minor & "." & App.Revision
    
Next CC

ReDim GlobalMirrored(TotalSourcePhotos)
ReDim ListaNOMI(TotalSourcePhotos)
GlobI = 0
For CC = 1 To NumberOfCollections
    For P = 1 To Coll(CC).NofPhotos
        GlobI = GlobI + 1
        '    Stop
        
        GlobalMirrored(GlobI) = Coll(CC).Z(P).Mirrored
        'ListaNOMI(GlobI) = Coll(CC).Z(p).FileName
        
    Next
Next

''''''''''''''''''''''
''''''''''''''''''''''
''''''''''''''''''''''

LClabel = " Source Photos: " & TotalSourcePhotos & IIf(chMIRROR.Value = Checked, "(M)", "")

LWait.Visible = False

End Sub

Private Sub LoadCollection_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 2

End Sub

Private Sub LoadCS_Click()
Dim sOpen As SelectedFile
Dim cS As String
Dim cL As String
Dim ind As Integer
Dim NotFound As String
Dim Yfound As String ''

Dim F As Boolean '

If Right$(FileDialog.sFile, 3) <> "TXT" Then FileDialog.sFile = ""
FileDialog.sFile = ""
FileDialog.sFilter = "Collection SET" & Chr$(0) & "*.txt"
FileDialog.sInitDir = App.Path & "\COLL\"
'FileDialog.Action = 1

sOpen = ShowOpen(Me.hWnd)

If sOpen.nFilesSelected = 0 Then Exit Sub

cS = sOpen.sLastDirectory & sOpen.sFiles(1)
'CS = FileDialog.FileName
If Right$(cS, 3) <> "txt" Then Exit Sub
For ind = 1 To ListCol.ListItems.Count
    ListCol.ListItems.Item(ind).Checked = False
Next
NotFound = ""
Yfound = ""

Open cS For Input As 3
Do
    Input #3, cL
    
    F = False
    
    For ind = 1 To ListCol.ListItems.Count
        If ListCol.ListItems.Item(ind) = cL Then
            ListCol.ListItems.Item(ind).Checked = True
            Yfound = Yfound & cL & vbCrLf
            F = True
        End If
    Next
    
    If F = False Then NotFound = NotFound & cL & vbCrLf
    
    
Loop While Not (EOF(3))
Close 3

If Len(NotFound) <> 0 Then NotFound = vbCrLf & "List of not found collection:" & vbCrLf & NotFound

MsgBox "Collection SET:  " & Chr$(34) & DopoUltimaBarra(cS) & Chr$(34) & " " & vbCrLf & _
        vbCrLf & "List: " & vbCrLf & Yfound & vbCrLf & NotFound & vbCrLf & vbCrLf & "NOW START LOAD COLLECION(s)", vbInformation

LoadCollection_Click


End Sub

Private Sub LoadCS_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 12


End Sub

Private Sub LoadSUB_Click()
Dim CC As Integer

If TotalSourcePhotos = 0 Then MsgBox "Go to Point(2)", vbExclamation, "No Collection Loaded!": Exit Sub
CC = Int(Rnd * UBound(Coll) + 1)
SubjFileName = Coll(CC).Z(Int(Rnd * Coll(CC).NofPhotos + 1)).filename
STATUS = "Subject: " & SubjFileName

loads (SubjFileName)
OUTS_Scroll

If InStr(1, FmTYPE, "CIRCLED") Then
    ShapeCenter.Visible = True
    ShapeCenter.left = PicR.Width / 2 - ShapeCenter.Width / 2
    ShapeCenter.top = PicR.Height / 2 - ShapeCenter.Height / 2
End If

End Sub
Sub loads(filename As String)
Me.Cls


picLoad = LoadPicture(filename)
picLoad.Refresh

FM.TotalW = picLoad.Width
FM.TotalH = picLoad.Height


WW = picLoad.Width
HH = picLoad.Height
If WW > HH Then
    H2 = 140 '6 '12
    W2 = Round(WW * H2 / HH)
    Debug.Print WW & "x" & HH & " " & W2 & "x" & H2 & "   " & WW / HH & " " & W2 / H2
Else
    W2 = 140 '6 '12
    H2 = Round(HH * W2 / WW)
    Debug.Print WW & "x" & HH & " " & W2 & "x" & H2 & "   " & WW / HH & " " & W2 / H2
End If

PicR.Width = W2
PicR.Height = H2

Call SetStretchBltMode(PicR.hdc, STRETCHMODE)
Call StretchBlt(PicR.hdc, 0, 0, W2, H2, _
        picLoad.hdc, 0, 0, picLoad.Width - 1, picLoad.Height - 1, vbSrcCopy)
PicR.Refresh




End Sub

Private Sub LoadSUB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 16

End Sub

Private Sub LoadSubDlg_Click()
Dim sOpen As SelectedFile
' See Standard CommonDialog Flags for all options

FileDialog.sFilter = "Picture " & Chr$(0) & "*.jpg;*.bmp;*.jpeg"

' See Standard CommonDialog Flags for all options
'FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
FileDialog.sDlgTitle = "Load Select Subject Picture"
FileDialog.sInitDir = App.Path & "\"

sOpen = ShowOpen(Me.hWnd)

'If Right$(filedialog.sFileTitle, 3) <> "JPG" Then filedialog.sFileTitle = ""
'filedialog.Filter = "BMP AND JPG|*.BMP;*.JPG;*.JPEG|bmp|*.BMP|JPG|*.JPG"
'filedialog.Action = 1

If sOpen.nFilesSelected = 0 Then Exit Sub

SubjFileName = sOpen.sLastDirectory & sOpen.sFiles(1)

'MsgBox filedialog.FileTitle

STATUS = "Subject: " & SubjFileName

loads (SubjFileName)
OUTS_Scroll

If InStr(1, FmTYPE, "CIRCLED") Then
    ShapeCenter.Visible = True
    ShapeCenter.left = PicR.Width / 2 - ShapeCenter.Width / 2
    ShapeCenter.top = PicR.Height / 2 - ShapeCenter.Height / 2
End If


End Sub

Private Sub LoadSubDlg_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 3

End Sub

Private Sub MINI_Click(ThePicture As stdole.Picture, S As String)
'MsgBox ThePicture.Handle

picLoad = ThePicture
Call SetStretchBltMode(PicMask.hdc, STRETCHMODE)
Call StretchBlt(PicMask.hdc, 0, 0, PicMask.Width, PicMask.Height, _
        picLoad.hdc, 0, 0, picLoad.Width - 1, picLoad.Height - 1, vbSrcCopy)


PicMask.Refresh
FM.MaskName = S

End Sub

Private Sub NCENR_Click()
If NCENR.Value = Checked Then
    tNR = tNC
    tNR.Enabled = False
Else
    tNR.Enabled = True
End If

End Sub

Private Sub OUTS_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 10


End Sub

Private Sub OUTS_Scroll()
Dim S As String
Dim W As Single
Dim H As Single
W = FM.TotalW * OUTS / 10
H = FM.TotalH * OUTS / 10

S = "Image " & PixelsToCentimeter(W) & "x" & PixelsToCentimeter(H) & " cm" & vbCrLf
S = S & "___" & W & "x" & H & " pix" & vbCrLf
S = S & "Tiles " & PixelsToCentimeter((W) \ Val(tNC)) & "x" & PixelsToCentimeter((H) \ Val(tNR)) & " cm" & vbCrLf
S = S & "___" & (W) \ Val(tNC) & "x" & (H) \ Val(tNR) & " pix"

OUTsLabel = S

End Sub

Private Sub pbTimer_Timer()
On Error Resume Next

If IsBuilding Then PBvalue = (CLng(Z - 1) * TotalSourcePhotos + Real_I) / 1000
PB.Max = PBmax
PB.Value = PBvalue
'Stop

'PBlabel.Caption = Format(PBvalue / PBmax, "0%") & "  " & Format((Timer - TEMPO) / 86400, "HH:MM:SS") & "  Remain: " & Format((((Timer - TEMPO) / 86400) / (PBvalue + 1)) * (PBmax - PBvalue), "HH:MM:SS")
PBlabel.Text = Format(PBvalue / PBmax, "0.0%") & "  " & Format((Timer - Tempo) / 86400, "HH:MM:SS") & "  Remain: " & Format((((Timer - Tempo) / 86400) / (PBvalue + 1)) * (PBmax - PBvalue), "HH:MM:SS")
Me.Caption = STATUS
'Me.BackColor = IIf(lWAIT.Visible, &H404000, &H808000)

End Sub
Private Sub CREA_Click()

loads (SubjFileName)
CREATE2


End Sub

Private Sub putd_Click()
'
'''''' usato per mettere file date ---- non usare piu
'
'Dim iC As Long
'Dim iP As Long
'
'LoadCollection_Click
'
'For iC = 1 To NumberOfCollections
'
'    STATUS = Coll(iC).NAME & "  " & iC & " of " & NumberOfCollections
'    DoEvents
'    For iP = 1 To Coll(iC).NofPhotos
'
'        Set picFILE = FS.GetFile(Coll(iC).Z(iP).FileName)
'        Coll(iC).Z(iP).FileDate = picFILE.DateLastModified
'    Next
'
'
'    SAVECOLLECTION Coll(iC).NAME, iC '
'
'Next iC
'STATUS = "CONVERSION DONE"

End Sub

Private Sub RE_BUILD_Click()
Dim sOpen As SelectedFile
'filedialog.Filter = "Load Photomosaic|*.TXT"
'filedialog.InitDir = App.Path & "\MOSAIC\"
'filedialog.Action = 1
FileDialog.sFile = ""
FileDialog.sDlgTitle = "Select Photomosaic to ReBuild"
FileDialog.sFilter = "Photomosaic (*.TXT)" & Chr$(0) & "*.txt"
FileDialog.sInitDir = App.Path & "\MOSAIC\"

sOpen = ShowOpen(Me.hWnd)

If sOpen.nFilesSelected = 0 Then Exit Sub

'If (LoadFM(filedialog.FileName)) Then lWAIT.Visible = True: Ricostruz
If (LoadFM(sOpen.sLastDirectory & sOpen.sFiles(1))) Then LWait.Visible = True: Ricostruz

End Sub

Private Sub RE_BUILD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 5
End Sub

Private Sub SaveCS_Click()
Dim n As String
Dim ind As Integer
Dim numC As Long '
Dim SetName As String

n = InputBox("Collecion SET name", "Save Collection set", "Collection Set Name")

For ind = 1 To ListCol.ListItems.Count
    If ListCol.ListItems.Item(ind).Checked = True Then
        Open App.Path & "\coll\" & ListCol.ListItems.Item(ind) For Binary Access Read As 3
        numC = numC + ReadBin(3)
        Close 3
    End If
Next

'n = n & "_" & numC
SetName = "SET_" & n & ".txt"
Open App.Path & "\Coll\" & SetName For Output As 3

For ind = 1 To ListCol.ListItems.Count
    
    If ListCol.ListItems.Item(ind).Checked = True Then
        Print #3, ListCol.ListItems.Item(ind)
    End If
Next
Close 3

MsgBox "Collection SET   " & Chr$(34) & SetName & Chr$(34) & "   saved!", vbInformation


End Sub

Private Sub SaveCS_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 12

End Sub

Private Sub SORTbutton_Click()
If CollSORT Then
    ListCol.SortKey = 1
    ListCol.SortOrder = lvwDescending
    ListCol.Sorted = True
    CollSORT = False
    SORTbutton.Caption = "Sort by NAME"
Else
    
    ListCol.SortKey = 0
    ListCol.SortOrder = lvwAscending
    ListCol.Sorted = True
    CollSORT = True
    
    SORTbutton.Caption = "Sort by NUMBER"
End If
End Sub

Private Sub STATUS_Change()
Print #StatusFile, left$(Time, 5) & vbTab & STATUS

End Sub

Private Sub tMINDIST_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 7


End Sub

Private Sub tNC_Change()
If Val(tNC) < 1 Then tNC.Text = "1"
Ncelle = Val(tNR) * Val(tNC)
If NCENR.Value = Checked Then tNR = tNC
OUTS_Scroll
End Sub

Private Sub tNR_Change()
If Val(tNR) < 1 Then tNR.Text = "1"
Ncelle = Val(tNR) * Val(tNC)
OUTS_Scroll
End Sub


Sub Computazione()

Dim DR As Long
Dim DG As Long
Dim DB As Long
Dim DRT As Long
Dim DGT As Long
Dim DBT As Long
Dim agR As Long
Dim agB As Long
Dim agG As Long
Dim BestFIT As Single

Dim BESTindex As Long
Dim BEST As Single
Dim I_pic As Long
Dim ZonaBest As Long

Dim MyStep As Integer

Dim NofPixels As Long ' ??????? ' integer

Dim tmpFIT As Single
Dim tmpMR As Integer 'Single

Dim dx As Integer
Dim dy As Integer

Dim x As Long
Dim y As Long

Dim CC As Long
''''''''''''''''''''''''''''
Dim K As Single
Dim sX As Single
Dim sY As Single
Dim Y2 As Single
Dim X2 As Single
Dim SCX As Long
Dim Uguali As Boolean
Dim Z1 As Long
Dim Z2 As Long
Dim cH As Long
Dim ToG As Long
Dim ZZ As Long
Dim II As Long

Dim SameCounter As Long

Dim Msg As String



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''' computazione c o m p u t a z i o n e ''''''''''''''''''

PBmax = (CLng(FM.NZones) * TotalSourcePhotos) / 1000
PBvalue = 0
Real_I = 0
Tempo = Timer
IsBuilding = True

If FAST.Value = Checked Then MyStep = 2 Else: MyStep = 1



For Z = 1 To FM.NZones
    
    NofPixels = FM.toSEE(Z).WZone_W * FM.toSEE(Z).WZone_H
    If NofPixels = 0 Then Stop
    
    'PBvalue = Z
    BestFIT = 1E+23
    'Me.Cls
    For y = 1 To FM.toSEE(Z).WZone_H 'FM.WZone_H
        For x = 1 To FM.toSEE(Z).WZone_W 'FM.WZone_W
            MAIIN.Line (5 + x * 5, 5 + y * 5)-(5 + (x + 1) * 5, 5 + (y + 1) * 5), RGB(FM.toSEE(Z).R(x - 1, y - 1), FM.toSEE(Z).G(x - 1, y - 1), FM.toSEE(Z).B(x - 1, y - 1)), BF
        Next x
    Next y
    '    Stop
    
    DoEvents
    
    
    Real_I = 0
    
    With FM.toSEE(Z)
        
        For CC = 1 To UBound(Coll)
            
            For i = 1 To Coll(CC).NofPhotos
                
                Real_I = Real_I + 1
                
                ListaNOMI(Real_I) = Coll(CC).Z(i).filename 'moved to loadcoll
                
                
                DoEvents
                '
                'Debug.Print "2see " & FM.WZone_W & "/" & FM.WZone_H & "=" & FM.WZone_W / FM.WZone_H & "     coll " & Coll.Z(I).WW & "/" & Coll.Z(I).HH & "=" & Coll.Z(I).WW / Coll.Z(I).HH
                
                sX = 0
                sY = 0
                'If FM.WZone_W > FM.WZone_H Then
                If (.WZone_W / .WZone_H) > (Coll(CC).Z(i).WW / Coll(CC).Z(i).HH) Then
                    K = Coll(CC).Z(i).WW / .WZone_W
                    sY = (Coll(CC).Z(i).HH - .WZone_H * K) / 2
                    'OK
                    'Stop
                Else
                    K = Coll(CC).Z(i).HH / .WZone_H
                    sX = ((-.WZone_W * K + Coll(CC).Z(i).WW) / 2)
                    'Stop
                    
                End If
                'Else
                'End If
                '
                '
                DRT = 0
                DGT = 0
                DBT = 0
                DR = 0
                DG = 0
                DB = 0
                agR = 0
                agG = 0
                agB = 0
                tmpFIT = 0
                tmpMR = 0
                For y = 0 To .WZone_H - 1 Step MyStep
                    
                    Y2 = (y + 1) * K + sY
                    
                    For x = 0 To .WZone_W - 1 Step MyStep
                        
                        X2 = (x + 1) * K + sX
                        
                        '        Me.Line (5 + X * 5, 150 + Y * 5)-(5 + (X + 1) * 5, 150 + (Y + 1) * 5), RGB(Coll.Z(I).R(X2, Y2), Coll.Z(I).G(X2, Y2), Coll.Z(I).B(X2, Y2)), BF
                        
                        'DR = (CInt(.R(x, y) - CInt(Coll(CC).Z(i).R(X2, Y2))))
                        'DG = (CInt(.G(x, y) - CInt(Coll(CC).Z(i).G(X2, Y2))))
                        'DB = (CInt(.B(x, y) - CInt(Coll(CC).Z(i).B(X2, Y2))))
                        
                        
                        'HSP
                        'RGBtoHSP .R(x, y), .G(x, y), .B(x, y), cH1, cS1, cP1
                        'RGBtoHSP Coll(CC).Z(i).R(X2, Y2), Coll(CC).Z(i).G(X2, Y2), Coll(CC).Z(i).B(X2, Y2), cH2, cS2, cP2
                        
                        
                        
                        DR = .R(x, y) \ 1 - Coll(CC).Z(i).R(X2, Y2) \ 1
                        DG = .G(x, y) \ 1 - Coll(CC).Z(i).G(X2, Y2) \ 1
                        DB = .B(x, y) \ 1 - Coll(CC).Z(i).B(X2, Y2) \ 1
                        
                        
                        agR = agR + DR
                        agG = agG + DG
                        agB = agB + DB
                        
                        DR = Abs(DR)
                        DG = Abs(DG)
                        DB = Abs(DB)
                        
                        DRT = DRT + DR
                        DGT = DGT + DG
                        DBT = DBT + DB
                        
                        
                        
                        'tmpMR = (CInt(.R(x, y) + CInt(Coll(CC).Z(i).R(X2, Y2)))) \ 2
                       
                        'tmpMR = .R(x, y) \ 2 + Coll(CC).Z(i).R(X2, Y2) \ 2
                         tmpMR = (.R(x, y) \ 1 + Coll(CC).Z(i).R(X2, Y2) \ 1) \ 2
                        
                        'no need of sqr see color distance.txt
                        tmpFIT = tmpFIT + ( _
                                FASTColorDistanceR(DR, tmpMR) + _
                                FASTColorDistanceG(DG) + _
                                FASTColorDistanceB(DB, tmpMR) _
                                )
                        
            '            '***************** HSP
            '            cH1 = cH1 - cH2
            '            cS1 = cS1 - cS2
            '            cP1 = cP1 - cP2
            '            tmpFIT = tmpFIT + fastSQR(cH1 * cH1 + cS1 * cS1 + cP1 * cP1)
            '            '***********************************
                        
                        
                        
                        
                        
                        
                        'tmpFIT = tmpFIT + Sqr(DR * DR + DG * DG + DB + DB)
                        
                        
                        'risultato da 0 a 765 (con sqr) 585.225  (senza sqr)
                        'con FASTColorDistance senza /1024
                        
                        
                    Next x
                Next y
                '---------
                
                'tmpMR = tmpmpr / (NofPixels)
                
                DRT = DRT / NofPixels
                DGT = DGT / NofPixels
                DBT = DBT / NofPixels
                
                agR = agR / NofPixels
                agG = agG / NofPixels
                agB = agB / NofPixels
                
                
                .agR(Real_I) = agR
                .agG(Real_I) = agG
                .agB(Real_I) = agB
                
                
                
                .FIT(Real_I) = tmpFIT / NofPixels
                ' THIS
                'FM.toSEE(Z).FIT(Real_I) = ( _
                ((512 + tmpMR) * DRT * DRT) / 256 + _
                        4 * DGT * DGT + _
                        ((767 - tmpMR) * DBT * DBT) / 256 _
                        ) 'sqr
                
                'default
                'FM.toSEE(Z).FIT(Real_I) = (DRT * DRT + DGT * DGT + DBT * DBT) 'sqr
                
                
                'smartdesigntech
                'FM.toSEE(Z).FIT(Real_I) = 0.5 * DRT * DRT + DGT * DGT + 0.25 * DBT * DBT
                
                
                'weighted Euclidean distance
                '1
                'FM.toSEE(Z).FIT(Real_I) = Sqr(3 * DRT * DRT + 4 * DGT * DGT + 2 * DBT * DBT)
                '2
                'FM.toSEE(Z).FIT(Real_I) = Sqr(2 * DRT * DRT + 4 * DGT * DGT + 3 * DBT * DBT)
                '3
                'FM.toSEE(Z).FIT(Real_I) = Sqr(0.3 * DRT * DRT + 0.59 * DGT * DGT + 0.11 * DBT * DBT)
                
                
                If .FIT(Real_I) < BestFIT Then
                    
                    BestFIT = .FIT(Real_I)
                    .indexBESTFIT = Real_I
                    ' FM.toSEE(Z).indexBFfileName = Coll(CC).Z(I).filename
                    
                    '''''             Debug.Print "Bestfit " & BestFIT & "  z=" & Z & "  i=" & Real_I & "     "; FM.toSEE(Z).agR(i); " " & FM.toSEE(Z).agB(i); " "; FM.toSEE(Z).agB(i)
                    '' c era disegna qui
                    ''''Disegna
                    'For Y = 1 To FM.WZone_H
                    'For X = 1 To FM.WZone_W
                    'X2 = X * k + sx
                    'Y2 = Y * k + sy
                    'Me.Line (5 + X * 5, 100 + Y * 5)-(5 + (X + 1) * 5, 100 + (Y + 1) * 5), RGB(Coll.Z(I).R(X2, Y2), Coll.Z(I).G(X2, Y2), Coll.Z(I).B(X2, Y2)), BF
                    'Next
                    'Next
                    '''''
                End If
                
            Next i
            
        Next CC
        
    End With 'fm.tosee
Next Z


IsBuilding = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' remove duplicates

Do
    SCX = 0
    
    Uguali = False
    For Z1 = 1 To FM.NZones
        For Z2 = 1 To FM.NZones
            If Z1 <> Z2 Then
                '            Stop
                
                If FM.toSEE(Z1).indexBESTFIT = FM.toSEE(Z2).indexBESTFIT Then
                    'ListaNOMI(FM.toSEE(z1).indexBESTFIT) = ListaNOMI(FM.toSEE(z2).indexBESTFIT) Then 'To Consider Mirror tiles Equals. But Don't know why do not work
                    '                    Stop
                    
                    ''''''''''''''''''''''''''''''******
                    
                    dx = Abs(FM.toSEE(Z1).CX - FM.toSEE(Z2).CX) / FM.GlobTileW
                    dy = Abs(FM.toSEE(Z1).CY - FM.toSEE(Z2).CY) / FM.GlobTileH
                    
                    
                    
                    If (dx <= MINDIST) And (dy <= MINDIST) Then
                        
                        '*****************************************
                        
                        Uguali = True
                        SCX = SCX + 1
                        
                        ''''                   Debug.Print "z1=" & Z1 & " z2=" & Z2 & "    " & FM.toSEE(Z1).FIT(FM.toSEE(Z1).indexBESTFIT) & " " & FM.toSEE(Z2).FIT(FM.toSEE(Z2).indexBESTFIT) & " index=" & FM.toSEE(Z2).indexBESTFIT
                        
                        If FM.toSEE(Z1).FIT(FM.toSEE(Z1).indexBESTFIT) > FM.toSEE(Z2).FIT(FM.toSEE(Z2).indexBESTFIT) Then
                            cH = Z1
                            ToG = Z2
                        Else
                            cH = Z2
                            ToG = Z1
                        End If
                        
                        
                        
                        '**************** secondo modo ******************
SecondoModo:
                        'Stop
                        ' trova minimo FM.toSEE(ch).indexBESTFIt
                        BEST = 1E+25
                        '        BEST = -1
                        I_pic = FM.toSEE(cH).indexBESTFIT
                        For ZZ = 1 To FM.NZones
                            If FM.toSEE(ZZ).indexBESTFIT = I_pic Then
                                If FM.toSEE(ZZ).FIT(FM.toSEE(ZZ).indexBESTFIT) < BEST Then
                                    BEST = FM.toSEE(ZZ).FIT(FM.toSEE(ZZ).indexBESTFIT)
                                    ZonaBest = ZZ
                                End If
                            End If
                        Next ZZ
                        '''''            Debug.Print "Best zone for pic " & I_pic & " is zone " & ZonaBest
                        
                        'tutti tranne zonabest  NOT usable
                        For ZZ = 1 To FM.NZones
                            If ZZ <> ZonaBest Then
                                FM.toSEE(ZZ).FitINDEXusable(I_pic) = False
                            End If
                        Next ZZ
                        
                        
                        
                        'trova nuovo BEST per zona CH
                        ''' trova ch nuovo bestfit
                        BestFIT = 1E+23
                        For II = 1 To TotalSourcePhotos 'Coll(CC).NofPhotos
                            If (FM.toSEE(cH).FIT(II) < BestFIT) Then
                                If (FM.toSEE(cH).FitINDEXusable(II) = True) Then
                                    BestFIT = FM.toSEE(cH).FIT(II)
                                    
                                    FM.toSEE(cH).indexBESTFIT = II
                                    'Debug.Print FM.toSEE(ch).FitINDEXusable(ii)
                                End If
                            End If
                        Next II
                        
                        
                        ''''             Debug.Print "new bestfit pic " & FM.toSEE(CH).indexBESTFIT & " for Zone " & CH
                        ''''             Debug.Print "-------------------------------------------------"
                        '**************** fine secondo modo
                        'Stop
                        
                    End If
                End If 'FM.toSEE(z1).indexBESTFIT = FM.toSEE(z2).indexBESTFIT
            End If 'If z1 <> z2
        Next Z2
    Next Z1
    '    Stop
    
    If STATUS = "removing duplicates... " & SCX Then
        SameCounter = SameCounter + 1
    Else
        SameCounter = 0
    End If
    
    STATUS = "removing duplicates... " & SCX
    Label1 = SCX
    DoEvents
    ''''''Loop While uguali = True
    If SameCounter = 100 Then
        
        Msg = "In Collection(s) There are " & TotalSourcePhotos & " tiles." & _
                vbCrLf & "Photomosaic should be made by " & FM.NZones & " Tiles." & vbCrLf & _
                "Minimal Distance between Identical tiles is " & MINDIST & "." & vbCrLf & _
                vbCrLf & "Suggest to Have more Photos in Collection(s)" & _
                vbCrLf & "or decrease Minimal Distance between Tiles." & vbCrLf & _
                vbCrLf & "Now process will continue breaking 'Minimal Distance' Rule."
        STATUS = Msg
        MsgBox Msg, vbInformation, "Trouble!"
        
        SCX = 0 'means exit loop
    End If
    
    
Loop While SCX > 0
'Stop

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''' assegna nome file index bestfit + INDICE CORRETTEZZA
ErrorI = 0
For Z = 1 To FM.NZones
    With FM.toSEE(Z)
        
        '.indexBFfileName = (Coll.Z(.indexBESTFIT).filename)
        .indexBFfileName = (ListaNOMI(.indexBESTFIT))
        ErrorI = ErrorI + .FIT(.indexBESTFIT)
    End With
Next Z
'ErrorI = (((ErrorI / FM.NZones) / 3) / 256)
ErrorI = (((ErrorI / FM.NZones) / Sqr(255 * 3)))

''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''
'''''''''''''''''' DRAWORDER
STATUS = "SORTING Draw Order... (Wrongest first)"
DoEvents
'Stop

QuickSort 1, FM.NZones


' Mirror  '
''''''''''''''''''''''''''''
For Z1 = 1 To FM.NZones
    FM.toSEE(Z1).Mirrored = GlobalMirrored(FM.toSEE(Z1).indexBESTFIT)
Next
''''''''''''''''''''''
'Stop
'verifica drawo
For Z1 = 1 To FM.NZones
    Debug.Print FM.toSEE(Z1).FIT(FM.toSEE(Z1).indexBESTFIT)
Next Z1

'Stop

End Sub

Sub Ricostruz()
Dim SCALA As Single
''''''''''''''''''''''''
Dim ix As Single
Dim iy As Single
Dim K As Single
Dim sX As Single
Dim sY As Single
Dim eX As Single
Dim eY As Single
Dim c As Long
''''''''''''''''''''''''
Dim wX As Single
Dim wY As Single
Dim id As Single
Dim iX2 As Single
Dim iY2 As Single
Dim iD2 As Single
'''''''''''''''''''''''
Dim adjP As Single


SCALA = OUTS / 10
PicR.Cls
PicR.Width = FM.TotalW * SCALA + 1
PicR.Height = FM.TotalH * SCALA + 1


PicMask.Width = FM.GlobTileW * SCALA + 1
PicMask.Height = FM.GlobTileH * SCALA + 1
'''' Load PicMask
If FM.MaskName <> "NOTHING" Then
    
    picLoad = LoadPicture(App.Path & "\masks\" & FM.MaskName)
    
    Call SetStretchBltMode(PicMask.hdc, STRETCHMODE)
    Call StretchBlt(PicMask.hdc, 0, 0, PicMask.Width, PicMask.Height - 1, _
            picLoad.hdc, 0, 0, picLoad.Width - 1, picLoad.Height - 1, vbSrcCopy)
    
    
    PicMask.Refresh
End If
'''''


Tempo = Timer
PBmax = FM.NZones
PBvalue = 0
For Z = 1 To FM.NZones
    
    If (Z = 1) Or (Z \ 10 = Z / 10) Then STATUS = "Drawing... " & Z & "/" & FM.NZones
    
    PBvalue = Z
    ix = 0
    iy = 0
    With FM.toSEE(Z)
        picLoad = LoadPicture(.indexBFfileName)
        
        
        'If ADJ.Value = Checked Then Aggiusta
        
        
        Debug.Print FM.GlobTileW & " X " & FM.GlobTileH & "  " & picLoad.Width & " X " & picLoad.Height
        
        'If picLoad.Width / picLoad.Height > FM.GlobTileW / FM.GlobTileH Then
        If picLoad.Width / picLoad.Height > .OnPmWidht / .OnPmHeight Then
            
            K = .OnPmHeight / picLoad.Height
            ix = (picLoad.Width - .OnPmWidht / K) / 2
            
        Else
            K = .OnPmWidht / picLoad.Width
            iy = (picLoad.Height - .OnPmHeight / K) / 2
        End If
        Debug.Print "                               ix-iy " & ix & "-" & iy & "   k=" & K
        '    Stop
        
        
        If FmTYPE <> "CIRCLED_LR" _
          And FmTYPE <> "CIRCLED_UD" _
          And FmTYPE <> "ANG_OVERLAP_RND" _
          And FmTYPE <> "ANG_OVERLAP_COL" Then
           
           
      
           
            If FM.MaskName = "NOTHING" Then
                ' Mirror  '
                If .Mirrored = False Then
                
                    Call SetStretchBltMode(PicR.hdc, STRETCHMODE)
                    Call StretchBlt(PicR.hdc, (.CX - .OnPmWidht \ 2) * SCALA - 1, _
                        (.CY - .OnPmHeight \ 2) * SCALA, _
                        .OnPmWidht * SCALA + 1, .OnPmHeight * SCALA + 1, _
                        picLoad.hdc, ix, iy, picLoad.Width - ix * 2 - 1, picLoad.Height - iy * 2 - 1, vbSrcCopy)
                Else 'is Mirrored
                
                    Call SetStretchBltMode(PicR.hdc, STRETCHMODE)
                    Call StretchBlt(PicR.hdc, (.CX - .OnPmWidht \ 2) * SCALA + .OnPmWidht * SCALA - 1, _
                        (.CY - .OnPmHeight \ 2) * SCALA, _
                        -(.OnPmWidht * SCALA + 1), .OnPmHeight * SCALA + 1, _
                        picLoad.hdc, ix, iy, picLoad.Width - ix * 2 - 1, picLoad.Height - iy * 2 - 1, vbSrcCopy)
                
                End If
            
            
            Else
            
            '            DRAWmask2 (.cX - FM.GlobTileW \ 2) * SCALA, _
            (.cY - FM.GlobTileH \ 2) * SCALA, _
                    FM.GlobTileW * SCALA, FM.GlobTileH * SCALA, _
                    ix, iy, _
                    picLoad.Width - ix * 2 - 1, picLoad.Height - iy * 2 - 1
            
            DRAWMASK2 (.CX - .OnPmWidht \ 2) * SCALA, _
                    (.CY - .OnPmHeight \ 2) * SCALA, _
                    .OnPmWidht * SCALA + 1, .OnPmHeight * SCALA + 1, _
                    ix, iy, _
                    picLoad.Width - ix * 2 - 1, picLoad.Height - iy * 2 - 1, .Mirrored
            
            
            End If
        
      
     
            
        Else
        
            '"CIRCLED_LR" ,"CIRCLED_UD","ANG_OVERLAP_RND","ANG_OVERLAP_COL"
        
            RotaPIC.Cls
            RotaPIC.Width = .OnPmWidht * SCALA + 1
            RotaPIC.Height = .OnPmHeight * SCALA + 1
        
            If .Mirrored = False Then
            
                Call SetStretchBltMode(RotaPIC.hdc, STRETCHMODE)
                Call StretchBlt(RotaPIC.hdc, 0, 0, _
                    .OnPmWidht * SCALA + 2, .OnPmHeight * SCALA + 2, _
                    picLoad.hdc, ix, iy, picLoad.Width - ix * 2 - 1, picLoad.Height - iy * 2 - 1, vbSrcCopy)
            Else
                Call SetStretchBltMode(RotaPIC.hdc, STRETCHMODE)
                Call StretchBlt(RotaPIC.hdc, .OnPmWidht * SCALA, 0, _
                    -.OnPmWidht * SCALA - 2, .OnPmHeight * SCALA + 2, _
                    picLoad.hdc, ix, iy, picLoad.Width - ix * 2 - 1, picLoad.Height - iy * 2 - 1, vbSrcCopy)
            
            End If
        
        
        
        If ADJ.Value = Checked Then
            adjP = adjPERC.Value / 100
            FX.GetBits RotaPIC.Image.handle
            FX.MYADD .agR(.indexBESTFIT) * adjP, .agG(.indexBESTFIT) * adjP, .agB(.indexBESTFIT) * adjP
            FX.SetBits RotaPIC.Image.handle
        End If
        
        RotaPIC.Refresh
                
        ROTATE PicR.hdc, -.ANG * PI / 180, .CX * SCALA, .CY * SCALA, RotaPIC.Width, RotaPIC.Height, RotaPIC.Image.handle
        
        
        'PicR.Refresh ''
                
             
    End If

    

    ''
    If FmTYPE <> "CIRCLED_LR" _
            And FmTYPE <> "CIRCLED_UD" _
            And FmTYPE <> "ANG_OVERLAP_RND" _
            And FmTYPE <> "ANG_OVERLAP_COL" Then
        If ADJ.Value = Checked And FM.MaskName = "NOTHING" Then
            sX = (.CX - .OnPmWidht \ 2) * SCALA
            sY = (.CY - .OnPmHeight \ 2) * SCALA
            eX = sX + .OnPmWidht * SCALA
            eY = sY + .OnPmHeight * SCALA
            aggiusta2 sX, sY, eX, eY: PicR.Refresh
        End If
    End If


End With

If Z / 100 = Z \ 100 Then
    PicR.Refresh
    If asBMP Then
        SavePicture PicR.Image, FM.FilePathOUT & ".bmp"
    Else
        SaveJPG PicR.Image, FM.FilePathOUT & ".jpg", jpgQuality \ 2
    End If
    
End If
DoEvents
Next Z

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BLEND
If AdjBLEND > 0 Then
    STATUS = "Blending....  Please wait."
    DoEvents
    PicMask = LoadPicture(FM.FilePathIN)
    picLoad.Width = PicR.Width
    picLoad.Height = PicR.Height
    Call SetStretchBltMode(picLoad.hdc, STRETCHMODE)
    Call StretchBlt(picLoad.hdc, 0, 0, picLoad.Width, picLoad.Height, _
            PicMask.hdc, 0, 0, PicMask.Width, PicMask.Height, vbSrcCopy)
    
    PicMask.Width = PicR.Width
    PicMask.Height = PicR.Height
    c = RGB(255 - 255 * AdjBLEND / 100, 255 - 255 * AdjBLEND / 100, 255 - 255 * AdjBLEND / 100)
    PicMask.Line (0, 0)-(PicMask.Width, PicMask.Height), c, BF
    
    ModMask_Setup picLoad, PicMask, PicR
    ModMask_BLTIT PicR.Width / 2, PicR.Height / 2, PicR
    ModMask_CleanUp
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If asBMP Then
    SavePicture PicR.Image, FM.FilePathOUT & ".bmp"
Else
    SaveJPG PicR.Image, FM.FilePathOUT & ".jpg", jpgQuality
End If

STATUS = "done!"
LWait.Visible = False

If chShowFolder = Checked Then mySHELL.Open App.Path & "\MOSAIC"


End Sub

Private Sub DRAWMASK2(tx, ty, WWW, HHH, sX, sY, sw, SH As Single, Mirr As Boolean)
Dim R As Byte
Dim G As Byte
Dim B As Byte
Dim RR As Integer
Dim GG As Integer
Dim BB As Integer
Dim c As Long
Dim adjP As Single
Dim agRibfp As Single
Dim agGibfp As Single
Dim agBibfp As Single
Dim x As Long
Dim y As Long

' OK
' Since V7.6 Borders are right

PicSRC.Width = WWW
PicSRC.Height = HHH
PicMask.Width = WWW
PicMask.Height = HHH

Call SetStretchBltMode(PicSRC.hdc, STRETCHMODE)
If Not (Mirr) Then
    Call StretchBlt(PicSRC.hdc, 0, 0, _
            WWW, HHH, _
            picLoad.hdc, sX, sY, sw, SH, vbSrcCopy)
    PicSRC.Refresh
Else
    
    Call StretchBlt(PicSRC.hdc, WWW - 1, 0, _
            -WWW, HHH, _
            picLoad.hdc, sX, sY, sw, SH, vbSrcCopy)
    PicSRC.Refresh
    
    
End If


'----------------------------------------------------------------
'Yet to improve with alphablendmask.bas
''''''''''''''''' for adjust color '''''''''''''''''''''''''''
adjP = adjPERC.Value / 100
With FM.toSEE(Z)
    agRibfp = .agR(.indexBESTFIT) * adjP
    agGibfp = .agG(.indexBESTFIT) * adjP
    agBibfp = .agB(.indexBESTFIT) * adjP
End With
For y = 0 To PicSRC.Height - 1
    For x = 0 To PicSRC.Width - 1
        c = GetPixel(PicSRC.hdc, x, y)
        Long2RGB c, R, G, B
        RR = R + agRibfp
        GG = G + agGibfp
        BB = B + agBibfp
        
        RR = IIf(RR < 0, 0, RR)
        GG = IIf(GG < 0, 0, GG)
        BB = IIf(BB < 0, 0, BB)
        
        RR = IIf(RR > 255, 255, RR)
        GG = IIf(GG > 255, 255, GG)
        BB = IIf(BB > 255, 255, BB)
        
        SetPixel PicSRC.hdc, x, y, RGB(RR, GG, BB)
    Next x
Next y
PicSRC.Refresh
'''''''''''''''''''''''''''''''''''
'---------------------------------------------------------------


'MASK'
'''''''''''''''''''
picLoad = LoadPicture(App.Path & "\masks\" & FM.MaskName)
picLoad.Refresh
Call SetStretchBltMode(PicMask.hdc, STRETCHMODE)
Call StretchBlt(PicMask.hdc, 0, 0, _
        WWW, HHH, _
        picLoad.hdc, 0, 0, picLoad.Width, picLoad.Height, vbSrcCopy)
PicMask.Refresh
'''''''''''''''''''''''

PicMask.Visible = True

'--------------------------------------------------------------------
ModMask_Setup PicSRC, PicMask, PicR
ModMask_BLTIT tx + WWW / 2, ty + HHH / 2, PicR
ModMask_CleanUp

'---------------------------------------------------------------

End Sub
Private Sub DRAWMASK(tx, ty, GlobTileW, GlobTileH, sX, sY, sw, SH As Single)
'
'Dim maskC As Long
'Dim mR As Byte
'Dim mG As Byte
'Dim MB As Byte
'Dim soC As Long
'Dim soR As Byte
'Dim soG As Byte
'Dim soB As Byte
'Dim pC As Long
'Dim pR As Byte
'Dim pG As Byte
'Dim PB As Byte
'Dim TR As Single
'Dim TG As Single
'Dim TB As Single'
'
'
'Dim FX As Single
'Dim fY As Single'
'
'Dim MASKsi As Single
'Dim MASKno As Single
'
'Dim adjP As Single
'Dim agRibfp As Single
'Dim agGibfp As Single
'Dim agBibfp As Single
'Dim x As Long
'Dim y As Long'
'
'Dim trR As Integer
'Dim trG As Integer
'Dim trB As Integer
''
'
'adjP = adjPERC.Value / 100
'With FM.toSEE(Z)
'    agRibfp = .agR(.indexBESTFIT) * adjP
'    agGibfp = .agG(.indexBESTFIT) * adjP
'    agBibfp = .agB(.indexBESTFIT) * adjP
'End With''
'
'For fY = ty To ty + GlobTileH - 1
'    For FX = tx To tx + GlobTileW - 1
'
'        maskC = GetPixel(PicMask.hdc, FX - tx, fY - ty)
'        If maskC < 0 Then maskC = 0
'        Long2RGB maskC, mR, mG, MB
'''
'        Call SetStretchBltMode(RotaPIC.hdc, STRETCHMODE)
'        Call StretchBlt(RotaPIC.hdc, 0, 0, _
'                1, 1, _
'                picLoad.hdc, _
'                sX + (FX - tx) * (sw / GlobTileW), sY + (fY - ty) * (SH / GlobTileH), _
'                (sw / GlobTileW), (SH / GlobTileH), vbSrcCopy)
'
'
'        soC = GetPixel(PicR.hdc, FX, fY)
'        pC = GetPixel(RotaPIC.hdc, 0, 0)
'
'        If soC < 0 Then soC = 0
'        If pC < 0 Then pC = 0
'
'        Long2RGB soC, soR, soG, soB
'        Long2RGB pC, pR, pG, PB
'
'        MASKsi = (255 - mR) / 256
'        MASKno = 1 - MASKsi
'
'        trR = (soR) * MASKno + (pR + agRibfp) * MASKsi
'        trG = (soG) * MASKno + (pG + agGibfp) * MASKsi
'        trB = (soB) * MASKno + (PB + agBibfp) * MASKsi
'
'
'        trR = IIf(trR < 0, 0, trR)
'        trG = IIf(trG < 0, 0, trG)
'        trB = IIf(trB < 0, 0, trB)
'
'        trR = IIf(trR > 255, 255, trR)
'        trG = IIf(trG > 255, 255, trG)
'        trB = IIf(trB > 255, 255, trB)
'
'
'        SetPixel PicR.hdc, FX, fY, RGB(trR, trG, trB)
'
'
'    Next FX
'    DoEvents
'Next fY
'PicR.Refresh
'DoEvents
'
''Stop
'
'
End Sub

Sub Aggiusta()
'Dim XX As Long
'Dim YY As Long
'Dim ncR As Integer
'Dim ncG As Integer
'Dim ncB As Integer
'Dim R1 As Byte
'Dim G1 As Byte '
'Dim B1 As Byte
'With FM.toSEE(Z)
'    STATUS = "Rebuildin with Adjusting... " & FM.toSEE(Z).agR(.indexBESTFIT) & " " & _
'            FM.toSEE(Z).agG(.indexBESTFIT) & " " & _
'            FM.toSEE(Z).agB(.indexBESTFIT)
'
'    For YY = 0 To picLoad.Height - 1
'        For XX = 0 To picLoad.Width - 1
'            ccc = GetPixel(picLoad.hdc, XX, YY)
'
'
'            Long2RGB ccc, R1, G1, B1
'
'            ncR = R1 + .agR(.indexBESTFIT)
'            ncG = G1 + .agG(.indexBESTFIT)
'            ncB = B1 + .agB(.indexBESTFIT)
'
'            ncR = IIf(ncR < 0, 0, ncR)
'            ncG = IIf(ncG < 0, 0, ncG)
'            ncB = IIf(ncB < 0, 0, ncB)
'
'            ncR = IIf(ncR > 255, 255, ncR)
'            ncG = IIf(ncG > 255, 255, ncG)
'            ncB = IIf(ncB > 255, 255, ncB)
'
'
'            SetPixel picLoad.hdc, XX, YY, RGB(ncR, ncG, ncB)
'
'
'        Next XX
'    Next YY
'
'
'End With
'DoEvents'
'
End Sub
Sub aggiusta2(sX, sY, eX, eY)
'Yet to improve with alphablendmask.bas

Dim XX As Long
Dim YY As Long
Dim ncR As Integer
Dim ncG As Integer
Dim ncB As Integer
Dim R1 As Byte
Dim G1 As Byte
Dim B1 As Byte

Dim agRibfp As Single
Dim agGibfp As Single
Dim agBibfp As Single

Dim adjP As Single

adjP = adjPERC.Value / 100

With FM.toSEE(Z)
    'STATUS = "Rebuildin with Adjusting... " & FM.toSEE(Z).agR(.indexBESTFIT) & " " & _
    FM.toSEE(Z).agG(.indexBESTFIT) & " " & _
            FM.toSEE(Z).agB(.indexBESTFIT)
    
    agRibfp = .agR(.indexBESTFIT) * adjP
    agGibfp = .agG(.indexBESTFIT) * adjP
    agBibfp = .agB(.indexBESTFIT) * adjP
    
    For YY = sY To eY - 1
        For XX = sX To eX - 1
            ccc = GetPixel(PicR.hdc, XX, YY)
            ' If ccc < 0 Then MsgBox "Error ccc=" & ccc & " at " & xx & " " & yy: ccc = 0
            If ccc < 0 Then ccc = 0
            
            Long2RGB ccc, R1, G1, B1
            
            'If FM.MaskName <> "NOTHING" Then
            '     maskC = GetPixel(PicMask.hdc, xx - sx, yy - sy)
            '     Long2RGB maskC, mR, mG, mB
            '     MASKsi = (255 - mR) / 256
            '     MASKno = 1 - MASKsi
            'End If
            
            
            ncR = (R1 + agRibfp) '* MASKsi
            ncG = (G1 + agGibfp) '* MASKsi
            ncB = (B1 + agBibfp) '* MASKsi
            
            ncR = IIf(ncR < 0, 0, ncR)
            ncG = IIf(ncG < 0, 0, ncG)
            ncB = IIf(ncB < 0, 0, ncB)
            
            ncR = IIf(ncR > 255, 255, ncR)
            ncG = IIf(ncG > 255, 255, ncG)
            ncB = IIf(ncB > 255, 255, ncB)
            
            SetPixel PicR.hdc, XX, YY, RGB(ncR, ncG, ncB)
            
        Next XX
    Next YY
    
    
End With
DoEvents
End Sub
Sub AGGIUSTAROT()
''Yet to improve with alphablendmask.bas
'
'Dim XX As Long
'Dim YY As Long
'Dim ncR As Integer
'Dim ncG As Integer
'Dim ncB As Integer
'Dim R1 As Byte
'Dim G1 As Byte
'Dim B1 As Byte
'
'Dim agRibfp As Single
'Dim agGibfp As Single
'Dim agBibfp As Single
'
'Dim adjP As Single'
'
'adjP = adjPERC.Value / 100'
'
'With FM.toSEE(Z)
'
'    agRibfp = .agR(.indexBESTFIT) * adjP
'    agGibfp = .agG(.indexBESTFIT) * adjP
'    agBibfp = .agB(.indexBESTFIT) * adjP
'
'    For YY = 0 To RotaPIC.Height - 1
'        For XX = 0 To RotaPIC.Width - 1
'            ccc = GetPixel(RotaPIC.hdc, XX, YY)
'            ' If ccc < 0 Then MsgBox "Error ccc=" & ccc & " at " & xx & " " & yy: ccc = 0
'            If ccc < 0 Then ccc = 0
'            Long2RGB ccc, R1, G1, B1
'
'            ncR = (R1 + agRibfp) '* MASKsi
'            ncG = (G1 + agGibfp) '* MASKsi
'            ncB = (B1 + agBibfp) '* MASKsi
'
'            ncR = IIf(ncR < 0, 0, ncR)
'            ncG = IIf(ncG < 0, 0, ncG)
'            ncB = IIf(ncB < 0, 0, ncB)
'
'            ncR = IIf(ncR > 255, 255, ncR)
'            ncG = IIf(ncG > 255, 255, ncG)
'            ncB = IIf(ncB > 255, 255, ncB)
'
'            SetPixel RotaPIC.hdc, XX, YY, RGB(ncR, ncG, ncB)
'
'        Next XX
'    Next YY
'
'
'End With
'DoEvents
End Sub



Function DopoUltimaBarra(S As String) As String
For i = Len(S) To 1 Step -1
    If Mid$(S, i, 1) = "/" Or Mid$(S, i, 1) = "\" Then
        DopoUltimaBarra = Right$(S, Len(S) - i)
        i = 0
    End If
Next
End Function

Function FinoUltimaBarra(S As String) As String
For i = Len(S) To 1 Step -1
    If Mid$(S, i, 1) = "/" Or Mid$(S, i, 1) = "\" Then
        FinoUltimaBarra = left$(S, i - 1)
        i = 0
    End If
Next
End Function

Sub SaveFM()
Dim CC As Long

STATUS = "SAVING... " & DopoUltimaBarra(FM.FilePathOUT)

Open App.Path & "\MOSAIC\" & DopoUltimaBarra(FM.FilePathOUT) & ".txt" For Binary Access Write As 1

WRITEbin FmTYPE & "|" & Replace(ErrorI, ",", ".") & "|"
WRITEbin FM.MaskName & "|"
WRITEbin "var2|var3|var4|var5|" & vbCrLf
WRITEbin FM.FilePathIN & "|" & vbCrLf
WRITEbin "(" & FM.FilePathOUT & ")|" & vbCrLf
WRITEbin "Size|" & FM.TotalW & "|" & FM.TotalH & "|" & vbCrLf
'WRITEbin FM.GlobTileW & "|" & FM.GlobTileH & "|" & vbCrLf
WRITEbin Replace(FM.GlobTileW, ",", ".") & "|" & Replace(FM.GlobTileH, ",", ".") & "|" & vbCrLf
WRITEbin FM.WZone_W & "|" & FM.WZone_H & "|" & vbCrLf
WRITEbin "Tiles|" & FM.NZones & "|" & vbCrLf
WRITEbin "Coll Used|" & UBound(Coll) & "|" & TotalSourcePhotos & "|" & vbCrLf
For CC = 1 To UBound(Coll)
    WRITEbin Coll(CC).NAME & "|"
Next CC
WRITEbin vbCrLf


For Z = 1 To FM.NZones
    With FM.toSEE(Z)
        WRITEbin vbCrLf
        
        WRITEbin .indexBESTFIT & "|" & vbCrLf
        WRITEbin .indexBFfileName & "|" & vbCrLf
        
        WRITEbin Replace(.CX, ",", ".") & "|" & Replace(.CY, ",", ".") & "|" & vbCrLf
        
        WRITEbin Replace(.OnPmWidht, ",", ".") & "|" & Replace(.OnPmHeight, ",", ".") & "|" & vbCrLf
        
        
        WRITEbin .agR(.indexBESTFIT) & "|" & .agG(.indexBESTFIT) & "|" & .agB(.indexBESTFIT) & "|" & vbCrLf
        WRITEbin Replace(CStr(.ANG), ",", ".") & "|" & vbCrLf
        WRITEbin IIf(.Mirrored, "True", "False") & "|" & vbCrLf
    End With
Next Z

Close 1


End Sub

Private Function LoadFM(filename As String) As Boolean
'Stop
Dim st As String
Dim aCapo As Byte
Dim cUsed As Long
Dim CC As Long

On Error GoTo notValidFIle

STATUS = "Loading... " & DopoUltimaBarra(filename)

'Open App.Path & "\MOSAIC\" & DopoUltimaBarra(FM.FilePathOUT) & ".txt" For Binary Access Write As 1
Open filename For Binary Access Read As 1


'WRITEbin FmTYPE & "|" & ErrorI & "|1|2|3|4|5|" & vbCrLf
FmTYPE = ReadBin
ErrorI = ReadBin

FM.MaskName = ReadBin
'FM.MaskPercX = CSng(ReadBin)
'FM.MaskPercY = CSng(ReadBin)
ReadBin
ReadBin
ReadBin
ReadBin

Get #1, , aCapo
Get #1, , aCapo
'WRITEbin FM.FilePathIN & "|" & vbCrLf
FM.FilePathIN = ReadBin
Get #1, , aCapo
Get #1, , aCapo
'WRITEbin "(" & FM.FilePathOUT & ")|" & vbCrLf
FM.FilePathOUT = ReadBin
FM.FilePathOUT = Mid$(FM.FilePathOUT, 2, Len(FM.FilePathOUT) - 2)
Get #1, , aCapo
Get #1, , aCapo
'WRITEbin "Size|" & FM.TotalW & "|" & FM.TotalH & "|" & vbCrLf
ReadBin
FM.TotalW = CInt(ReadBin)
FM.TotalH = CInt(ReadBin)
Get #1, , aCapo
Get #1, , aCapo
'WRITEbin FM.GlobTileW & "|" & FM.GlobTileH & "|" & vbCrLf
'FM.GlobTileW = CSng(readbin)
'FM.GlobTileH = CSng(readbin)
FM.GlobTileW = Val(ReadBin)
FM.GlobTileH = Val(ReadBin)


Get #1, , aCapo
Get #1, , aCapo
'WRITEbin FM.WZone_W & "|" & FM.WZone_H & "|" & vbCrLf
FM.WZone_W = CInt(ReadBin)
FM.WZone_H = CInt(ReadBin)
Get #1, , aCapo
Get #1, , aCapo
'WRITEbin "Tiles|" & FM.NZones & "|" & vbCrLf
ReadBin
FM.NZones = CInt(ReadBin)
Get #1, , aCapo
Get #1, , aCapo
'WRITEbin "Coll Used|" & UBound(Coll) & "|" & TotalSourcePhotos & "|" & vbCrLf
ReadBin
cUsed = CInt(ReadBin)
TotalSourcePhotos = CLng(ReadBin)
Get #1, , aCapo
Get #1, , aCapo
''
For CC = 1 To cUsed
    ReadBin
Next CC
Get #1, , aCapo
Get #1, , aCapo

ReDim FM.toSEE(FM.NZones)
For Z = 1 To FM.NZones
    
    With FM.toSEE(Z)
        ReDim .agR(TotalSourcePhotos)
        ReDim .agG(TotalSourcePhotos)
        ReDim .agB(TotalSourcePhotos)
        
        'WRITEbin vbCrLf
        Get #1, , aCapo
        Get #1, , aCapo
        
        'WRITEbin .indexbesfit & "|" & vbCrLf
        .indexBESTFIT = ReadBin
        Get #1, , aCapo
        Get #1, , aCapo
        
        
        'WRITEbin .indexBFfileName & "|" & vbCrLf
        .indexBFfileName = ReadBin
        Get #1, , aCapo
        Get #1, , aCapo
        '**************************
        'WRITEbin .cX & "|" & .cY & "|" & vbCrLf
        '.cX = CSng(readbin)
        '.cY = CSng(readbin) due to   "," "."
        .CX = Val(ReadBin)
        .CY = Val(ReadBin)
        '        Stop
        '**************************
        
        
        Get #1, , aCapo
        Get #1, , aCapo
        .OnPmWidht = Val(ReadBin)
        .OnPmHeight = Val(ReadBin)
        
        
        
        
        
        
        Get #1, , aCapo
        Get #1, , aCapo
        'WRITEbin .agR(.indexBESTFIT) & "|" & .agG(.indexBESTFIT) & "|" & .agB(.indexBESTFIT) & "|" & vbCrLf
        .agR(.indexBESTFIT) = CInt(ReadBin)
        .agG(.indexBESTFIT) = CInt(ReadBin)
        .agB(.indexBESTFIT) = CInt(ReadBin)
        Get #1, , aCapo
        
        
        'WRITEbin .ANG & "|" & vbCrLf
        .ANG = Val(ReadBin)
        Get #1, , aCapo
        Get #1, , aCapo
        st = ReadBin
        .Mirrored = IIf(st = "True", True, False)
        Get #1, , aCapo
        Get #1, , aCapo
    End With
Next Z
'Stop

Close 1

LoadFM = True

GoTo fileOK
notValidFIle:
Close 1
MsgBox "WRONG FILE!!!" & vbCr & "(" & Err.Description & ")", vbCritical, "cant load!"

LoadFM = False
Err.Clear

fileOK:
End Function
Sub RefreshListCOl()
LWait.Visible = True

Dim numC As Long

''--------------------
'''' Riempie LISTCOL
ListCol.ListItems.Clear
ListCol.Sorted = False

Set FO = fs.GetFolder(App.Path & "\coll")
i = 0
For Each picFILE In FO.Files
    '    Stop
    
    If Right$(picFILE, 3) = "COL" Then
        
        i = i + 1
        ListCol.ListItems.Add i, , picFILE.NAME
        Open App.Path & "\coll\" & picFILE.NAME For Binary Access Read As 3
        numC = ReadBin(3)
        ListCol.ListItems.Item(i).SubItems(1) = Format(numC, "00000")
        Close 3
        
    End If
    
Next
'ListCol.SortKey = 1
ListCol.SortOrder = lvwDescending
ListCol.Sorted = True
'''_________________
LWait.Visible = False
End Sub

Sub InitFastSQR()
Dim i7 As Long
For i7 = 0 To 195075
'For i7 = 0 To 1530
    fastSQR(i7) = Sqr(i7)
Next i7

End Sub



Sub InitFASTColorDistanceR()
Dim Min
Dim Max
Max = -99999999
Min = 99999999999#
Dim rrr
Dim mer
Dim c As Byte

For rrr = 0 To 255 '-255
    For mer = 0 To 255
        
        FASTColorDistanceR(rrr, mer) = ((2 + mer / 256) * rrr * rrr) / 768
        c = FASTColorDistanceR(rrr, mer) * 1 '0.5
        If c > 255 Then Stop
        SetPixel Me.hdc, Abs(rrr), mer, RGB(c, c, c)
        If c > Max Then Max = c
        If c < Min Then Min = c
    Next
Next
Debug.Print Min, Max

End Sub

Sub InitFASTColorDistanceG()
Dim Min
Dim Max
Max = -99999999
Min = 99999999999#
Dim ggg
Dim mer
Dim c As Byte
For ggg = 0 To 255 '-255
    For mer = 0 To 255
        FASTColorDistanceG(ggg) = (4 * ggg * ggg) / 1024
        'Stop
        
        c = ((FASTColorDistanceG(ggg))) * 1 '0.5
        If c > 255 Then Stop
        SetPixel Me.hdc, Abs(ggg), mer + 256, RGB(c, c, c)
        If c > Max Then Max = c
        If c < Min Then Min = c
    Next
Next
Debug.Print Min, Max

End Sub


Sub InitFASTColorDistanceB()
Dim Min
Dim Max
Max = -99999999
Min = 99999999999#
Dim bbb
Dim mer
Dim c As Byte
For bbb = 0 To 255 '-255
    For mer = 0 To 255
        FASTColorDistanceB(bbb, mer) = ((2 + (255 - mer) / 256) * bbb * bbb) / 768
        c = FASTColorDistanceB(bbb, mer) * 1 ' 0.5
        If c > 255 Then Stop
        SetPixel Me.hdc, Abs(bbb), 256 + mer + 256 + 2, RGB(c, c, c)
        If c > Max Then Max = c
        If c < Min Then Min = c
    Next
Next
Debug.Print Min, Max
'Stop

End Sub


Sub SAVECOLLECTION(S As String, NumColl)
Dim x As Long
Dim y As Long

'S = S & "_PROVA.COL"
Open App.Path & "\Coll\" & S For Binary Access Write As 1

PBvalue = 0
PBmax = Coll(NumColl).NofPhotos
Tempo = Timer
STATUS = "Saving Collection " & S & "(" & PBmax & " photos)   "


WRITEbin Coll(NumColl).NofPhotos & "|" & vbCrLf
WRITEbin Coll(NumColl).STARTdir & "|" & vbCrLf


For i = 1 To Coll(NumColl).NofPhotos
    PBvalue = i
    DoEvents
    ''''''''''''''''''''''''''
    
    WRITEbin i & "|" & CStr(Coll(NumColl).Z(i).FileDate) & "|"
    WRITEbin Coll(NumColl).Z(i).filename & "|"
    WRITEbin Coll(NumColl).Z(i).oW & "|"
    WRITEbin Coll(NumColl).Z(i).oH & "|"
    WRITEbin Coll(NumColl).Z(i).WW & "|"
    WRITEbin Coll(NumColl).Z(i).HH & "|"
    'Coll(CC).Z(I).oW = CInt(readbin)
    'Coll(CC).Z(I).oH = CInt(readbin)
    'Coll(CC).Z(I).WW = CInt(readbin)
    'Coll(CC).Z(I).HH = CInt(readbin)
    For y = 1 To Coll(NumColl).Z(i).HH
        For x = 1 To Coll(NumColl).Z(i).WW
            'Get #1, , bR
            'Get #1, , bG
            'Get #1, , bB
            'Coll(CC).Z(I).R(X, Y) = bR
            'Coll(CC).Z(I).G(X, Y) = bG
            'Coll(CC).Z(I).b(X, Y) = bB
            Put #1, , Coll(NumColl).Z(i).R(x, y)
            Put #1, , Coll(NumColl).Z(i).G(x, y)
            Put #1, , Coll(NumColl).Z(i).B(x, y)
        Next x
        'Get #1, , aCapo
        'Get #1, , aCapo
        WRITEbin vbCrLf
    Next y
    
    
Next i

Close 1

End Sub

Private Sub UpdateCollecions_Click()
Dim pCONTA As Long
Dim tmpDATE As Double
Dim picI As Long
Dim ZI As Long
Dim iC As Long
Dim Added As Long
Dim Removed As Long
Dim AD As Long
Dim CantRead As Long


AD = IIf(chMIRROR.Value = Checked, True, False)
chMIRROR.Value = Unchecked


LoadCollection_Click
'Stop

For iC = 1 To NumberOfCollections
    
    Added = 0
    Removed = 0
    
    pCONTA = 0
    FindPICFiles (Coll(iC).STARTdir)
    
    ZI = 0
    
    CantRead = 0
    
    For picI = 1 To UBound(Pic2Read)
        
        ZI = ZI + 1
        If picI <= Coll(iC).NofPhotos Then
            
            
            If Pic2Read(picI) <> Coll(iC).Z(ZI).filename Then
                
                If UCase(Pic2Read(picI)) < UCase(Coll(iC).Z(ZI).filename) Then
                    'Stop
                    
                    ' MsgBox "ADD" & vbCrLf & picI & " " & Pic2Read(picI) & vbCrLf & ZI & " " & Coll(iC).Z(ZI).FileName
                    Added = Added + 1
                    STATUS = "Adding " & Added & " - " & Pic2Read(picI)
                    DoEvents
                    CantRead = CantRead + AddHere(picI, ZI, iC)
                    
                Else
                    '            Stop
                    Removed = Removed + 1
                    '                MsgBox "DoNOT know if ti works REMOVE second one " & vbCrLf & picI & " " & Pic2Read(picI) & vbCrLf & ZI & " " & Coll(iC).Z(ZI).FileName
                    RemovePic ZI, iC
                    'ZI = ZI - 1
                End If
            End If
            
            
        Else
            'MsgBox "ADD to END  -   DoNOT know if ti works" & vbCrLf & Pic2Read(picI)
            Added = Added + 1
            CantRead = CantRead + AddHere(picI, ZI, iC)
            
        End If
    Next picI
    '
    Coll(iC).NofPhotos = picI - 1 - CantRead
    
    ' This function don't remove last one... simply change NofPhotos
    'MsgBox "Added " & Added & " Removed " & Removed
    'If Added <> 0 Or Removed <> 0 Then SAVECOLLECTION Left$(Coll(iC).NAME, Len(Coll(iC).NAME) - 4) & "_UPDATED.COL", iC
    If Added <> 0 Or Removed <> 0 Then SAVECOLLECTION Coll(iC).NAME, iC
    
Next iC


RefreshListCOl


chMIRROR.Value = IIf(AD, Checked, Unchecked)

STATUS = "Update Done!"


End Sub


Sub RemovePic(Wich As Long, CC As Long)
Dim AI As Long

Coll(CC).NofPhotos = Coll(CC).NofPhotos - 1


For AI = Wich To Coll(CC).NofPhotos
    
    Coll(CC).Z(AI) = Coll(CC).Z(AI + 1)
    
Next

ReDim Preserve Coll(CC).Z(Coll(CC).NofPhotos)

End Sub


Function AddHere(Source As Long, Target As Long, CC As Long) As Long



Dim AI As Long
Dim XX As Long
Dim YY As Long

Dim RET As Long


On Error GoTo someError

picLoad = LoadPicture(Pic2Read(Source))



Coll(CC).NofPhotos = Coll(CC).NofPhotos + 1
ReDim Preserve Coll(CC).Z(Coll(CC).NofPhotos)

For AI = Coll(CC).NofPhotos To Target + 1 Step -1
    
    Coll(CC).Z(AI) = Coll(CC).Z(AI - 1)
    
Next

''''''''''''''''*****************************''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Manca assegnazione di tutti i valori (scann pic)
Coll(CC).Z(Target).filename = Pic2Read(Source)
Coll(CC).Z(Target).FileDate = Pic2ReadDATE(Source)

'picLoad = LoadPicture(Pic2Read(Source))

WW = picLoad.Width
HH = picLoad.Height


If WW > HH Then
    H2 = ComputationalComplexity '12 '6 '12
    W2 = Round(WW * H2 / HH)
    Debug.Print WW & "x" & HH & " " & W2 & "x" & H2 & "   " & WW / HH & " " & W2 / H2
Else
    W2 = ComputationalComplexity '12 '6 '12
    H2 = Round(HH * W2 / WW)
    Debug.Print WW & "x" & HH & " " & W2 & "x" & H2 & "   " & WW / HH & " " & W2 / H2
End If

PicR.Width = W2
PicR.Height = H2

Call SetStretchBltMode(PicR.hdc, STRETCHMODE)
Call StretchBlt(PicR.hdc, 0, 0, W2, H2, _
        picLoad.hdc, 0, 0, picLoad.Width - 1, picLoad.Height - 1, vbSrcCopy)

PicR.Refresh

'WRITEbin I & "|" & Pic2ReadDATE(I) & "|" & Pic2Read(I) & "|"
'WRITEbin WW & "|"
'WRITEbin HH & "|"
'WRITEbin w2 & "|"
'WRITEbin h2 & "|"
'Stop

Coll(CC).Z(Target).oW = WW
Coll(CC).Z(Target).oH = HH
Coll(CC).Z(Target).WW = W2
Coll(CC).Z(Target).HH = H2

ReDim Coll(CC).Z(Target).R(W2, H2)
ReDim Coll(CC).Z(Target).G(W2, H2)
ReDim Coll(CC).Z(Target).B(W2, H2)

Me.Cls
For YY = 0 To PicR.Height - 1
    For XX = 0 To PicR.Width - 1
        
        ccc = GetPixel(PicR.hdc, XX, YY)
        Long2RGB ccc, Cr, cG, cb
        
        '        Put #1, , Cr
        '        Put #1, , cG
        '        Put #1, , cB
        
        Coll(CC).Z(Target).R(XX + 1, YY + 1) = Cr
        Coll(CC).Z(Target).G(XX + 1, YY + 1) = cG
        Coll(CC).Z(Target).B(XX + 1, YY + 1) = cb
        
        
        Me.Line (5 + XX * 5, 5 + YY * 5)-(5 + (XX + 1) * 5, 5 + (YY + 1) * 5), RGB(Cr, cG, cb), BF
        
    Next XX
Next YY
DoEvents
''''''''''''''''*****************************''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
AddHere = 0
GoTo NOERROR

someError:

RET = MsgBox("Error:" & Err.Number & " " & Err.Description & vbCrLf & Pic2Read(Source) & vbCrLf & _
        vbCrLf & "Yes = Delete and Continue." & vbCrLf & _
        "No = Don't Delete and Continue.", vbYesNo)
If RET = 6 Then Kill Pic2Read(Source)


AddHere = 1


Exit Function





NOERROR:

End Function

Private Sub UpdateCollecions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
DoHelp 11
End Sub

Sub DoHelp(n)

frmHELP.TXT.Text = H(n)
If chHELP.Value = Checked Then frmHELP.Visible = True: ' frmHELP.Top = moY: frmHELP.Left = moX

End Sub

Private Sub QuickSort(ByVal First As Long, ByVal Last As Long)

Dim Low                  As Long
Dim High                 As Long
Dim MidValue             As Double
Dim MidElement           As Long
Dim SawpToSee             As tPMzone

Low = First
High = Last

MidElement = (First + Last) \ 2
MidValue = FM.toSEE(MidElement).FIT(FM.toSEE(MidElement).indexBESTFIT)
Do
    While FM.toSEE(Low).FIT(FM.toSEE(Low).indexBESTFIT) > MidValue '> Bigger first < Lowest first
        Low = Low + 1
    Wend
    While FM.toSEE(High).FIT(FM.toSEE(High).indexBESTFIT) < MidValue '< Bigger first > Lowest first
        High = High - 1
    Wend
    If Low <= High Then
        GoSub SWAP
        Low = Low + 1
        High = High - 1
    End If
Loop While Low <= High
If First < High Then QuickSort First, High
If Low < Last Then QuickSort Low, Last

Exit Sub

SWAP:
SWAPtoSee = FM.toSEE(High)
FM.toSEE(High) = FM.toSEE(Low)
FM.toSEE(Low) = SWAPtoSee

Return

End Sub


Private Sub CREATE2()
Dim nC As Single
Dim Nr As Single
Dim RR As Integer
Dim CC As Integer

Dim GlobTileW As Single
Dim GlobTileH As Single
Dim WZone_W As Integer
Dim WZone_H As Integer
Dim XX As Long
Dim YY As Long
Dim x As Long
Dim y As Long
Dim SizeVariation As Single
Dim Stopped As Boolean
Dim Cntloop As Long
Dim X1 As Single
Dim Y1 As Single
Dim X2 As Single
Dim Y2 As Single

Dim C1 As Long
Dim C2 As Long
Dim R1 As Byte
Dim G1 As Byte
Dim B1 As Byte
Dim R2 As Byte
Dim G2 As Byte
Dim B2 As Byte
Dim DR As Long
Dim DG As Long
Dim DB As Long
Dim mR As Long
Dim dt As Single

Dim TR As Integer
Dim TG As Integer
Dim TB As Integer
Dim Vmin As Integer


STATUS = "Looking Zones...   (Type: " & FmTYPE & ")"
DoEvents

If SubjFileName = "" Then MsgBox "Load Suject Photo! (3)", vbExclamation, "Missing Subject Photo": Exit Sub

If FmTYPE = "STND_MASK" And FM.MaskName = "NOTHING" Then MsgBox "Click a Mask Picture first!", vbInformation: Exit Sub

'If chMIRROR.Value = Checked Then LoadCollection_Click 'because have to load collection in different way "mirror"
'better do LoadCollection Allways
LoadCollection_Click

If TotalSourcePhotos = 0 Then MsgBox "No Collection Selected!", vbInformation: Exit Sub

If chAllowDuplicates.Value = Checked Then
    MINDIST = Val(tMINDIST)
Else
    MINDIST = 999999999
End If

loads (SubjFileName)
FM.FilePathIN = SubjFileName

OUTS_Scroll
'SavePicture picLoad.Image, App.Path & "\IN.bmp"
If asBMP Then
    SavePicture picLoad.Image, App.Path & "\MOSAIC\" & DopoUltimaBarra(FM.FilePathIN)
Else
    SaveJPG picLoad.Image, App.Path & "\MOSAIC\" & DopoUltimaBarra(FM.FilePathIN), 80
End If



PicSRC.Visible = False
PicMask.Visible = False


Select Case FmTYPE
    Case "STANDARD", "OVERLAP_MASK"
    Case "STND_MASK"
        PicSRC.Visible = True
        PicMask.Visible = True
    Case "OVERLAP"
End Select


nC = Val(tNC)
Nr = Val(tNR)

'WW = picLoad.Width
'HH = picLoad.Height
GlobTileW = picLoad.Width / nC 'ww/
GlobTileH = picLoad.Height / Nr 'hh/
'Stop


FM.GlobTileW = GlobTileW
FM.GlobTileH = GlobTileH



If chAllowDuplicates.Value = Unchecked And _
        TotalSourcePhotos < nC * Nr Then MsgBox "You Tried to build a Photomosaic with " _
        & Nr * nC & " Tiles " & vbCrLf & "but you select start Collection(s) with " & TotalSourcePhotos & _
        " photos!" & vbCrLf & vbCrLf & "Cant Continue!" & vbCrLf & vbCrLf & _
        "(You could retry Checking 'Allow Duplicates' Option)", vbExclamation, "Too Few Photos in Collection(s)": Exit Sub


LWait.Visible = True

If nC < 1 Or Nr < 1 Then MsgBox "N Cols or N Rows not valid! (" & nC & " " & Nr & ")", vbExclamation, "not valid input": Exit Sub

Debug.Print "GlobTileW GlobTileH " & GlobTileW & " " & GlobTileH



Select Case FmTYPE
        '-----------------------------------------------------------------------------
    Case "STANDARD"
        FM.NZones = Nr * nC
        myAllocate (TotalSourcePhotos)
        For i = 1 To FM.NZones
            '..................................................................
            With FM.toSEE(i)
                .OnPmWidht = GlobTileW
                .OnPmHeight = GlobTileH
                If .OnPmWidht > .OnPmHeight Then
                    .WZone_W = ComputationalComplexity '6 '12
                    .WZone_H = Round(.OnPmHeight * .WZone_W / .OnPmWidht)
                Else
                    .WZone_H = ComputationalComplexity '6 '12
                    .WZone_W = Round(.OnPmWidht * .WZone_H / .OnPmHeight)
                End If
                Debug.Print "WZone_W WZone_H " & .WZone_W & " " & .WZone_H
            End With
            '................................................................
        Next
        i = 0
        For y = 1 To Nr
            For x = 1 To nC
                i = i + 1
                FM.toSEE(i).CX = (x - 1) * GlobTileW + GlobTileW / 2
                FM.toSEE(i).CY = (y - 1) * GlobTileH + GlobTileH / 2
            Next
        Next
        '-----------------------------------------------------------------------------
        
    Case "CIRCLED_LR", "CIRCLED_UD"
    
    
        FM.NZones = 30 + Nr * nC * 1.5
        myAllocate (TotalSourcePhotos)
        For i = 1 To FM.NZones
            '..................................................................
            With FM.toSEE(i)
                .OnPmWidht = GlobTileW
                .OnPmHeight = GlobTileH
                If .OnPmWidht > .OnPmHeight Then
                    .WZone_W = ComputationalComplexity '6 '12
                    .WZone_H = Round(.OnPmHeight * .WZone_W / .OnPmWidht)
                Else
                    .WZone_H = ComputationalComplexity '6 '12
                    .WZone_W = Round(.OnPmWidht * .WZone_H / .OnPmHeight)
                End If
                Debug.Print "WZone_W WZone_H " & .WZone_W & " " & .WZone_H
            End With
            '................................................................
        Next
        Dim R As Single
        Dim A As Single
        Dim A2 As Single
        
        Dim iW As Single
        Dim iH As Single
        Dim D As Single
        Dim StepR As Single
        Dim StartR As Single
        
'        X = FM.TotalW \ 2
'        Y = FM.TotalH \ 2
        x = FM.TotalW * ((ShapeCenter.left + ShapeCenter.Width / 2) / PicR.Width)
        y = FM.TotalH * ((ShapeCenter.top + ShapeCenter.Height / 2) / PicR.Height)
      
        iW = FM.TotalW
        iH = FM.TotalH
        
        D = Sqr(iW * iW + iH * iH) '/ 2
        
        If Right$(FmTYPE, 2) = "LR" Then
            StepR = (GlobTileW * 0.97)
            StartR = GlobTileW / 2
        Else
            StepR = (GlobTileH * 0.97)
            StartR = GlobTileH / 2
        End If
        
        i = 1
        For R = StartR To D Step StepR
            
            A = 1
            
            Do
                
                With FM.toSEE(i)
                    
                    If Right$(FmTYPE, 2) = "LR" Then
                        .ANG = A
                        If Not (.ANG < 90 Or .ANG > 270) Then
                            .ANG = A - 180: If .ANG < 0 Then .ANG = .ANG + 360
                        End If
                    Else
                        .ANG = A + 90
                        If .ANG > 90 And .ANG < 270 Then
                            .ANG = .ANG - 180: If .ANG < 0 Then .ANG = .ANG + 360
                        End If
                    End If
                    
                    .CX = x + Cos(A * PI2 / 360) * R
                    .CY = y + Sin(A * PI2 / 360) * R
                    
                    If .CX > -.OnPmWidht \ 2 And .CX < FM.TotalW + .OnPmWidht \ 2 And _
                            .CY > -.OnPmHeight \ 2 And .CY < FM.TotalH + .OnPmHeight \ 2 Then i = i + 1
                    
                    If Right$(FmTYPE, 2) = "LR" Then
                        A = A + (GlobTileH / (R + GlobTileW \ 2)) * 57
                    Else
                        A = A + (GlobTileW / (R + GlobTileH \ 2)) * 57
                    End If
                    
                End With
            
            Loop While A < 360
            
        Next R
        
        FM.NZones = i - 1
        ReDim Preserve FM.toSEE(FM.NZones)
        '--------------------------------------------
        
'    Case "ANG_OVERLAP_RND", "ANG_OVERLAP_COL"
'        FM.NZones = Nr * nC * 2.5 + 50
'        myAllocate (TotalSourcePhotos)
'        PicR.Width = FM.TotalW
'        PicR.Height = FM.TotalH
'        PicR.Line (0, 0)-(PicR.Width, PicR.Height), 0, BF
'        PicR.Refresh
'
'        RotaPIC.Cls
'        RotaPIC.Width = GlobTileW
'        RotaPIC.Height = GlobTileH
'        RotaPIC.Line (0, 0)-(RotaPIC.Width, RotaPIC.Height), vbRed, BF
'        RotaPIC.Refresh
'        'Vmin = IIf(GlobTileW < GlobTileH, GlobTileW, GlobTileH) * 0.1
'        'Stop
'
'
'        Stopped = True
'        i = 0
'        Do
'            i = i + 1
''            Stop
'
'            With FM.toSEE(i)
'
'                Cntloop = 0
'                Do
'                    .CX = Rnd * (FM.TotalW - 4) + 2
'                    .CY = Rnd * (FM.TotalH - 4) + 2
'                    '.cx = (.cx \ Vmin) * Vmin
'                    '.cy = (.cy \ Vmin) * Vmin
'
'                    Cntloop = Cntloop + 1
'                Loop While (GetPixel(PicR.hdc, .CX, .CY) = vbRed) And Cntloop < 50000
'
'                .OnPmWidht = GlobTileW
'                .OnPmHeight = GlobTileH
'                If .OnPmWidht > .OnPmHeight Then
'                    .WZone_W = ComputationalComplexity '6 '12
'                    .WZone_H = Round(.OnPmHeight * .WZone_W / .OnPmWidht)
'                Else
'                    .WZone_H = ComputationalComplexity '6 '12
'                    .WZone_W = Round(.OnPmWidht * .WZone_H / .OnPmHeight)
'                End If
'                Debug.Print "WZone_W WZone_H " & .WZone_W & " " & .WZone_H
'
'                .ANG = Rnd * 180 - 90
'                If .ANG < 0 Then .ANG = .ANG + 360
'
'                ROTATE PicR.hdc, .ANG * PI / 180, _
'                    CLng(.CX), CLng(.CY), RotaPIC.Width, RotaPIC.Height, RotaPIC.Image.handle
'                PicR.Refresh
'                DoEvents
'
'            End With
'
'        Loop While Cntloop <> 50000
''        Stop
'
'        FM.NZones = i
'        ReDim Preserve FM.toSEE(FM.NZones)
'        '-----------------------------------------------------------------------------
        
    Case "ANG_OVERLAP_RND", "ANG_OVERLAP_COL"

        Vmin = IIf(GlobTileW < GlobTileH, GlobTileW, GlobTileH) / Sqr(2)
        
        FM.NZones = ((FM.TotalH + Vmin) / Vmin) * ((FM.TotalW + Vmin) / Vmin)
        myAllocate (TotalSourcePhotos)
        
        If Right$(FmTYPE, 3) = "COL" Then
        PicR.Width = FM.TotalW
        PicR.Height = FM.TotalH
        Call SetStretchBltMode(PicR.hdc, STRETCHMODE)
        Call StretchBlt(PicR.hdc, 0, 0, PicR.Width, PicR.Height, _
            picLoad.hdc, 0, 0, picLoad.Width - 1, picLoad.Height - 1, vbSrcCopy)
        PicR.Refresh
        End If
        
        Stopped = True
        
        i = 0
        For y = Vmin \ 2 To FM.TotalH + Vmin \ 2 Step Vmin
        For x = Vmin \ 2 To FM.TotalW + Vmin \ 2 Step Vmin
            i = i + 1
            
            With FM.toSEE(i)
                    
                    .CX = x
                    .CY = y
                
                .OnPmWidht = GlobTileW
                .OnPmHeight = GlobTileH
                If .OnPmWidht > .OnPmHeight Then
                    .WZone_W = ComputationalComplexity '6 '12
                    .WZone_H = Round(.OnPmHeight * .WZone_W / .OnPmWidht)
                Else
                    .WZone_H = ComputationalComplexity '6 '12
                    .WZone_W = Round(.OnPmWidht * .WZone_H / .OnPmHeight)
                End If
                Debug.Print "WZone_W WZone_H " & .WZone_W & " " & .WZone_H
'                                Stop
                                
                If Right$(FmTYPE, 3) = "COL" Then
                    C1 = GetPixel(PicR.hdc, .CX, .CY)
                    Long2RGB C1, R1, B1, G1
                    TR = R1: TG = G1: TB = B1
                    .ANG = -90 + TR / 255 * 60 + TG / 255 * 60 + TB / 255 * 60
                Else
                    .ANG = Rnd * 180 - 90
                End If
                
                If .ANG < 0 Then .ANG = .ANG + 360
               
            End With
            
        Next x
        Next y
        
        
        FM.NZones = i
        ReDim Preserve FM.toSEE(FM.NZones)
        '-----------------------------------------------------------------------------
           
    Case "STND_MASK", "OVERLAP"
        FM.NZones = Nr * nC
        myAllocate (TotalSourcePhotos)
        For i = 1 To FM.NZones
            '..................................................................
            With FM.toSEE(i)
                .OnPmWidht = GlobTileW * 2
                .OnPmHeight = GlobTileH * 2
                If .OnPmWidht > .OnPmHeight Then
                    .WZone_W = ComputationalComplexity '6 '12
                    .WZone_H = Round(.OnPmHeight * .WZone_W / .OnPmWidht)
                Else
                    .WZone_H = ComputationalComplexity '6 '12
                    .WZone_W = Round(.OnPmWidht * .WZone_H / .OnPmHeight)
                End If
                Debug.Print "WZone_W WZone_H " & .WZone_W & " " & .WZone_H
            End With
            '................................................................
        Next
        i = 0
        For y = 1 To Nr
            For x = 1 To nC
                i = i + 1
                FM.toSEE(i).CX = (x - 1) * GlobTileW + GlobTileW / 2
                FM.toSEE(i).CY = (y - 1) * GlobTileH + GlobTileH / 2
            Next
        Next
        '-----------------------------------------------------------------------------

    
    Case "OVERLAP", "OVERLAP_MASK"
        FM.NZones = Nr * nC * IIf(FmTYPE = "OVERLAP", 2.1, 4.2)
        myAllocate (TotalSourcePhotos)
        PicR.Width = FM.TotalW
        PicR.Height = FM.TotalH
        PicR.Line (0, 0)-(PicR.Width, PicR.Height), 0, BF
        PicR.Refresh
        
        Stopped = True
        i = 0
        Do
            i = i + 1
            
            With FM.toSEE(i)
                
                Cntloop = 0
                Do
                    .CX = Rnd * (FM.TotalW - 4) + 2
                    .CY = Rnd * (FM.TotalH - 4) + 2
                    Cntloop = Cntloop + 1
                Loop While (GetPixel(PicR.hdc, .CX, .CY) = vbRed) And Cntloop < 50000
                
                SizeVariation = fnRND(0.8, 1.5, False)
                '                Stop
                
                .OnPmWidht = GlobTileW * SizeVariation
                .OnPmHeight = GlobTileH * SizeVariation
                If .OnPmWidht > .OnPmHeight Then
                    .WZone_W = ComputationalComplexity '6 '12
                    .WZone_H = Round(.OnPmHeight * .WZone_W / .OnPmWidht)
                Else
                    .WZone_H = ComputationalComplexity '6 '12
                    .WZone_W = Round(.OnPmWidht * .WZone_H / .OnPmHeight)
                End If
                Debug.Print "WZone_W WZone_H " & .WZone_W & " " & .WZone_H
                
                If FmTYPE = "OVERLAP" Then
                    PicR.Line (.CX - .OnPmWidht \ 2, .CY - .OnPmHeight \ 2)- _
                            (.CX + .OnPmWidht \ 2, .CY + .OnPmHeight \ 2), vbRed, BF
                Else
                    PicR.Line (.CX - .OnPmWidht \ 3, .CY - .OnPmHeight \ 3)- _
                            (.CX + .OnPmWidht \ 3, .CY + .OnPmHeight \ 3), vbRed, BF
                End If
                PicR.Refresh
                DoEvents
                
            End With
            
        Loop While Cntloop <> 50000
        
        FM.NZones = i
        ReDim Preserve FM.toSEE(FM.NZones)
        
        '------------------------------------------------------------
    Case "ART_1"
        
        
        FM.NZones = Nr * nC * 1.2
        myAllocate (TotalSourcePhotos)
        PicR.Width = FM.TotalW
        PicR.Height = FM.TotalH
        PicR.Line (0, 0)-(PicR.Width, PicR.Height), 0, BF
        PicR.Refresh
        
        Stopped = True
        i = 1
        
        Do
            With FM.toSEE(i)
NewPoint:
                Cntloop = 0
                Do
                    .CX = Int(Rnd * (FM.TotalW - 4) + 2)
                    .CY = Int(Rnd * (FM.TotalH - 4) + 2)
                    X1 = .CX
                    Do
                        X1 = X1 - 1
                    Loop While (GetPixel(PicR.hdc, X1, .CY) = 0) And (Abs(X1 - .CX) / FM.GlobTileW < 1)
                    X2 = .CX
                    Do
                        X2 = X2 + 1
                    Loop While (GetPixel(PicR.hdc, X2, .CY) = 0) And (Abs(X2 - .CX) / FM.GlobTileW < 1)
                    Y1 = .CY
                    Do
                        Y1 = Y1 - 1
                    Loop While (GetPixel(PicR.hdc, .CX, Y1) = 0) And (Abs(Y1 - .CY) / FM.GlobTileH < 1)
                    Y2 = .CY
                    Do
                        Y2 = Y2 + 1
                    Loop While (GetPixel(PicR.hdc, .CX, Y2) = 0) And (Abs(Y2 - .CY) / FM.GlobTileH < 1)
                    
                    Cntloop = Cntloop + 1
                Loop While (GetPixel(PicR.hdc, .CX, .CY) = vbRed) And Cntloop < 50000
                
                
                '          Stop
                
                .CX = (X2 + X1) / 2
                .CY = (Y2 + Y1) / 2
                
                .OnPmWidht = X2 - X1 - 1.5
                .OnPmHeight = Y2 - Y1 - 1.5
                
                If ((GetPixel(PicR.hdc, .CX, .CY) = vbRed) Or _
                        (GetPixel(PicR.hdc, .CX - .OnPmWidht / 2, .CY - .OnPmHeight / 2) = vbRed) Or _
                        (GetPixel(PicR.hdc, .CX + .OnPmWidht / 2, .CY - .OnPmHeight / 2) = vbRed) Or _
                        (GetPixel(PicR.hdc, .CX - .OnPmWidht / 2, .CY + .OnPmHeight / 2) = vbRed) Or _
                        (GetPixel(PicR.hdc, .CX + .OnPmWidht / 2, .CY + .OnPmHeight / 2) = vbRed)) And (Cntloop < 10000) Then GoTo NewPoint
                
                '
                .OnPmWidht = X2 - X1 + 1
                .OnPmHeight = Y2 - Y1 + 1
                
                
                If .OnPmWidht > .OnPmHeight Then
                    .WZone_W = ComputationalComplexity '6 '12
                    .WZone_H = Round(.OnPmHeight * .WZone_W / .OnPmWidht)
                Else
                    .WZone_H = ComputationalComplexity '6 '12
                    .WZone_W = Round(.OnPmWidht * .WZone_H / .OnPmHeight)
                End If
                If .WZone_H < 1 Then .WZone_H = 1 ': Stop
                If .WZone_W < 1 Then .WZone_W = 1 ': Stop
                Debug.Print "WZone_W WZone_H " & .WZone_W & " " & .WZone_H
                
                PicR.Line (.CX - .OnPmWidht \ 2, .CY - .OnPmHeight \ 2)- _
                        (.CX + .OnPmWidht \ 2, .CY + .OnPmHeight \ 2), vbRed, BF
                PicR.Refresh
                DoEvents
                
            End With
            i = i + 1
        Loop While Cntloop <> 50000
        
        FM.NZones = i - 1
        ReDim Preserve FM.toSEE(FM.NZones)
        '-----------------------
        
    Case "ART_brain"
        Dim maxV
        maxV = 200
        
        FM.NZones = Nr * nC * 5
        myAllocate (TotalSourcePhotos)
        PicR.Width = FM.TotalW
        PicR.Height = FM.TotalH
        '        Stop
        
        PicR = LoadPicture(FM.FilePathIN)
        
        'PicR.Line (0, 0)-(PicR.Width, PicR.Height), 0, BF
        PicR.Refresh
        
        Stopped = True
        i = 1
        
        Do
            With FM.toSEE(i)
NewPoint2:
                Cntloop = 0
                Do
                    .CX = Int(Rnd * (FM.TotalW - 4) + 2)
                    .CY = Int(Rnd * (FM.TotalH - 4) + 2)
                    
                    dt = 0
                    X1 = .CX
                    Do
                        '                       Stop
                        
                        X1 = X1 - 1
                        C1 = GetPixel(PicR.hdc, X1, .CY)
                        Long2RGB C1, R1, G1, B1
                        If Abs(X1 - .CX) <> 1 Then GoSub CalcolaDT
                        '
                        R2 = R1
                        G2 = G1
                        B2 = B1
                        
                    Loop While (C1 <> vbRed) And (dt < maxV) And (Abs(X1 - .CX) / FM.GlobTileW < 1)
                    
                    
                    dt = 0
                    X2 = .CX
                    Do
                        X2 = X2 + 1
                        C1 = GetPixel(PicR.hdc, X2, .CY)
                        Long2RGB C1, R1, G1, B1
                        If Abs(X2 - .CX) <> 1 Then GoSub CalcolaDT
                        '
                        '
                        R2 = R1
                        G2 = G1
                        B2 = B1
                        
                        
                    Loop While (C1 <> vbRed) And (dt < maxV) And (Abs(X2 - .CX) / FM.GlobTileW < 1)
                    
                    dt = 0
                    Y1 = .CY
                    Do
                        Y1 = Y1 - 1
                        C1 = GetPixel(PicR.hdc, .CX, Y1)
                        Long2RGB C1, R1, G1, B1
                        If Abs(Y1 - .CY) <> 1 Then GoSub CalcolaDT
                        '
                        
                        R2 = R1
                        G2 = G1
                        B2 = B1
                        
                    Loop While (C1 <> vbRed) And (dt < maxV) And (Abs(Y1 - .CY) / FM.GlobTileH < 1)
                    
                    dt = 0
                    Y2 = .CY
                    Do
                        Y2 = Y2 + 1
                        C1 = GetPixel(PicR.hdc, .CX, Y2)
                        Long2RGB C1, R1, G1, B1
                        If Abs(Y2 - .CY) <> 1 Then GoSub CalcolaDT
                        '
                        
                        R2 = R1
                        G2 = G1
                        B2 = B1
                        
                    Loop While (C1 <> vbRed) And (dt < maxV) And (Abs(Y2 - .CY) / FM.GlobTileH < 1)
                    
                    Cntloop = Cntloop + 1
                    
                    'Stop
                    
                Loop While (GetPixel(PicR.hdc, .CX, .CY) = vbRed) And Cntloop < 50000
                
                
                '                         Stop
                
                .CX = (X2 + X1) / 2
                .CY = (Y2 + Y1) / 2
                
                .OnPmWidht = X2 - X1 - 2 '1.5
                .OnPmHeight = Y2 - Y1 - 2 '1.5
                
                If ((GetPixel(PicR.hdc, .CX, .CY) = vbRed) Or _
                        (GetPixel(PicR.hdc, .CX - .OnPmWidht / 2, .CY - .OnPmHeight / 2) = vbRed) Or _
                        (GetPixel(PicR.hdc, .CX + .OnPmWidht / 2, .CY - .OnPmHeight / 2) = vbRed) Or _
                        (GetPixel(PicR.hdc, .CX - .OnPmWidht / 2, .CY + .OnPmHeight / 2) = vbRed) Or _
                        (GetPixel(PicR.hdc, .CX + .OnPmWidht / 2, .CY + .OnPmHeight / 2) = vbRed)) And (Cntloop < 10000) Then GoTo NewPoint2
                
                '
                .OnPmWidht = X2 - X1 + 1
                .OnPmHeight = Y2 - Y1 + 1
                
                
                If .OnPmWidht > .OnPmHeight Then
                    .WZone_W = ComputationalComplexity '6 '12
                    .WZone_H = Round(.OnPmHeight * .WZone_W / .OnPmWidht)
                Else
                    .WZone_H = ComputationalComplexity '6 '12
                    .WZone_W = Round(.OnPmWidht * .WZone_H / .OnPmHeight)
                End If
                If .WZone_H < 1 Then .WZone_H = 1 ': Stop
                If .WZone_W < 1 Then .WZone_W = 1 ': Stop
                Debug.Print "WZone_W WZone_H " & .WZone_W & " " & .WZone_H
                
                PicR.Line (.CX - .OnPmWidht \ 2, .CY - .OnPmHeight \ 2)- _
                        (.CX + .OnPmWidht \ 2, .CY + .OnPmHeight \ 2), vbRed, BF
                PicR.Refresh
                DoEvents
                
            End With
            i = i + 1
        Loop While Cntloop <> 50000
        
        FM.NZones = i - 1
        ReDim Preserve FM.toSEE(FM.NZones)
        '-----------------------
        
End Select

'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------







FM.FilePathOUT = App.Path & "\MOSAIC\" & left$(DopoUltimaBarra(FM.FilePathIN), Len(DopoUltimaBarra(FM.FilePathIN)) - 4) & "_PM_" & FmTYPE & "_" & FM.NZones
'..............


STATUS = "Creating Photomosaic..."
DoEvents
'--------------------------------------------------------------------------
'---------------------------------------------------------------------------
'----------------------------------------------------------------------------
'Stop

i = 1
'FM.toSEE(i).cx = (GlobTileW / 2) / FM.MaskPercX ''mettere uguale anche su FOR "WWW"
'FM.toSEE(i).cy = (GlobTileH / 2) / FM.MaskPercY


Dim ix As Single
Dim iy As Single
Dim id As Single
Dim idSmall As Single

Dim iX2 As Single
Dim iY2 As Single
Dim iD2 As Single



For i = 1 To FM.NZones
    
    PicR.Width = FM.toSEE(i).WZone_W
    PicR.Height = FM.toSEE(i).WZone_H
    
    If FM.toSEE(i).ANG = 0 Then
        
        Call SetStretchBltMode(PicR.hdc, STRETCHMODE)
        Call StretchBlt(PicR.hdc, 0, 0, PicR.Width, PicR.Height, _
                picLoad.hdc, _
                FM.toSEE(i).CX - FM.toSEE(i).OnPmWidht \ 2, _
                FM.toSEE(i).CY - FM.toSEE(i).OnPmHeight \ 2, _
                FM.toSEE(i).OnPmWidht, _
                FM.toSEE(i).OnPmHeight, _
                vbSrcCopy)
        PicR.Refresh
        
    Else
        '        Stop
        
        ix = FM.toSEE(i).OnPmWidht
        iy = FM.toSEE(i).OnPmHeight
        id = Sqr(ix * ix + iy * iy)
        idSmall = id / 2
      
        RotaPIC.Width = idSmall
        RotaPIC.Height = idSmall
        
        
        'Mette Porzione PicLoad In RotaPIC
        Call SetStretchBltMode(RotaPIC.hdc, STRETCHMODE)
        Call StretchBlt(RotaPIC.hdc, 0, 0, idSmall, idSmall, _
                picLoad.hdc, _
                FM.toSEE(i).CX - id \ 2, _
                FM.toSEE(i).CY - id \ 2, _
                id, _
                id, _
                vbSrcCopy)
        RotaPIC.Refresh
           
           
'          Stop
          
    ROTATE RotaPIC.hdc, FM.toSEE(i).ANG * PI / 180, _
            idSmall \ 2, idSmall \ 2, _
            CLng(idSmall), CLng(idSmall), RotaPIC.Image.handle
    
    RotaPIC.Refresh
    
'    Stop
    

 Call SetStretchBltMode(PicR.hdc, STRETCHMODE)
        Call StretchBlt(PicR.hdc, 0, 0, PicR.Width, PicR.Height, _
                RotaPIC.hdc, _
                idSmall \ 2 - FM.toSEE(i).OnPmWidht \ 4, _
                idSmall \ 2 - FM.toSEE(i).OnPmHeight \ 4, _
                FM.toSEE(i).OnPmWidht \ 2, _
                FM.toSEE(i).OnPmHeight \ 2, _
                vbSrcCopy)
        PicR.Refresh
    
'Stop

        
    End If
    
    'Me.Refresh
    '''''''''''''''''''''''''''''''' legge colori zone
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    For YY = 0 To PicR.Height - 1
        For XX = 0 To PicR.Width - 1
            ccc = GetPixel(PicR.hdc, XX, YY)
            Long2RGB ccc, Cr, cG, cb
            FM.toSEE(i).R(XX, YY) = Cr
            FM.toSEE(i).G(XX, YY) = cG
            FM.toSEE(i).B(XX, YY) = cb
        Next XX
    Next YY
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''
Next i

'--------------------------------------------------------------------------'
'--------------------------------------------------------------------------'
'Stop

Computazione

SaveFM

Beep
Ricostruz
Beep
Exit Sub
'----------------------------------------------------
'bRAIN
CalcolaDT:

DR = Abs(R1 \ 1 - R2 \ 1)
DG = Abs(G1 \ 1 - G2 \ 1)
DB = Abs(B1 \ 1 - B2 \ 1)
mR = R1 \ 2 + R2 \ 2

dt = dt + ( _
        FASTColorDistanceR(DR, mR) + _
        FASTColorDistanceG(DG) + _
        FASTColorDistanceB(DB, mR) _
        )




Return
'----------------------------------------------------
End Sub


Sub myAllocate(TotalSourcePhotos)
STATUS = "Allocating memory..."
DoEvents
ReDim FM.toSEE(FM.NZones + 1)
For i = 1 To FM.NZones
    ReDim FM.toSEE(i).FIT(TotalSourcePhotos)
    ReDim FM.toSEE(i).agR(TotalSourcePhotos)
    ReDim FM.toSEE(i).agG(TotalSourcePhotos)
    ReDim FM.toSEE(i).agB(TotalSourcePhotos)
    ReDim FM.toSEE(i).FitINDEXusable(TotalSourcePhotos)
    DoEvents
Next i
'INIT
For Z = 1 To FM.NZones
    For i = 1 To TotalSourcePhotos
        FM.toSEE(Z).FitINDEXusable(i) = True
        FM.toSEE(Z).FIT(i) = 1E+20
        FM.toSEE(Z).indexBESTFIT = 0
        FM.toSEE(Z).indexBFfileName = vbNullString
    Next i
Next Z
End Sub
Private Function Atan2(ByVal dx As Single, ByVal dy As Single) As Single
'This Should return Angle
'Stop

Dim theta As Single

If (Abs(dx) < 0.0000001) Then
    If (Abs(dy) < 0.0000001) Then
        theta = 0#
    ElseIf (dy > 0#) Then
        theta = 1.5707963267949
        'theta = PI / 2
    Else
        theta = -1.5707963267949
        'theta = -PI / 2
    End If
Else
    theta = Atn(dy / dx)
    
    If (dx < 0) Then
        If (dy >= 0#) Then
            theta = PI + theta
        Else
            theta = theta - PI
        End If
    End If
End If

Atan2 = theta
End Function
