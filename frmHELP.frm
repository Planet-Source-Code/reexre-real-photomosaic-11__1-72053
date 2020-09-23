VERSION 5.00
Begin VB.Form frmHELP 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4230
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   5520
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   368
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox TXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   3975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmHELP.frx":0000
      Top             =   120
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   4215
      Left            =   45
      Top             =   0
      Width           =   5400
   End
End
Attribute VB_Name = "frmHELP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim S As String

'Me.Width = Me.Height * 1.618
'TXT.Width = TXT.Height * 1.618
'Shape1.Width = Shape1.Height * 1.618


Me.Height = Me.Width * 1.618
TXT.Height = TXT.Width * 1.618
Shape1.Width = Me.ScaleWidth
Shape1.left = 0
Shape1.Height = Shape1.Width * 1.618


H(1) = vbCrLf & "C R E A T E  C O L L E C T I O N" & vbCrLf & vbCrLf
H(1) = H(1) & "Collection is a Collection of Photos to use as tiles." & vbCrLf
H(1) = H(1) & "Click Here, It is prompted to choose a folder, then each jpg picture in that folder will be scanned and a Collection 'FolderName' will be created." & vbCrLf
H(1) = H(1) & vbCrLf
H(1) = H(1) & "This is the first step to do if Collections List is empty." & vbCrLf

H(2) = vbCrLf & "L O A D   (checked)   C O L L E C T I O N" & vbCrLf & vbCrLf
H(2) = H(2) & "Click Here to Load in Memory the selected Collection(s) that will be used in Creation Process." & vbCrLf & vbCrLf
H(2) = H(2) & "If your Collection List is empty go to point (1)." & vbCrLf & vbCrLf

H(3) = vbCrLf & "L O A D   S U B J E C T   P I C" & vbCrLf & vbCrLf
H(3) = H(3) & "Click here to Select a Picture to Use as SUBJECT" & vbCrLf & vbCrLf


H(4) = vbCrLf & "C R E A T E   P H O T O M O S A I C" & vbCrLf & vbCrLf
H(4) = H(4) & "Click here to Start Photomosaic Creation Process." & vbCrLf & vbCrLf
H(4) = H(4) & "(At least 1 Collection Must be Created and Loaded)" & vbCrLf & vbCrLf



H(5) = vbCrLf & "O P E N   A N D   R E B U I L D" & vbCrLf & vbCrLf
H(5) = H(5) & "With here left OUTPUTSIZE SlideBar Choose Output Dimension of Photomosaic." & vbCrLf & vbCrLf
H(5) = H(5) & "Can Modify 'Adjust Color' and 'BLEND' Values too." & vbCrLf & vbCrLf

H(5) = H(5) & "Click Here to Select wich (perivous Created) Photomosaic you want Re-Build and Start ReBuild Process." & vbCrLf & vbCrLf

H(6) = vbCrLf & "C O L L E C T I O N S  L I S T" & vbCrLf & vbCrLf
H(6) = H(6) & "This is the Collections List. Here you can Select/Deselect Collections involved in Creation Process." & vbCrLf & vbCrLf
H(6) = H(6) & "When choise is done go to Point (2) to Load it/them in memory" & vbCrLf & vbCrLf

H(7) = vbCrLf & "ALLOW DUPLICATES / SELECT DISTANCE" & vbCrLf & vbCrLf
H(7) = H(7) & "Allow or Deny Same Tile to appear Multiple Times." & vbCrLf & vbCrLf
H(7) = H(7) & "If Duplicates are Allowed then Type Minimal 'number of Tiles'(distance) Between Identical Tiles." & vbCrLf & vbCrLf

H(8) = vbCrLf & "ADJUST COLORS" & vbCrLf & vbCrLf
H(8) = H(8) & "With SlideBar choose % of colors adjustment in Creation Process." & vbCrLf & vbCrLf
H(8) = H(8) & "0% means: tiles Color is not changed." & vbCrLf
H(8) = H(8) & "100% means: tiles Color is full changed to best match the tile Zone." & vbCrLf

H(9) = vbCrLf & "FAST (Less Accurate)" & vbCrLf & vbCrLf
H(9) = H(9) & "If Checked then Creation Process is (about 4 times) Faster but Less Accurate." & vbCrLf & vbCrLf

H(10) = vbCrLf & "OUTPUT SIZE" & vbCrLf & vbCrLf
H(10) = H(10) & "10 Means: Output Size is equal to Subject Picture Size." & vbCrLf
H(10) = H(10) & "20 Means: Output Size is Two time the Subject Picture Size." & vbCrLf & vbCrLf
H(10) = H(10) & "... and so on..." & vbCrLf & vbCrLf
H(10) = H(10) & "Centimeters size is calculate as 300 DPI image." & vbCrLf & vbCrLf


H(11) = vbCrLf & "UPDATE COLLECTION(S)" & vbCrLf & vbCrLf
H(11) = H(11) & "If you added/removed/modified photos in some Collection Folder then it's needed to Update Collection." & vbCrLf
H(11) = H(11) & "Go to Collection List and Check the Collection(s) you want to update." & vbCrLf
H(11) = H(11) & "Click Here and Checked Collection(s) will be updated" & vbCrLf & vbCrLf
H(11) = H(11) & "(For added Photos works Good, otherwise may not work. Alternatively You can Go to Point (1) and reCreate Collection from zero.)"

H(12) = vbCrLf & "LOAD SET   /   SAVE SET" & vbCrLf & vbCrLf
H(12) = H(12) & "A Collection Set is a List of Collections." & vbCrLf & vbCrLf
H(12) = H(12) & "Load: Load  Collection Set." & vbCrLf & vbCrLf
H(12) = H(12) & "Save: Check Collections to Group as a 'SET' and then click Save Set." & vbCrLf & vbCrLf


H(13) = vbCrLf & "MIRRORED TILES" & vbCrLf & vbCrLf
H(13) = H(13) & "If Checked , Consider and use even Mirrored Tiles in Creation Process." & vbCrLf
H(13) = H(13) & "So the Number of Photos to Use as Tiles is  Doubled." & vbCrLf

H(14) = vbCrLf & "P H O T O M O S A I C   T Y P E" & vbCrLf & vbCrLf
H(14) = H(14) & "STANDARD = Tiles are placed side by side" & vbCrLf & vbCrLf
H(14) = H(14) & "STND_MASK = Tiles are MASKED. Click a picture here Right to choose the Mask. The 3RD picture is suggested." & vbCrLf & vbCrLf
H(14) = H(14) & "OVERLAP = Tiles are Overlapped and have Random size. (Total Tiles Numbers is about 3 Times than dispalyed)" & vbCrLf & vbCrLf
H(14) = H(14) & "OVERLAP_MASK = Tiles are Overlapped With MASK and have Random size. (Total Tiles Numbers is about 3 Times than dispalyed)" & vbCrLf & vbCrLf
H(14) = H(14) & "ART_1 = Random Size NOT OVERLAPPED" & vbCrLf & vbCrLf
H(14) = H(14) & "ART_brain = Similar to ART_1. Requires a lot of memory" & vbCrLf & vbCrLf
H(14) = H(14) & "CIRCLED_LR = Tiles are Placed in Circles. Click Point on subject Picture to change circle Center. Orizontal Tiles are Left And Right. (Total Tiles Numbers is about 1.2-1.5 Times than dispalyed)" & vbCrLf & vbCrLf
H(14) = H(14) & "CIRCLED_UD = Same as above but Orizontal Tiles are Up and Down" & vbCrLf & vbCrLf
H(14) = H(14) & "ANG_OVERLAP_RND = Tiles are overlapped with Random Angle. (Total Tiles Numbers is about 3 Times than dispalyed)" & vbCrLf & vbCrLf
H(14) = H(14) & "ANG_OVERLAP_COL = Same as above but the Angle is given by subject pic Color." & vbCrLf & vbCrLf


H(15) = vbCrLf & "BLEND" & vbCrLf & vbCrLf
H(15) = H(15) & "With SlideBar choose % of how much Subject Picture Image will appear in Photomosaic." & vbCrLf & vbCrLf


H(16) = vbCrLf & "G E T   R A N D O M   P I C   A S   S U B J E C T" & vbCrLf & vbCrLf
H(16) = H(16) & "Click here to Select a Random Picture From Loaded Collection(s) to Use as SUBJECT" & vbCrLf & vbCrLf
H(16) = H(16) & "(at Least 1 Collection must be loaded, [ See(2) ] )" & vbCrLf



S = "R E A L  P H O T O M O S A I C " & vbTab & App.Major & vbCrLf & vbCrLf & vbCrLf & vbCrLf
For i = 1 To UBound(H)
S = S & H(i) & "---------------------------------------------------------" & vbCrLf & vbCrLf
Next
Open App.Path & "\HELP.txt" For Output As 55
Print #55, S
Close 55


End Sub

