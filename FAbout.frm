VERSION 5.00
Begin VB.Form FAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About HTMLLabel"
   ClientHeight    =   3435
   ClientLeft      =   4755
   ClientTop       =   3705
   ClientWidth     =   5220
   Icon            =   "FAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin HTMLLabelEdit.HTMLLabel ctlHTML 
      Height          =   2445
      Left            =   900
      TabIndex        =   2
      Top             =   180
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   4313
      Appearance      =   0
      BorderStyle     =   0
      BackColor       =   -2147483633
      EnableAnchors   =   -1  'True
      EnableScroll    =   0   'False
      DefaultFontName =   "Tahoma"
      DefaultFontSize =   8
   End
   Begin VB.PictureBox picLogo 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   30
      Picture         =   "FAbout.frx":000C
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   2910
      Width           =   1215
   End
End
Attribute VB_Name = "FAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   Form FAbout.
'
'   About... box.
'

'
' cmdOK_Click
'
' Dismiss the dialog.
'
Private Sub cmdOK_Click()
    Unload Me
End Sub
'
' ctlHTML_HyperlinkClick()
'
' Follow any clicked hyperlinks.
'
Private Sub ctlHTML_HyperlinkClick(Href As String)
    Shell "start " & Href
End Sub
'
' Form_Load
'
' Initialisation.
'
Private Sub Form_Load()
    ' Set screen display.
    ctlHTML.DocumentHTML = "<html><body>" & _
                           "<h2><b>HTMLLabel</b></h2>" & _
                           "<font face='tahoma' size='3'>Version: " & ctlHTML.Version & "<br><hr>" & _
                           "<p>For further information, version updates and other products, " & _
                           "<a href='http://www.damnet.freeserve.co.uk/products/'>" & _
                           "visit our products web site</a>.</p><br><hr>" & _
                           "<p>Copyright &copy; 2001 <a href='http://www.woodbury.co.uk'>" & _
                           "Woodbury Associates</a></p>" & _
                           "</font></body></html>"
End Sub
'
' Form_Unload
'
Private Sub Form_Unload(Cancel As Integer)
    Set mfrmParent = Nothing
End Sub
'
' Form_Resize()
'
' Refresh the HTMLLabel control.
'
Private Sub Form_Resize()
    ctlHTML.Refresh False
End Sub

