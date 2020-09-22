VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Scrollbars A++ Example"
   ClientHeight    =   7290
   ClientLeft      =   420
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      SmallChange     =   10
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6720
      Width           =   6255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5535
      Left            =   6360
      SmallChange     =   10
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   6615
      Left            =   0
      ScaleHeight     =   437
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   413
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   8640
         Left            =   0
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   8640
         ScaleWidth      =   5895
         TabIndex        =   3
         Top             =   0
         Width           =   5895
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hi everyone, this is my first submission to PSC, hope
'you all get what you need
'
'I made this because the examples on PSC wernt clear enough for
'new coders, or the examples were bundled into other code.
'This should hopefully answer any questions about how to make
'scrollbars and have them work correctly.
'
'Note: My form and the picturebox's have a scalemode of PIXEL, I dont
'      know why they have TwipsPerPixel, but hey it works
'
'
'You are free to use this code where ever you want, its just text!!

Private Sub Form_Load()
    'This calls the form_resize event
    Form_Resize
End Sub

Private Sub Form_Resize()


    'These are just condidtion to exit the function
    'You can make things smaller then 0
    'and there is no point in resizing if the window
    'is minimized
    
    'Windowstate = 1 This is if the window is minimezed,2 is maximised
    If WindowState = 1 Then Exit Sub
    If Form1.ScaleWidth < 20 Then Exit Sub
    If Form1.ScaleHeight < 20 Then Exit Sub
    
    
    'This sets the ammount the scrollbars can move and
    'the size of the scroller
    VScroll1.Max = Picture2.Height - Picture1.Height
    VScroll1.LargeChange = Form1.ScaleHeight
    
    'This sets whether the scroll bars are visible
    If VScroll1.Max < 0 Then
        VScroll1.Visible = False
        Picture1.Width = Form1.ScaleWidth
    Else
        VScroll1.Visible = True
        Picture1.Width = Form1.ScaleWidth - 17
    End If
    
    'This sets the ammount the scrollbars can move and
    'the size of the scroller
    HScroll1.Max = Picture2.Width - Picture1.Width
    HScroll1.LargeChange = Form1.ScaleWidth
    
    'This sets whether the scroll bars are visible
    If HScroll1.Max < 0 Then
        HScroll1.Visible = False
        Picture1.Height = Form1.ScaleHeight
    Else
        HScroll1.Visible = True
        Picture1.Height = Form1.ScaleHeight - 17
    End If
    
    'This connects the scroll bars to the edges of the form
    VScroll1.Top = 0
    VScroll1.Left = Form1.ScaleWidth - VScroll1.Width
    
    HScroll1.Top = Form1.ScaleHeight - HScroll1.Height
    HScroll1.Left = 0

    'This part sets the horizontal scrollbars width
    'depending on whether the vertical scrollbar is visble
    If HScroll1.Visible = True Then
        VScroll1.Height = Form1.ScaleHeight - 17
    Else
        VScroll1.Height = Form1.ScaleHeight
    End If
    
    
    'This part sets the vertical scrollbars height
    'depending on whether the horizontal scrollbar is visble
    If VScroll1.Visible = True Then
        HScroll1.Width = Form1.ScaleWidth - 17
    Else
        HScroll1.Width = Form1.ScaleWidth
    End If
End Sub

Private Sub HScroll1_Change()
    'Updates picture2.left location when you
    'click to scroll the scrollbars
    Picture2.Left = 0 - (HScroll1.Value)
End Sub

Private Sub HScroll1_Scroll()
    'Updates picture2.left location when you
    'grab and scroll the scrollbars
    Picture2.Left = 0 - (HScroll1.Value)
End Sub

Private Sub Picture2_Click()

    'This just lets you load a picture into picture2
    'when you click it once
    
    'This code might be usefull to new coders
    CommonDialog1.Filter = "*.bmp|*.bmp|*.jpeg|*.jpeg|*.jpg|*.jpg|*.gif|*.gif|All Picture Files|*.bmp;*.jpeg;*.jpg;*.gif"
    CommonDialog1.FilterIndex = 5
    
    'When the commondialog is open, the program waits for
    'its returned input
    CommonDialog1.ShowOpen
    Picture2.Picture = LoadPicture(CommonDialog1.FileName)
    
    'call the resize function to set the scrollbars
    Form_Resize
End Sub

Private Sub VScroll1_Change()
    'Updates picture2.top location when you
    'click to scroll the scrollbars
    Picture2.Top = 0 - (VScroll1.Value)
End Sub

Private Sub VScroll1_Scroll()
    'Updates picture2.top location when you
    'grab and scroll the scrollbars
    Picture2.Top = 0 - (VScroll1.Value)
End Sub
