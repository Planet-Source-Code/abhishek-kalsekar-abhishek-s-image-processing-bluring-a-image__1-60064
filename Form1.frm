VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Image Bluring"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   3315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBlur 
      Caption         =   "Blur Image"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4710
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   312
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   120
      Width           =   3030
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Type ColorRGBType
    Red As Integer
    Blue As Integer
    Green As Integer
End Type


Private Function RgbColor(Color As Long) As ColorRGBType
    RgbColor.Red = (Int(Color And 255)) And 255
    RgbColor.Green = (Int(Color / 256)) And 255
    RgbColor.Blue = (Int(Color / 65536)) And 255
End Function


Private Sub cmdBlur_Click()
    Dim x As Integer
    Dim y As Integer
    Dim Pixel As Long
    Dim Pixel2 As Long
    Dim Col As ColorRGBType
    
   

    With P
        Me.Caption = "Image Bluring - Processing"
        For x = 0 To .ScaleWidth - 1
            For y = 0 To .ScaleHeight - 1
            
                Pixel = GetPixel(.HDC, x, y)
                
                If x < .ScaleWidth - 3 Then
                    Pixel2 = GetPixel(.HDC, x + 2, y)
                End If
                
                Col.Red = (RgbColor(Pixel).Red + RgbColor(Pixel2).Red) / 2
                Col.Green = (RgbColor(Pixel).Green + RgbColor(Pixel2).Green) / 2
                Col.Blue = (RgbColor(Pixel).Blue + RgbColor(Pixel2).Blue) / 2
                
                SetPixelV .HDC, x + 1, y, RGB(Col.Red, Col.Green, Col.Blue)
            
            Next y
            .Refresh
        Next x
        
    End With
    Me.Caption = "Image Bluring"
    
End Sub

