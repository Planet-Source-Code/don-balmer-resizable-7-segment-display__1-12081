VERSION 5.00
Begin VB.UserControl Rsz_7Seg 
   BackColor       =   &H00000000&
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10185
   ScaleHeight     =   6975
   ScaleWidth      =   10185
   Begin VB.Timer tmrTimeBase_Sub 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   27
      Left            =   2280
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   27
      Left            =   1800
      Shape           =   2  'Oval
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   27
      Left            =   3720
      Shape           =   2  'Oval
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   26
      Left            =   1440
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   26
      Left            =   1440
      Shape           =   2  'Oval
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   26
      Left            =   1440
      Shape           =   2  'Oval
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   25
      Left            =   1440
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   25
      Left            =   1440
      Shape           =   2  'Oval
      Top             =   3720
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   25
      Left            =   1440
      Shape           =   2  'Oval
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   24
      Left            =   2280
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   24
      Left            =   3840
      Shape           =   2  'Oval
      Top             =   6480
      Width           =   735
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   24
      Left            =   1920
      Shape           =   2  'Oval
      Top             =   6480
      Width           =   735
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   23
      Left            =   4440
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   23
      Left            =   4440
      Shape           =   2  'Oval
      Top             =   3720
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   23
      Left            =   4440
      Shape           =   2  'Oval
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   22
      Left            =   4440
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   22
      Left            =   4440
      Shape           =   2  'Oval
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   22
      Left            =   4440
      Shape           =   2  'Oval
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   21
      Left            =   2160
      Top             =   0
      Width           =   1935
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   21
      Left            =   1800
      Shape           =   2  'Oval
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   21
      Left            =   3720
      Shape           =   2  'Oval
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   20
      Left            =   7080
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   20
      Left            =   6600
      Shape           =   2  'Oval
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   20
      Left            =   8520
      Shape           =   2  'Oval
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   19
      Left            =   6240
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   19
      Left            =   6240
      Shape           =   2  'Oval
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   19
      Left            =   6240
      Shape           =   2  'Oval
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   18
      Left            =   6240
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   18
      Left            =   6240
      Shape           =   2  'Oval
      Top             =   3720
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   18
      Left            =   6240
      Shape           =   2  'Oval
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   17
      Left            =   7080
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   17
      Left            =   8640
      Shape           =   2  'Oval
      Top             =   6480
      Width           =   735
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   17
      Left            =   6720
      Shape           =   2  'Oval
      Top             =   6480
      Width           =   735
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   16
      Left            =   9240
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   16
      Left            =   9240
      Shape           =   2  'Oval
      Top             =   3720
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   16
      Left            =   9240
      Shape           =   2  'Oval
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   15
      Left            =   9240
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   15
      Left            =   9240
      Shape           =   2  'Oval
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   15
      Left            =   9240
      Shape           =   2  'Oval
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   14
      Left            =   6960
      Top             =   0
      Width           =   1935
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   14
      Left            =   8520
      Shape           =   2  'Oval
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   14
      Left            =   6600
      Shape           =   2  'Oval
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape shpCOLON_LEFT 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   220
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape shpCOLON_LEFT 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   220
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape shpCOLON_LEFT 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   220
      Shape           =   3  'Circle
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape shpCOLON_LEFT 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   220
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   13
      Left            =   3720
      Shape           =   2  'Oval
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   13
      Left            =   1800
      Shape           =   2  'Oval
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   13
      Left            =   2160
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   12
      Left            =   1440
      Shape           =   2  'Oval
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   12
      Left            =   1440
      Shape           =   2  'Oval
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   12
      Left            =   1440
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   11
      Left            =   1440
      Shape           =   2  'Oval
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   11
      Left            =   1440
      Shape           =   2  'Oval
      Top             =   3720
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   11
      Left            =   1440
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   10
      Left            =   1920
      Shape           =   2  'Oval
      Top             =   6480
      Width           =   735
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   10
      Left            =   3840
      Shape           =   2  'Oval
      Top             =   6480
      Width           =   735
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   10
      Left            =   2280
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   9
      Left            =   4440
      Shape           =   2  'Oval
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   9
      Left            =   4440
      Shape           =   2  'Oval
      Top             =   3720
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   9
      Left            =   4440
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   8
      Left            =   4440
      Shape           =   2  'Oval
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   8
      Left            =   4440
      Shape           =   2  'Oval
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   8
      Left            =   4440
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   7
      Left            =   3720
      Shape           =   2  'Oval
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   7
      Left            =   1800
      Shape           =   2  'Oval
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   7
      Left            =   2160
      Top             =   0
      Width           =   1935
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   6
      Left            =   8520
      Shape           =   2  'Oval
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   6
      Left            =   6600
      Shape           =   2  'Oval
      Top             =   3240
      Width           =   735
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   6
      Left            =   6960
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   5
      Left            =   6240
      Shape           =   2  'Oval
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   5
      Left            =   6240
      Shape           =   2  'Oval
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   5
      Left            =   6240
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   4
      Left            =   6240
      Shape           =   2  'Oval
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   4
      Left            =   6240
      Shape           =   2  'Oval
      Top             =   3720
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   4
      Left            =   6240
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   6720
      Shape           =   2  'Oval
      Top             =   6480
      Width           =   735
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   8640
      Shape           =   2  'Oval
      Top             =   6480
      Width           =   735
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   7080
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   2
      Left            =   9240
      Shape           =   2  'Oval
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   2
      Left            =   9240
      Shape           =   2  'Oval
      Top             =   3720
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   2
      Left            =   9240
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   1
      Left            =   9240
      Shape           =   2  'Oval
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   1
      Left            =   9240
      Shape           =   2  'Oval
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   1
      Left            =   9240
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape shpSEG_END2 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   8520
      Shape           =   2  'Oval
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape shpSEG_END1 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   6600
      Shape           =   2  'Oval
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape shpSEG_MID 
      BorderColor     =   &H00000080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   6960
      Top             =   0
      Width           =   1935
   End
   Begin VB.Image imgBACKDROP 
      Height          =   6975
      Left            =   0
      Top             =   0
      Width           =   10185
   End
End
Attribute VB_Name = "Rsz_7Seg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Dim gColons_Overide As Integer
'Default Property Values:
Const m_def_ColonsState = 1
Const m_def_Value = 0
Const m_def_BlankMSB = 0
Const m_def_MaxCount = 99
Const m_def_MinCount = 0
Const m_def_CountStep = 1
Const m_def_ForceColons = 0
Const m_def_DisableColons = 0
Const m_def_ActiveColour = &HFF&
Const m_def_Backcolour = &H0&
Const m_def_InactiveColour = &H40&
Const m_def_BackTransparent = 0
Const m_def_AutoTimebase = 1000
'Property Variables:
Dim m_ColonsState As Boolean
Dim m_value As Long
Dim m_BlankMSB As Boolean
Dim m_Maxcount As Integer
Dim m_Mincount As Integer
Dim m_CountStep As Long
Dim m_ForceColons As Boolean
Dim m_DisableColons As Boolean
Dim m_ActiveColour As OLE_COLOR
Dim m_Backcolour As OLE_COLOR
Dim m_InactiveColour As OLE_COLOR
Dim m_BackTransparent As Boolean
Dim m_AutoTimebase As Integer
'Event Declarations:
Event Overflow(Carry_Out)
Attribute Overflow.VB_Description = "When trigger cause .Value to exceed .CountMax the difference is carried "
Event Underflow(Carry_Out)
Attribute Underflow.VB_Description = "When trigger cause .Value todrop below  .CountMin the difference is carried "
Event ColonsChangeState(State As Boolean)


'
'During control resizing it is necesary to recalculate size and position
'of segment elements (Inactive and Active) and then redraw. This can be ugly
'so the segments are blanked during resize and are only redrawn when
'scaling is complete
'
Private Sub Display_Inactive_Elements(State_Flag As Boolean)

Dim Visi_Flag As Boolean
Dim Visi_Loop As Integer
Dim End_Element As Integer

'Last element of the Inactive set of the elements array
End_Element = 13

If State_Flag = True Then Visi_Flag = True

If m_BlankMSB = True And Visi_Flag = True Then End_Element = 6


'loop through element and set to value of State_Flag
For Visi_Loop = 0 To End_Element

'Segments are made up of three elementals
'
'the left hand 'Curve'
UserControl.shpSEG_END1(Visi_Loop).Visible = Visi_Flag
'the centre 'Bar'
UserControl.shpSEG_MID(Visi_Loop).Visible = Visi_Flag
'the right hand 'Curve'
UserControl.shpSEG_END2(Visi_Loop).Visible = Visi_Flag

Next Visi_Loop

'The colons are made of elements to and require blanking etc.
'but the control provides an 'Overide' for colons so they are never shown
'or always shown. This code alows the overide to have priority over the
'State_Flag
If m_DisableColons = True Then Visi_Flag = False
If m_ForceColons = True Then Visi_Flag = True

UserControl.shpCOLON_LEFT(0).Visible = Visi_Flag
UserControl.shpCOLON_LEFT(1).Visible = Visi_Flag


End Sub

'
'During control resizing it is necesary to recalculate size and position
'of segment elements (Inactive and Active) and then redraw. This can be ugly
'so the segments are blanked during resize and are only redrawn when
'scaling is complete
'

Private Sub Display_Active_Elements(State_Flag As Boolean)

Dim Visi_Flag As Boolean
Dim Visi_Loop As Integer
Dim End_Element As Integer

'Last element of the Inactive set of the elements array
End_Element = 27

If State_Flag = True Then Visi_Flag = True

If m_BlankMSB = True And Visi_Flag = True Then End_Element = 20

'loop through element and set to value of State_Flag
For Visi_Loop = 14 To End_Element
'
'Segments are made up of three elementals
'
'the left hand 'Curve'
UserControl.shpSEG_END1(Visi_Loop).Visible = Visi_Flag
'the centre 'Bar'
UserControl.shpSEG_MID(Visi_Loop).Visible = Visi_Flag
'the right hand 'Curve'
UserControl.shpSEG_END2(Visi_Loop).Visible = Visi_Flag

Next Visi_Loop


'The colons are made of elements to and require blanking etc.
'but the control provides an 'Overide' for colons so they are never shown
'or always shown. This code alows the overide to have priority over the
'State_Flag
If m_DisableColons = True Then Visi_Flag = False
If m_ForceColons = True Then Visi_Flag = True

UserControl.shpCOLON_LEFT(2).Visible = Visi_Flag
UserControl.shpCOLON_LEFT(3).Visible = Visi_Flag

End Sub









'When the user sets the AutoTimebase it is divided by two and used to
'Control the colons and AutoTrigger at 50% duty cycle
Private Sub tmrTimeBase_Sub_Timer()

'remember what action was taken last time (Count or Colons flash)
Static toggle As Boolean

If toggle = False Then
toggle = True
'switch the colons to active
Colons_Control True
'increment the counter as per the value of CountStep
Counter_trigger
Else
toggle = False
'switch the colons to inactive
Colons_Control False
End If

End Sub

'Divide the incoming Autotimebase by two for auto colon control
Private Sub Divide_Timebase(m_AutoTimebase As Integer)

'Because the timebase is divided by 2 the minimum value is 2
If m_AutoTimebase < 2 Then
MsgBox "Minimum Value for Auto Timebase is 2 Milliseconds", vbOKOnly + vbInformation, "Invalid Property Value"
m_AutoTimebase = 2
End If

'update the timer with the new value
tmrTimeBase_Sub.Interval = Int(m_AutoTimebase / 2)
End Sub
'
'the image control sits behind the segements and needs to be updated if the control changes size
Private Sub Update_Imagebox()

'ADJUST IMAGE BOX
UserControl.imgBACKDROP.Height = UserControl.ScaleHeight
UserControl.imgBACKDROP.Width = UserControl.ScaleWidth

End Sub

'
'When to control is resized EVERY element on the control needs to have its size and position
'calculated relative to the controls new size
Private Sub UserControl_Resize()

Dim WRS As Variant
Dim HRS As Variant
Dim Element_Loop As Integer

'BLANK DISPLAY TO HIDE REDRAW OF ELEMENTS
Display_Active_Elements False
Display_Inactive_Elements False


'CALCULATE SCALE RATIO BY DIVIDING CURRENT SIZE BY DESIGN SIZE
'original width of control during design time is 10185
'so of new width is 5092.5 then 5092.5 divided by 10185 = .5
'so every size and position is multiplied by .5 - easy!

'width scaling
WRS = CDec(UserControl.ScaleWidth / 10185)
'height scaling
HRS = CDec(UserControl.ScaleHeight / 6975)

'Size of form is changing, resize image box so that any image remains
'sized if the 'Stretch' flag is true
Update_Imagebox

'SET UP LOOP TO GO THROUGH ALL INACTIVE ELEMENTS
For Element_Loop = 0 To 27

'by using the case statement i can group elements that need identical
'scaling calculations
Select Case Element_Loop

'SELECT HORIZONTAL ELEMENTS
Case 0, 14, 3, 17, 6, 20, 7, 21, 10, 24, 13, 27

'SETUP NEW WIDTH HORIZONTAL ELEMENTS
UserControl.shpSEG_END1(Element_Loop).Width = 735 * WRS
UserControl.shpSEG_MID(Element_Loop).Width = 1935 * WRS
UserControl.shpSEG_END2(Element_Loop).Width = 735 * WRS

'SETUP NEW HEIGHT HORIZONTAL ELEMENTS
UserControl.shpSEG_END1(Element_Loop).Height = 495 * HRS
UserControl.shpSEG_MID(Element_Loop).Height = 495 * HRS
UserControl.shpSEG_END2(Element_Loop).Height = 495 * HRS

'SELECT VERTICALLY ALIGNED ELEMENTS THAT ARE GOING TO MOVE
'IN THE SAME WAY AS EACH OTHER AND APPLY NEW CO-ORDS
If Element_Loop = 0 Or Element_Loop = 3 Or Element_Loop = 6 Or Element_Loop = 14 Or Element_Loop = 17 Or Element_Loop = 20 Then
UserControl.shpSEG_END1(Element_Loop).Left = 6600 * WRS
UserControl.shpSEG_MID(Element_Loop).Left = 6960 * WRS
UserControl.shpSEG_END2(Element_Loop).Left = 8520 * WRS
End If

'SELECT VERTICALLY ALIGNED ELEMENTS THAT ARE GOING TO MOVE
'IN THE SAME WAY AS EACH OTHER AND APPLY NEW CO-ORDS
If Element_Loop = 7 Or Element_Loop = 10 Or Element_Loop = 13 Or Element_Loop = 21 Or Element_Loop = 24 Or Element_Loop = 27 Then
UserControl.shpSEG_END1(Element_Loop).Left = 1800 * WRS
UserControl.shpSEG_MID(Element_Loop).Left = 2160 * WRS
UserControl.shpSEG_END2(Element_Loop).Left = 3720 * WRS
End If

'SELECT VERTICALLY ALIGNED ELEMENTS THAT ARE GOING TO MOVE
'IN THE SAME WAY AS EACH OTHER AND APPLY NEW CO-ORDS
If Element_Loop = 6 Or Element_Loop = 13 Or Element_Loop = 20 Or Element_Loop = 27 Then
UserControl.shpSEG_END1(Element_Loop).Top = 3240 * HRS
UserControl.shpSEG_MID(Element_Loop).Top = 3240 * HRS
UserControl.shpSEG_END2(Element_Loop).Top = 3240 * HRS
End If

'SELECT VERTICALLY ALIGNED ELEMENTS THAT ARE GOING TO MOVE
'IN THE SAME WAY AS EACH OTHER AND APPLY NEW CO-ORDS
If Element_Loop = 3 Or Element_Loop = 10 Or Element_Loop = 17 Or Element_Loop = 24 Then
UserControl.shpSEG_END1(Element_Loop).Top = 6480 * HRS
UserControl.shpSEG_MID(Element_Loop).Top = 6480 * HRS
UserControl.shpSEG_END2(Element_Loop).Top = 6480 * HRS
End If

'SELECT ALL VERTICAL ELEMENTS
Case 1, 15, 2, 16, 4, 18, 5, 19, 8, 22, 9, 23, 11, 25, 12, 26

'APPLY NEW WIDTH TO ALL VERTICAL ELEMENTS
UserControl.shpSEG_END1(Element_Loop).Width = 495 * WRS
UserControl.shpSEG_MID(Element_Loop).Width = 495 * WRS
UserControl.shpSEG_END2(Element_Loop).Width = 495 * WRS

'APPLY NEW HEIGHT TO ALL VERTICAL ELEMENTS
UserControl.shpSEG_END1(Element_Loop).Height = 735 * HRS
UserControl.shpSEG_MID(Element_Loop).Height = 1935 * HRS
UserControl.shpSEG_END2(Element_Loop).Height = 735 * HRS

'SELECT HORIZONTALLY ALIGNED ELEMENTS THAT ARE GOING TO
'MOVE IN THE SAME WAY AS EACH OTHER AND APPLY NEW CO-ORDS
If Element_Loop = 1 Or Element_Loop = 2 Or Element_Loop = 15 Or Element_Loop = 16 Then
UserControl.shpSEG_END1(Element_Loop).Left = 9240 * WRS
UserControl.shpSEG_MID(Element_Loop).Left = 9240 * WRS
UserControl.shpSEG_END2(Element_Loop).Left = 9240 * WRS
End If

'SELECT HORIZONTALLY ALIGNED ELEMENTS THAT ARE GOING TO
'MOVE IN THE SAME WAY AS EACH OTHER AND APPLY NEW CO-ORDS
If Element_Loop = 4 Or Element_Loop = 5 Or Element_Loop = 18 Or Element_Loop = 19 Then
UserControl.shpSEG_END1(Element_Loop).Left = 6240 * WRS
UserControl.shpSEG_MID(Element_Loop).Left = 6240 * WRS
UserControl.shpSEG_END2(Element_Loop).Left = 6240 * WRS
End If

'SELECT HORIZONTALLY ALIGNED ELEMENTS THAT ARE GOING TO
'MOVE IN THE SAME WAY AS EACH OTHER AND APPLY NEW CO-ORDS
If Element_Loop = 8 Or Element_Loop = 9 Or Element_Loop = 22 Or Element_Loop = 23 Then
UserControl.shpSEG_END1(Element_Loop).Left = 4440 * WRS
UserControl.shpSEG_MID(Element_Loop).Left = 4440 * WRS
UserControl.shpSEG_END2(Element_Loop).Left = 4440 * WRS
End If

'SELECT HORIZONTALLY ALIGNED ELEMENTS THAT ARE GOING TO
'MOVE IN THE SAME WAY AS EACH OTHER AND APPLY NEW CO-ORDS
If Element_Loop = 11 Or Element_Loop = 12 Or Element_Loop = 25 Or Element_Loop = 26 Then
UserControl.shpSEG_END1(Element_Loop).Left = 1440 * WRS
UserControl.shpSEG_MID(Element_Loop).Left = 1440 * WRS
UserControl.shpSEG_END2(Element_Loop).Left = 1440 * WRS
End If

'SELECT HORIZONTALLY ALIGNED ELEMENTS THAT ARE GOING TO
'MOVE IN THE SAME WAY AS EACH OTHER AND APPLY NEW CO-ORDS
If Element_Loop = 1 Or Element_Loop = 5 Or Element_Loop = 8 Or Element_Loop = 12 Or Element_Loop = 15 Or Element_Loop = 19 Or Element_Loop = 22 Or Element_Loop = 26 Then
UserControl.shpSEG_END1(Element_Loop).Top = 600 * HRS
UserControl.shpSEG_MID(Element_Loop).Top = 960 * HRS
UserControl.shpSEG_END2(Element_Loop).Top = 2520 * HRS
End If

'SELECT HORIZONTALLY ALIGNED ELEMENTS THAT ARE GOING TO
'MOVE IN THE SAME WAY AS EACH OTHER AND APPLY NEW CO-ORDS
If Element_Loop = 2 Or Element_Loop = 4 Or Element_Loop = 9 Or Element_Loop = 11 Or Element_Loop = 16 Or Element_Loop = 18 Or Element_Loop = 23 Or Element_Loop = 25 Then
UserControl.shpSEG_END1(Element_Loop).Top = 3720 * HRS
UserControl.shpSEG_MID(Element_Loop).Top = 4080 * HRS
UserControl.shpSEG_END2(Element_Loop).Top = 5640 * HRS
End If

End Select

Next Element_Loop

'APPLY SCALED POSITION CHANGES TO LEFT HAND COLONS
UserControl.shpCOLON_LEFT(0).Top = 1496 * HRS
UserControl.shpCOLON_LEFT(0).Left = 220 * WRS

UserControl.shpCOLON_LEFT(3).Top = 4984 * HRS
UserControl.shpCOLON_LEFT(3).Left = 220 * WRS

UserControl.shpCOLON_LEFT(2).Top = 1496 * HRS
UserControl.shpCOLON_LEFT(2).Left = 220 * WRS

UserControl.shpCOLON_LEFT(1).Top = 4984 * HRS
UserControl.shpCOLON_LEFT(1).Left = 220 * WRS


'If the scaling of the colons results in a size of less than 23 then
'due to restriction is resolution they are invisible!
'so restrict the smallest size to 23
If 495 * HRS < 23 Then
UserControl.shpCOLON_LEFT(0).Width = 23
UserControl.shpCOLON_LEFT(0).Height = 23
UserControl.shpCOLON_LEFT(2).Width = 23
UserControl.shpCOLON_LEFT(2).Height = 23
UserControl.shpCOLON_LEFT(1).Width = 23
UserControl.shpCOLON_LEFT(1).Height = 23
UserControl.shpCOLON_LEFT(3).Width = 23
UserControl.shpCOLON_LEFT(3).Height = 23
Else
UserControl.shpCOLON_LEFT(0).Width = 495 * HRS
UserControl.shpCOLON_LEFT(0).Height = 495 * HRS
UserControl.shpCOLON_LEFT(2).Width = 495 * HRS
UserControl.shpCOLON_LEFT(2).Height = 495 * HRS
UserControl.shpCOLON_LEFT(1).Width = 495 * HRS
UserControl.shpCOLON_LEFT(1).Height = 495 * HRS
UserControl.shpCOLON_LEFT(3).Width = 495 * HRS
UserControl.shpCOLON_LEFT(3).Height = 495 * HRS
End If


'SHOW DISPLAY NOW THE REDRAW OF ELEMENTS IS COMPLETE
Display_Active_Elements True
Display_Inactive_Elements True

'RE-APPLY NUMBER TO NEWLY REFRESHED DISPLAY
Display_Number m_value

'RESTORE COLUNS TO PREVIOUS STATE
Colons_Control m_ColonsState

End Sub
'
'if the display is used as a basic counter then it has to be triggered either
'manually or automatically
Private Sub Counter_trigger(Optional Direct_Add As Long)

'm_value holds the current value of the display
'm_countstep holds the increment value applied during each trigger
If Direct_Add = 0 Then m_value = m_value + m_CountStep

'allows direct adding to counter, for carrys etc
If Direct_Add <> 0 Then m_value = m_value + Direct_Add

'user can set custom m_Maxcount for display EG 60 for simple clock etc.
'when current count exceeds m_Maxcount raise an overflow event for carry
'out functions, alarm clock etc.
If m_value > m_Maxcount Then Overflow

'user can set custom m_Mincount for display EG 60 for simple Egg timer etc.
'when current count goes below m_Mincount raise an Underflow event for carry
'out functions, time elapsed etc.
If m_value < m_Mincount Then Underflow

'Display new value of m_value
Display_Number m_value

End Sub
'
''
''m_Maxcount reached by trigger
''
Private Sub Overflow()

Dim Carry_Out As Long

Carry_Out = (m_value - m_Maxcount) - 1

'reset counter to user lower setting
m_value = m_Mincount

'call external event that user can use to control other functions
RaiseEvent Overflow(Carry_Out)

End Sub

'
'm_Mincount reached by trigger
'
Private Sub Underflow()

Dim Carry_Out As Long

Carry_Out = (m_value - m_Mincount) + 1

'reset counter to user upper setting
m_value = m_Maxcount

'call external event that user can use to control other functions
RaiseEvent Underflow(Carry_Out)


End Sub

'
'initial routine for handling the display of a two digit number into the duak 7 segment
'array
'
Private Sub Display_Number(V_dec As Variant)

Dim Units_Val As Variant
Dim Tens_Val As Variant

'keep the value of m_value up to date (this enables the routine to be called from anywhere
'in the control whilst maintaining the sync of queries to the .Value variable)
m_value = V_dec
'pass onto external reference
PropertyChanged "Value"

'BREAK NUMBER TO UNITS AND TENS FOR DISPLAY BY THE (M)OST (S)IGNIFICANT (B)IT
'AND THE (L)EAST (S)IGNIFICANT (B)IT
Tens_Val = Int(V_dec / 10)
Units_Val = CDec(V_dec - (Tens_Val * 10))

'PASS ONTO INDIVIDUAL SEGMENT ENCODING ROUTINES
Display_Encode Units_Val, "Units"

'THE CONTROL PROVIDES FOR A MSB BLANKING OPTION, INCASE THE USER WANT TO BE PRECISE AND ONLY
'WANTS SAY [8]8 88 88 88
'          / \BLANKED, SO DISPLAY ONLY READS for EXAMPLE 1,000,000 INSTEAD OF 01,000,000

'SO MSB IS NOT DRAWN IS m_BlankMSB = true
If m_BlankMSB = False Then Display_Encode Tens_Val, "Tens"

End Sub

'
'this routine encodes a number from 0 to 9 into an array of boolean
'flags that individualy controls the elements .Visible property
'
Private Sub Display_Encode(Number As Variant, Selection As String)

Dim Segment_array() As Boolean
Dim Encode_Loop As Integer
Dim Selection_Offset As Integer

'as all 'segment' elements are part of one control array it is east to create
'one routine to control them all, by altering the start and end to loops groups
'of elements can be altered together

'the 'Units' LSB occupies elements 14 to 20
If Selection = "Units" Then Selection_Offset = 14

'the 'Tens' MSB occupies elements 21 to 27
If Selection = "Tens" Then Selection_Offset = 21


'size the array to accept 7 elements
ReDim Segment_array(6)

'0 is top
'1 is top right
'2 is bottom right
'3 is bottom
'4 is bottom left
'5 is top left
'6 is centre

Select Case Number

'individual segment assignmets
'0 is top
'1 is top right
'2 is bottom right
'3 is bottom
'4 is bottom left
'5 is top left
'6 is centre

Case Is = 0

Segment_array(0) = 1
Segment_array(1) = 1
Segment_array(2) = 1
Segment_array(3) = 1
Segment_array(4) = 1
Segment_array(5) = 1
Segment_array(6) = 0

Case Is = 1

Segment_array(0) = 0
Segment_array(1) = 1
Segment_array(2) = 1
Segment_array(3) = 0
Segment_array(4) = 0
Segment_array(5) = 0
Segment_array(6) = 0

Case Is = 2

Segment_array(0) = 1
Segment_array(1) = 1
Segment_array(2) = 0
Segment_array(3) = 1
Segment_array(4) = 1
Segment_array(5) = 0
Segment_array(6) = 1

Case Is = 3
Segment_array(0) = 1
Segment_array(1) = 1
Segment_array(2) = 1
Segment_array(3) = 1
Segment_array(4) = 0
Segment_array(5) = 0
Segment_array(6) = 1

Case Is = 4
Segment_array(0) = 0
Segment_array(1) = 1
Segment_array(2) = 1
Segment_array(3) = 0
Segment_array(4) = 0
Segment_array(5) = 1
Segment_array(6) = 1

Case Is = 5
Segment_array(0) = 1
Segment_array(1) = 0
Segment_array(2) = 1
Segment_array(3) = 1
Segment_array(4) = 0
Segment_array(5) = 1
Segment_array(6) = 1

Case Is = 6
Segment_array(0) = 1
Segment_array(1) = 0
Segment_array(2) = 1
Segment_array(3) = 1
Segment_array(4) = 1
Segment_array(5) = 1
Segment_array(6) = 1

Case Is = 7
Segment_array(0) = 1
Segment_array(1) = 1
Segment_array(2) = 1
Segment_array(3) = 0
Segment_array(4) = 0
Segment_array(5) = 0
Segment_array(6) = 0

Case Is = 8
Segment_array(0) = 1
Segment_array(1) = 1
Segment_array(2) = 1
Segment_array(3) = 1
Segment_array(4) = 1
Segment_array(5) = 1
Segment_array(6) = 1

Case Is = 9
Segment_array(0) = 1
Segment_array(1) = 1
Segment_array(2) = 1
Segment_array(3) = 0
Segment_array(4) = 0
Segment_array(5) = 1
Segment_array(6) = 1
End Select



'apply encoded array onto selected group of element array
For Encode_Loop = 0 To 6

'as each segment is made from three elementals, there are three elements to set
UserControl.shpSEG_END1(Encode_Loop + Selection_Offset).Visible = Segment_array(Encode_Loop)
UserControl.shpSEG_MID(Encode_Loop + Selection_Offset).Visible = Segment_array(Encode_Loop)
UserControl.shpSEG_END2(Encode_Loop + Selection_Offset).Visible = Segment_array(Encode_Loop)

Next Encode_Loop


End Sub
'
'colons can be set to either active or inactive (the same as segments)
'
Private Sub Colons_Control(Optional State As Boolean)

'm_DisbleColons hold the overide flag for the colons, if its true then
'colons cannot be displayed no matter the setting of state
If m_DisableColons = True Then
State = False
UserControl.shpCOLON_LEFT(0).Visible = 0
UserControl.shpCOLON_LEFT(1).Visible = 0
Else
UserControl.shpCOLON_LEFT(0).Visible = 1
UserControl.shpCOLON_LEFT(1).Visible = 1
End If

'm_ColonsForce holds the flag for the overide flag for the colons, if
'it's true then colons will be shown no matter the setting of any other colons flag
If m_ForceColons = True Then
State = True
UserControl.shpCOLON_LEFT(0).Visible = 1
UserControl.shpCOLON_LEFT(1).Visible = 1
End If

'assuming no overide flags have set the state flag to anything else, the users original
'flag will be applied to the colons
If State = False Then
m_ColonsState = False
UserControl.shpCOLON_LEFT(2).Visible = 0
UserControl.shpCOLON_LEFT(3).Visible = 0
Else
m_ColonsState = True
UserControl.shpCOLON_LEFT(2).Visible = 1
UserControl.shpCOLON_LEFT(3).Visible = 1
End If
'
'raise an external event, user can use this to control other things (like other display
'controls on the form)
RaiseEvent ColonsChangeState(State)

End Sub
'
'colour setting for the active segments and colons can be set here
'
Private Sub update_Active_colour(m_ActiveColour As OLE_COLOR)

Dim Colour_Loop As Integer

'all active segments fall between 14 and 27 of the element aray
For Colour_Loop = 14 To 27

'left hand curve fill and border colour
shpSEG_END1(Colour_Loop).FillColor = m_ActiveColour
shpSEG_END1(Colour_Loop).BorderColor = m_ActiveColour

'centre bar fill and border colour
shpSEG_MID(Colour_Loop).FillColor = m_ActiveColour
shpSEG_MID(Colour_Loop).BorderColor = m_ActiveColour

'cright hand curve fill and border colour
shpSEG_END2(Colour_Loop).FillColor = m_ActiveColour
shpSEG_END2(Colour_Loop).BorderColor = m_ActiveColour

Next Colour_Loop

'colons fill and border colour
UserControl.shpCOLON_LEFT(2).FillColor = m_ActiveColour
UserControl.shpCOLON_LEFT(2).BorderColor = m_ActiveColour

UserControl.shpCOLON_LEFT(3).FillColor = m_ActiveColour
UserControl.shpCOLON_LEFT(3).BorderColor = m_ActiveColour

End Sub
'
'colour setting for the Inactive segments and colons can be set here
'
Private Sub update_Inactive_colour(New_colour As OLE_COLOR)

Dim Colour_Loop As Integer

'all Inactive segments fall between 0 and 13 of the element aray
For Colour_Loop = 0 To 13

'left hand curve fill and border colour
shpSEG_END1(Colour_Loop).FillColor = New_colour
shpSEG_END1(Colour_Loop).BorderColor = New_colour

'centre bar fill and border colour
shpSEG_MID(Colour_Loop).FillColor = New_colour
shpSEG_MID(Colour_Loop).BorderColor = New_colour

'right hand curve fill and border colour
shpSEG_END2(Colour_Loop).FillColor = New_colour
shpSEG_END2(Colour_Loop).BorderColor = New_colour

Next Colour_Loop

'colons fill and border colour
UserControl.shpCOLON_LEFT(0).FillColor = New_colour
UserControl.shpCOLON_LEFT(0).BorderColor = New_colour

UserControl.shpCOLON_LEFT(1).FillColor = New_colour
UserControl.shpCOLON_LEFT(1).BorderColor = New_colour

End Sub

'
'the backstyle transparency property of the control can be set by the user
'
Private Sub Update_Backstyle(Back_Flag As Boolean)

If Back_Flag = True Then

'make transparent
UserControl.BackStyle = 0

Else

'make opaque
UserControl.BackStyle = 1

End If

End Sub
'
'the background colour of the control can be set by the user
'
Private Sub update_backcolour(back_colour As OLE_COLOR)

UserControl.BackColor = back_colour

End Sub

'
'when the user interfaces with the control this routine syncronises internal variables
'
Private Sub update_colons_disable(m_DisableColons As Boolean)

If m_ForceColons = True Then
m_DisableColons = False
Exit Sub
End If

Colons_Control

End Sub
'
'when the user interfaces with the control this routine syncronises internal variables
'
Private Sub update_colons_force(m_ForceColons As Boolean)

If m_ForceColons = True Then
m_DisableColons = False
Colons_Control True
PropertyChanged "DisableColons"
End If

End Sub
'
'when the user interfaces with the control this routine syncronises internal variables
'm_Maxcoutn is the users highest number to count to.
Private Sub update_Maxcount(m_Maxcount As Integer)

'maximum is 99, anything else is obviously not allowed
If m_Maxcount > 99 Then
MsgBox "Maximum Value is 99", vbOKOnly + vbInformation, "Invalid Property Value"
m_Maxcount = 99
End If

'the maximum cannot be less than 1
If m_Maxcount < 1 Then
MsgBox "Minimum Value is 1", vbOKOnly + vbInformation, "Invalid Property Value"
m_Maxcount = m_Mincount + 1
End If

'if the maximum is set less that the users minimum then thats no good either!
If m_Maxcount < m_Mincount Then
MsgBox "Maximum Value cannot be less than Minimum Value", vbOKOnly + vbInformation, "Invalid Property Value"
m_Maxcount = m_Mincount + 1
End If


End Sub

'
'when the user interfaces with the control this routine syncronises internal variables
'
Private Sub update_Mincount(m_Mincount As Integer)

'the mimium user lever is 0
If m_Mincount < 0 Then
MsgBox "Minimum Value is 0", vbOKOnly + vbInformation, "Invalid Property Value"
m_Mincount = 0
End If

'the highest minimum user lever is 98
If m_Mincount > 98 Then
MsgBox "Maximum Value is 98", vbOKOnly + vbInformation, "Invalid Property Value"
m_Mincount = m_Maxcount - 1
End If

'if the user sets the min to greater than the max...
If m_Mincount > m_Maxcount Then
MsgBox "Minimum Value cannot be more than Maximum Value", vbOKOnly + vbInformation, "Invalid Property Value"
m_Mincount = m_Maxcount - 1
End If

End Sub

'
'when the user interfaces with the control this routine syncronises internal variables
'
Private Sub update_step_size(m_CountStep As Long)

If m_CountStep < -99 Then
MsgBox "Minimum Value is -99", vbOKOnly + vbInformation, "Invalid Property Value"
m_CountStep = -99
End If

If m_CountStep > 99 Then
MsgBox "Maximum Value is 99", vbOKOnly + vbInformation, "Invalid Property Value"
m_CountStep = 99
End If


End Sub

'
'when the user interfaces with the control this routine syncronises internal variables
'Blank the MSB
Private Sub update_msbblank(m_BlankMSB As Boolean)

'if the user chooses to blank the MSB and its not already blanked then
'procced to call other routines, if the MSB is already blanked then ignore
'to save overhead
If m_BlankMSB = True And UserControl.shpSEG_MID(7).Visible = True Then

Display_Inactive_Elements False
Display_Inactive_Elements True

Display_Active_Elements False

End If

If m_BlankMSB = False And UserControl.shpSEG_MID(7).Visible = False Then

Display_Inactive_Elements True

End If

'RE-APPLY NUMBER TO NEWLY REFRESHED DISPLAY
Display_Number m_value
'RESTORE COLUNS TO PREVIOUS STATE
Colons_Control m_ColonsState

End Sub
'
'when the user interfaces with the control this routine syncronises internal variables
'control the currently displyed value directly as opposed to trigger incrementation
Private Sub update_display_value(m_value As Long)

'max display value is 99
If m_value > 99 Then
MsgBox "Maximum value that can be displayed is 99", vbOKOnly + vbInformation, "Invalid Property Value"
m_value = 99
End If

'minimum display value is 0
If m_value < 0 Then
MsgBox "Minimum value that can be displayed is 0", vbOKOnly + vbInformation, "Invalid Property Value"
m_value = 0
End If


'RE-APPLY NUMBER TO NEWLY REFRESHED DISPLAY
Display_Number m_value

'RESTORE COLUNS TO PREVIOUS STATE
Colons_Control m_ColonsState

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1000
Public Property Get AutoTimebase() As Integer
Attribute AutoTimebase.VB_Description = "Sets / Returns value (in milliseconds), of the timer when in Auto Mode."
    AutoTimebase = m_AutoTimebase
    'send new value to be divided and sent to timer
    Divide_Timebase m_AutoTimebase
End Property

Public Property Let AutoTimebase(ByVal New_AutoTimebase As Integer)
    m_AutoTimebase = New_AutoTimebase
    
    'send new value to be divided and sent to timer
    Divide_Timebase m_AutoTimebase
    
    PropertyChanged "AutoTimebase"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_AutoTimebase = m_def_AutoTimebase
    m_ActiveColour = m_def_ActiveColour
    m_Backcolour = m_def_Backcolour
    m_InactiveColour = m_def_InactiveColour
    m_BackTransparent = m_def_BackTransparent
    m_DisableColons = m_def_DisableColons
    m_ForceColons = m_def_ForceColons
    m_Maxcount = m_def_MaxCount
    m_Mincount = m_def_MinCount
    m_CountStep = m_def_CountStep
    m_BlankMSB = m_def_BlankMSB
    m_value = m_def_Value
    m_ColonsState = m_def_ColonsState
 
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_AutoTimebase = PropBag.ReadProperty("AutoTimebase", m_def_AutoTimebase)
    m_ActiveColour = PropBag.ReadProperty("ActiveColour", m_def_ActiveColour)
    m_Backcolour = PropBag.ReadProperty("Backcolour", m_def_Backcolour)
    m_InactiveColour = PropBag.ReadProperty("InactiveColour", m_def_InactiveColour)
    m_BackTransparent = PropBag.ReadProperty("BackTransparent", m_def_BackTransparent)
    m_DisableColons = PropBag.ReadProperty("DisableColons", m_def_DisableColons)
    m_ForceColons = PropBag.ReadProperty("ForceColons", m_def_ForceColons)
    m_Maxcount = PropBag.ReadProperty("MaxCount", m_def_MaxCount)
    m_Mincount = PropBag.ReadProperty("MinCount", m_def_MinCount)
    m_CountStep = PropBag.ReadProperty("CountStep", m_def_CountStep)
    m_BlankMSB = PropBag.ReadProperty("BlankMSB", m_def_BlankMSB)
    m_value = PropBag.ReadProperty("Value", m_def_Value)
    m_ColonsState = PropBag.ReadProperty("ColonsState", m_def_ColonsState)
    global_update
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    imgBACKDROP.Stretch = PropBag.ReadProperty("Stretch", False)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub
'
'when the user interfaces with the control this routine syncronises internal variables
'
Private Sub global_update()
    update_Inactive_colour m_InactiveColour
    update_Active_colour m_ActiveColour
    update_backcolour m_Backcolour
    Update_Backstyle m_BackTransparent
    update_colons_disable m_DisableColons
    update_colons_force m_ForceColons
    update_Maxcount m_Maxcount
    update_Mincount m_Mincount
    update_step_size m_CountStep
    update_msbblank m_BlankMSB
    update_display_value m_value
    Divide_Timebase m_AutoTimebase
    Update_Imagebox
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AutoTimebase", m_AutoTimebase, m_def_AutoTimebase)
    Call PropBag.WriteProperty("ActiveColour", m_ActiveColour, m_def_ActiveColour)
    Call PropBag.WriteProperty("Backcolour", m_Backcolour, m_def_Backcolour)
    Call PropBag.WriteProperty("InactiveColour", m_InactiveColour, m_def_InactiveColour)
    Call PropBag.WriteProperty("BackTransparent", m_BackTransparent, m_def_BackTransparent)
    Call PropBag.WriteProperty("DisableColons", m_DisableColons, m_def_DisableColons)
    Call PropBag.WriteProperty("ForceColons", m_ForceColons, m_def_ForceColons)
    Call PropBag.WriteProperty("MaxCount", m_Maxcount, m_def_MaxCount)
    Call PropBag.WriteProperty("MinCount", m_Mincount, m_def_MinCount)
    Call PropBag.WriteProperty("CountStep", m_CountStep, m_def_CountStep)
    Call PropBag.WriteProperty("BlankMSB", m_BlankMSB, m_def_BlankMSB)
    Call PropBag.WriteProperty("Value", m_value, m_def_Value)
    Call PropBag.WriteProperty("ColonsState", m_ColonsState, m_def_ColonsState)
    Call PropBag.WriteProperty("Stretch", imgBACKDROP.Stretch, False)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&hff&
Public Property Get ActiveColour() As OLE_COLOR
Attribute ActiveColour.VB_Description = "Sets / Returns the colour of the segments when active"
    ActiveColour = m_ActiveColour
    update_Active_colour m_ActiveColour
End Property

Public Property Let ActiveColour(ByVal New_ActiveColour As OLE_COLOR)
    m_ActiveColour = New_ActiveColour
    update_Active_colour m_ActiveColour
    PropertyChanged "ActiveColour"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H0&
Public Property Get Backcolour() As OLE_COLOR
Attribute Backcolour.VB_Description = "Sets / Returns the background colour of the control"
    Backcolour = m_Backcolour
    update_backcolour m_Backcolour
    
End Property

Public Property Let Backcolour(ByVal New_Backcolour As OLE_COLOR)
    m_Backcolour = New_Backcolour
     update_backcolour m_Backcolour
    PropertyChanged "Backcolour"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&h80&
Public Property Get InactiveColour() As OLE_COLOR
Attribute InactiveColour.VB_Description = "Sets / Returns the colour of the segments when inactive"
    InactiveColour = m_InactiveColour
    update_Inactive_colour m_InactiveColour
End Property

Public Property Let InactiveColour(ByVal New_InactiveColour As OLE_COLOR)
    m_InactiveColour = New_InactiveColour
    update_Inactive_colour m_InactiveColour
    PropertyChanged "InactiveColour"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get BackTransparent() As Boolean
Attribute BackTransparent.VB_Description = "Sets / Returns the background of the control to either Transparent (True), or Opaque (False)"
    BackTransparent = m_BackTransparent
    Update_Backstyle m_BackTransparent
End Property

Public Property Let BackTransparent(ByVal New_BackTransparent As Boolean)
    m_BackTransparent = New_BackTransparent
    Update_Backstyle m_BackTransparent
    PropertyChanged "BackTransparent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get DisableColons() As Boolean
Attribute DisableColons.VB_Description = "Disable and hide colons for this control. Has priority over .Colons_On  &  .ColonsOff"
    DisableColons = m_DisableColons
    update_colons_disable m_DisableColons
End Property

Public Property Let DisableColons(ByVal New_DisableColons As Boolean)
    m_DisableColons = New_DisableColons
    update_colons_disable m_DisableColons
    PropertyChanged "DisableColons"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ForceColons() As Boolean
Attribute ForceColons.VB_Description = "Forces the colons to be in an 'On' state regardless of other influences. Has priority over all other colon triggers"
    ForceColons = m_ForceColons
    update_colons_force m_ForceColons
End Property

Public Property Let ForceColons(ByVal New_ForceColons As Boolean)
    m_ForceColons = New_ForceColons
    update_colons_force m_ForceColons
    PropertyChanged "ForceColons"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,99
Public Property Get MaxCount() As Integer
Attribute MaxCount.VB_Description = "Sets / Returns the Maximum Value that the control will count to before the next trigger will cause an 'Overflow' event."
    MaxCount = m_Maxcount
    update_Maxcount m_Maxcount
End Property

Public Property Let MaxCount(ByVal New_MaxCount As Integer)
    m_Maxcount = New_MaxCount
    update_Maxcount m_Maxcount
    PropertyChanged "MaxCount"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MinCount() As Integer
Attribute MinCount.VB_Description = "Sets / Returns the Minimum Value that the control will count to before the next trigger will cause an 'Underflow' event."
    MinCount = m_Mincount
    update_Mincount m_Mincount
End Property

Public Property Let MinCount(ByVal New_MinCount As Integer)
    m_Mincount = New_MinCount
    update_Mincount m_Mincount
    PropertyChanged "MinCount"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get CountStep() As Long
Attribute CountStep.VB_Description = "Sets / Returns the value that will be added to the counter during either a Manual Trigger event or a Auto Trigger event. Can be negative for backwards counting!"
    CountStep = m_CountStep
    update_step_size m_CountStep
End Property

Public Property Let CountStep(ByVal New_CountStep As Long)
    m_CountStep = New_CountStep
    update_step_size m_CountStep
    PropertyChanged "CountStep"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get BlankMSB() As Boolean
Attribute BlankMSB.VB_Description = "Sets / Returns the visibility of the (M)ost (S)ignificant (B)it."
    BlankMSB = m_BlankMSB
    update_msbblank m_BlankMSB
    
End Property

Public Property Let BlankMSB(ByVal New_BlankMSB As Boolean)
    m_BlankMSB = New_BlankMSB
    update_msbblank m_BlankMSB
    PropertyChanged "BlankMSB"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get value() As Long
Attribute value.VB_Description = "Sets / Returns the value that is displayed or is to be displayed."
    value = m_value
    update_display_value m_value
End Property

Public Property Let value(ByVal New_Value As Long)
    m_value = New_Value
    update_display_value m_value
    PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Automode() As Variant
UserControl.tmrTimeBase_Sub.Enabled = True
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function Reset() As Variant


If m_CountStep > 0 Then
m_value = m_Mincount
End If

If m_CountStep < 0 Then
m_value = m_Maxcount
End If

If m_CountStep = 0 Then
m_value = m_Mincount
End If


Display_Number m_value

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function ManualTrigger() As Variant
Counter_trigger
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function ManualMode() As Variant
UserControl.tmrTimeBase_Sub.Enabled = False
End Function
'

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,2,0
Public Property Get ColonsState() As Boolean
Attribute ColonsState.VB_Description = "Returns the current state of the colons as set by .ColonsOn & .ColonsOff"
Attribute ColonsState.VB_MemberFlags = "400"
    ColonsState = m_ColonsState
End Property

Public Property Let ColonsState(ByVal New_ColonsState As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_ColonsState = New_ColonsState
    PropertyChanged "ColonsState"
End Property

'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgBACKDROP,imgBACKDROP,-1,Stretch
Public Property Get Stretch() As Boolean
Attribute Stretch.VB_Description = "Returns/sets a value that determines whether a graphic resizes to fit the size of an Image control."
    Stretch = imgBACKDROP.Stretch
    Update_Imagebox
End Property

Public Property Let Stretch(ByVal New_Stretch As Boolean)
    imgBACKDROP.Stretch() = New_Stretch
    Update_Imagebox
    PropertyChanged "Stretch"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=imgBACKDROP,imgBACKDROP,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = imgBACKDROP.Picture
    Update_Imagebox
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set imgBACKDROP.Picture = New_Picture
    Update_Imagebox
    PropertyChanged "Picture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function ColonsDisplay(dFlag) As Boolean
Attribute ColonsDisplay.VB_Description = "Sets the Colons to Active:True or Inactive:False"

If dFlag = True Then
Colons_Control True
Else
Colons_Control False
End If

End Function
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=7
Public Function CountDirectAdd(AddValue) As Long
Dim PassVal As Long
PassVal = AddValue
Counter_trigger PassVal
End Function


