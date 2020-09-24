VERSION 5.00
Begin VB.Form fFFT 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Fast Fourier Transformation"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12030
   Icon            =   "fFFT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   12030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fr 
      Height          =   2025
      Index           =   4
      Left            =   195
      TabIndex        =   4
      Top             =   4500
      Width           =   11685
      Begin VB.Frame rf 
         Caption         =   "Reverse Transform"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1530
         Left            =   8730
         TabIndex        =   24
         Top             =   270
         Width           =   2565
         Begin VB.CheckBox ckRemoveNoise 
            Caption         =   "...with noise and    harmonics removed"
            Height          =   360
            Left            =   135
            TabIndex        =   26
            Top             =   1050
            Width           =   2160
         End
         Begin VB.CheckBox ckReverse 
            Caption         =   "This option will perform an extra pass to reverse-trans- form the current spectrum..."
            Height          =   780
            Left            =   135
            TabIndex        =   25
            Top             =   210
            Width           =   2235
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Samples"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1530
         Index           =   0
         Left            =   525
         TabIndex        =   18
         Top             =   270
         Width           =   1275
         Begin VB.OptionButton optSamples 
            Caption         =   "  256"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   23
            Top             =   285
            Width           =   720
         End
         Begin VB.OptionButton optSamples 
            Caption         =   "  512"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   22
            Top             =   510
            Width           =   720
         End
         Begin VB.OptionButton optSamples 
            Caption         =   "1024"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   21
            Top             =   735
            Width           =   720
         End
         Begin VB.OptionButton optSamples 
            Caption         =   "2048"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   20
            Top             =   960
            Value           =   -1  'True
            Width           =   720
         End
         Begin VB.OptionButton optSamples 
            Caption         =   "4096"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   19
            Top             =   1200
            Width           =   720
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Waveform"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1530
         Index           =   1
         Left            =   1950
         TabIndex        =   12
         Top             =   270
         Width           =   2265
         Begin VB.OptionButton optWaveform 
            Caption         =   "0 Hz"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   17
            Top             =   1200
            Width           =   675
         End
         Begin VB.OptionButton optWaveform 
            Caption         =   "Trapezoid"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   16
            Top             =   960
            Width           =   1050
         End
         Begin VB.OptionButton optWaveform 
            Caption         =   "Squarewave"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   15
            Top             =   735
            Width           =   1245
         End
         Begin VB.OptionButton optWaveform 
            Caption         =   "Sine with harmonics"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   14
            Top             =   510
            Width           =   1755
         End
         Begin VB.OptionButton optWaveform 
            Caption         =   "Pure Sine"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   13
            Top             =   285
            Value           =   -1  'True
            Width           =   1050
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Characteristics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1530
         Index           =   2
         Left            =   4395
         TabIndex        =   5
         Top             =   270
         Width           =   4170
         Begin VB.HScrollBar scrNoise 
            Height          =   255
            Left            =   225
            Max             =   5
            TabIndex        =   7
            Top             =   1095
            Width           =   3240
         End
         Begin VB.HScrollBar scrFreq 
            Height          =   255
            Left            =   225
            Max             =   2048
            Min             =   1
            TabIndex        =   6
            Top             =   480
            Value           =   1
            Width           =   3240
         End
         Begin VB.Label lb 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Add some noise"
            Height          =   180
            Index           =   1
            Left            =   1365
            TabIndex        =   11
            Top             =   870
            Width           =   1125
         End
         Begin VB.Label lb 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Normalized Frequency"
            Height          =   180
            Index           =   0
            Left            =   1095
            TabIndex        =   10
            Top             =   255
            Width           =   1575
         End
         Begin VB.Label lbNoise 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   3540
            TabIndex        =   9
            Top             =   1110
            Width           =   45
         End
         Begin VB.Label lbFreq 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   3540
            TabIndex        =   8
            Top             =   495
            Width           =   45
         End
      End
   End
   Begin VB.Frame fr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4395
      Index           =   3
      Left            =   195
      TabIndex        =   0
      Top             =   60
      Width           =   11670
      Begin VB.PictureBox picDisplay 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         DrawMode        =   15  'Stift und inverse Anzeige mischen
         Height          =   3645
         Left            =   495
         ScaleHeight     =   5.926
         ScaleMode       =   0  'Benutzerdefiniert
         ScaleWidth      =   10785
         TabIndex        =   1
         Top             =   405
         Width           =   10845
      End
      Begin VB.Label lbTiming 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   5985
         TabIndex        =   27
         Top             =   165
         Width           =   75
      End
      Begin VB.Label lbMaxAt 
         Alignment       =   2  'Zentriert
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   3345
         TabIndex        =   3
         Top             =   4020
         Width           =   60
      End
      Begin VB.Label lbValue 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   465
         TabIndex        =   2
         Top             =   270
         Width           =   45
      End
   End
   Begin VB.Timer tmrTick 
      Interval        =   200
      Left            =   11595
      Top             =   4860
   End
End
Attribute VB_Name = "fFFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cFourier    As clsFourier

Private NumSamples  As Long 'number of samples we're gonna use
Private SelWaveform As Long 'selector for waveform
Private InclReverse As Boolean 'include reverse transform
Private RemoveNoise As Boolean 'apply filter
Private Omega       As Double ' 2 * pi * f
Private SvdReal()   As Double 'saved reals for reverse transform
Private SvdImag()   As Double 'saved imaginaries for reverse transform

Private Sub ckRemoveNoise_Click()

    RemoveNoise = (ckRemoveNoise = vbChecked)

End Sub

Private Sub ckReverse_Click()

    InclReverse = (ckReverse = vbChecked)
    Form_Load

End Sub

Private Sub Form_Initialize()

    If InIDE Then
        MsgBox "Please compile me; I'm twelve times faster when compiled.", , "Fourier Transformation"
    End If
    Set cFourier = New clsFourier
    NumSamples = 2048
    scrFreq = NumSamples / 15

End Sub

Private Sub Form_Load()

    cFourier.NumberOfSamples = NumSamples
    If InclReverse Then
        ReDim SvdReal(1 To NumSamples), SvdImag(1 To NumSamples)
      Else 'INCLREVERSE = FALSE/0
        Erase SvdReal, SvdImag
    End If
    picDisplay.ScaleLeft = 1
    picDisplay.ScaleWidth = NumSamples / 2
    scrFreq.Max = NumSamples / 2 - 1 'Nyquist Theorem: sampling freq must be greater than twice sampled freq
    If ckReverse = vbChecked Then
        scrFreq.Value = 7 'make it real slow so people can see what's going on
        fr(3).Caption = "Spectrum and Waveform"
      Else 'NOT CKREVERSE...
        scrFreq = NumSamples / 15
        fr(3).Caption = "Spectrum"
    End If
    scrNoise_Change
    tmrTick_Timer

End Sub

Private Function InIDE(Optional c As Boolean = False) As Boolean

  Static b As Boolean

    b = c
    If b = False Then
        Debug.Assert InIDE(True)
    End If
    InIDE = b

End Function

Private Sub optSamples_Click(Index As Integer)

    NumSamples = 2 ^ (Index + 8)
    Form_Load

End Sub

Private Sub optWaveform_Click(Index As Integer)

    SelWaveform = Index
    Form_Load
    scrFreq.Enabled = (Index <> 4)
    lbFreq.Visible = (Index <> 4)
    ckRemoveNoise = vbUnchecked

End Sub

Private Sub scrFreq_Change()

    lbFreq = scrFreq
    Omega = 8 * Atn(1) * scrFreq

End Sub

Private Sub scrFreq_Scroll()

    scrFreq_Change

End Sub

Private Sub scrNoise_Change()

    lbNoise = "x " & scrNoise
    ckRemoveNoise = vbUnchecked

End Sub

Private Sub scrNoise_Scroll()

    scrNoise_Change

End Sub

Private Sub tmrTick_Timer()

  Dim i     As Long
  Dim j     As Long
  Dim tmp   As Double
  Dim Max   As Double

    With cFourier

        .TransformReverse = False

        For i = 1 To NumSamples 'a selection of homemade samples
            Select Case SelWaveform
              Case 0 'pure sine
                '=========================================================================================
                .RealIn(i) = Sin(Omega * i / NumSamples) + (Rnd - Rnd) * scrNoise
                '=========================================================================================
              Case 1 'sine with some harmonics
                'note: some aliased peaks at higher frequencies where harmonic frequencies exceed
                'the 'permitted' bandwidth
                .RealIn(i) = Sin(Omega * i / NumSamples) ^ 7 + (Rnd - Rnd) * scrNoise
                '=========================================================================================
              Case 2 'square wave
                'note: a square wave contains all odd harmonics and thus frequencies above the 'permitted'
                'bandwidth - you will therefore see aliased peaks in the fourier transformation
                cFourier.RealIn(i) = Sgn(Sin(Omega * i / NumSamples)) + (Rnd - Rnd) * scrNoise
                '=========================================================================================
              Case 3 'trapezoid
                'note: this reduces the aliased peaks because the waveform contains less harmonics
                tmp = Sin(Omega * i / NumSamples) * 3
                Select Case tmp
                  Case Is > 1
                    tmp = 1
                  Case Is < -1
                    tmp = -1
                End Select
                .RealIn(i) = tmp + (Rnd - Rnd) * scrNoise
                '=========================================================================================
              Case Else '0 Hertz
                .RealIn(i) = 1 + (Rnd - Rnd) * scrNoise
                '=========================================================================================
            End Select

            .ImagIn(i) = 0 'we have no imaginary part so we just suppy a zero

        Next i

        'find biggest out-value so that we can scale the picbox
        Max = -1
        For i = 1 To NumSamples
            If i <= NumSamples / 2 Then 'above that point are aliased echo peaks
                tmp = .ComplexOut(i)
                If tmp > Max Then
                    Max = tmp
                    j = i
                End If
            End If
            If InclReverse Then
                SvdReal(i) = .RealOut(i)
                SvdImag(i) = .ImagOut(i)
            End If
        Next i
    End With 'CFOURIER
    lbMaxAt = j - 1 ' at point 1 is zero hertz so we have to correct by 1 to show freq
    lbValue = Int(Max) & ">"

    With picDisplay
        lbMaxAt.Left = .Left + .ScaleX(j, .ScaleMode, ScaleMode) - lbMaxAt.Width / 2
        .Cls
        .ScaleHeight = Max + 2

        picDisplay.PSet (0, .ScaleHeight / 1.05) 'start drawing outside picbox and a little above ground

        For i = 1 To NumSamples / 2 + 1
            'draw the result
            picDisplay.Line -(i, .ScaleHeight / 1.05 - cFourier.ComplexOut(i)), vbGreen
        Next i
    End With 'PICDISPLAY
    If InclReverse Then 'reverse transform

        With cFourier
            .TransformReverse = True
            For i = 1 To NumSamples
                If RemoveNoise Then
                    If i <> j Then 'j still has the point of max ampl
                        .RealIn(i) = 0
                        .ImagIn(i) = 0
                      Else 'NOT I...
                        .RealIn(i) = SvdReal(i) * NumSamples / Max '...and Max still has the max value
                        .ImagIn(i) = SvdImag(i) * NumSamples / Max
                    End If
                  Else 'REMOVENOISE = FALSE/0
                    .RealIn(i) = SvdReal(i)
                    .ImagIn(i) = SvdImag(i)
                End If

            Next i
        End With 'CFOURIER

        With picDisplay
            j = .ScaleHeight / 2
            picDisplay.PSet (0, j) 'start drawing outside picbox and at midpoint
            For i = 1 To NumSamples
                picDisplay.Line -(i / 2, j - cFourier.RealOut(i) * j / 4), vbRed 'the 4 is just an arbitrary value so that it fits nicely in the box
            Next i
        End With 'PICDISPLAY
    End If
    lbTiming = Format$(cFourier.Timing, "0.00") & " mSec"

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Apr-06 22:44)  Decl: 11  Code: 216  Total: 227 Lines
':) CommentOnly: 13 (5,7%)  Commented: 32 (14,1%)  Empty: 51 (22,5%)  Max Logic Depth: 6
