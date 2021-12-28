VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H003399FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quantum Mechanics"
   ClientHeight    =   8115
   ClientLeft      =   225
   ClientTop       =   915
   ClientWidth     =   13170
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000C0&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   13170
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   255
      Left            =   3840
      TabIndex        =   48
      Top             =   720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Max             =   100
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H003399FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   49
      Text            =   "Text4"
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C000C0&
      Caption         =   "Temizle"
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H003399FF&
      Caption         =   "Hassasiyet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   31
      Top             =   2760
      Width           =   1935
      Begin VB.OptionButton Option9 
         Appearance      =   0  'Flat
         BackColor       =   &H003399FF&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option8 
         Appearance      =   0  'Flat
         BackColor       =   &H003399FF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   720
         TabIndex        =   33
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option7 
         Appearance      =   0  'Flat
         BackColor       =   &H003399FF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H003399FF&
      Caption         =   "Deðiþen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   28
      Top             =   1080
      Width           =   1935
      Begin VB.OptionButton Option6 
         Appearance      =   0  'Flat
         BackColor       =   &H003399FF&
         Caption         =   "Phi(1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   1080
         TabIndex        =   30
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Appearance      =   0  'Flat
         BackColor       =   &H003399FF&
         Caption         =   "Enerji"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H003399FF&
      Caption         =   "Deðiþim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   1935
      Begin VB.OptionButton Option4 
         Appearance      =   0  'Flat
         BackColor       =   &H003399FF&
         Caption         =   "Azalt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   1080
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Appearance      =   0  'Flat
         BackColor       =   &H003399FF&
         Caption         =   "Artýr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C000C0&
      Caption         =   "Fonksiyon Deðerlerini Sýfýrla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7680
      Width           =   2895
   End
   Begin VB.CommandButton Command9 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Fonksiyonu Bul"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6960
      Width           =   2895
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H003399FF&
      Caption         =   "Çizgisel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3240
      TabIndex        =   19
      Top             =   7080
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H003399FF&
      Caption         =   "Noktasal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H008080FF&
      Caption         =   "Koordinat Eksenlerini Çiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H008080FF&
      Caption         =   "Potansiyelleri Çiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H003399FF&
      Caption         =   "Periyodik Potansiyel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   3000
      TabIndex        =   15
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Potansiyel Say"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      Width           =   1095
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   1080
      TabIndex        =   13
      Top             =   4920
      Width           =   495
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   600
      TabIndex        =   12
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF00FF&
      Caption         =   "Seçili Olaný Sil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00404040&
      Caption         =   "Çýk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C000C0&
      Caption         =   "Tümünü Sil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "Fonksiyonu Çiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7320
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Caption         =   "Potansiyel Gir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H003399FF&
      Caption         =   "Eðri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3120
      TabIndex        =   22
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H003399FF&
      Caption         =   "Potansiyel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   1680
      TabIndex        =   47
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9480
      TabIndex        =   51
      Top             =   7560
      Width           =   3375
   End
   Begin VB.Label Label17 
      BackColor       =   &H003399FF&
      Caption         =   "m=h²x10"
      Height          =   255
      Left            =   2280
      TabIndex        =   50
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2400
      TabIndex        =   41
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   46
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H003399FF&
      Caption         =   "Enerji Seviyesi(n)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   2280
      TabIndex        =   45
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   39
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2160
      TabIndex        =   42
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H003399FF&
      Caption         =   "Phi Min Deðer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   40
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   38
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H003399FF&
      Caption         =   "Phi Max Deðer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   37
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   36
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H003399FF&
      Caption         =   "Gerçek phi(1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   35
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   27
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H003399FF&
      Caption         =   "Gerçek Enerji"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H003399FF&
      Caption         =   "   V     Xo    Xs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H003399FF&
      Caption         =   "   Lo               Ls"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H003399FF&
      Caption         =   "Enerji="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label14 
      BackColor       =   &H003399FF&
      Caption         =   "Phi(max)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2400
      TabIndex        =   44
      Top             =   3000
      Width           =   975
   End
   Begin VB.Menu dosya 
      Caption         =   "&Dosya"
      Index           =   0
      Begin VB.Menu dosya_ac 
         Caption         =   "&Aç"
      End
      Begin VB.Menu dosya_kaydet 
         Caption         =   "&Kaydet"
      End
      Begin VB.Menu dosya_excel 
         Caption         =   "&Excel olarak kaydet"
      End
      Begin VB.Menu ayir1 
         Caption         =   "-"
      End
      Begin VB.Menu dosya_cikis 
         Caption         =   "&Çýkýþ"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Yardým"
      Begin VB.Menu help_phi 
         Caption         =   "&Phi deðerini öðren"
      End
      Begin VB.Menu help_x 
         Caption         =   "&X deðerine göre"
      End
      Begin VB.Menu degerler 
         Caption         =   "Deðe&rler"
         Begin VB.Menu asagi 
            Caption         =   "0-9999"
         End
         Begin VB.Menu yukarý 
            Caption         =   "10000-10199"
         End
      End
      Begin VB.Menu help_about 
         Caption         =   "&Hakkýnda"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As New Form2
Dim f3 As New Form3
Dim kontrol As Boolean
Dim az_art As Integer
Dim ne_degis As Integer
Dim hassasiyet As Integer
Dim hassasiyet2 As Long
Dim fark As Integer
Dim oynandi As Boolean



Private Sub asagi_Click()
hangisi = True
f3.Show
End Sub

Private Sub Check1_Click()
On Error GoTo hata_

If Not (Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "") Then

If List1.ListCount = 0 Then
  Check1.Value = 0
  ciz
  pot_ciz
  Command2_Click
  Exit Sub
 End If

'baþlangýç deðeri aralýkta olmalý
  If List2.List(0) < CDbl(Text2.Text) Or List3.List(List1.ListCount - 1) < CDbl(Text2.Text) Then
    MsgBox "Lo<Xo<Ls olmalý", vbOKOnly, "HATA"
    ciz
    pot_ciz
    Command2_Click
    Exit Sub
  End If
  
Dim sr
Dim i, j As Integer
Dim uz, uz2 As Double
Dim tane, say As Integer

If List1.ListCount > 0 Then
    If Check1.Value = 1 Then
        sr = MsgBox("periyod için bu deðerler alýnacak emin misiniz?", vbYesNo, "Period")
        
        If sr = vbYes Then
          
            say = List1.ListCount
            uz = Abs(CDbl(Text3.Text) - CDbl(Text2.Text))
            uz2 = Abs(CDbl(List3.List(say - 1)) - CDbl(Text2.Text))
         
            tane = (uz / uz2)
                If tane > o Then
                    For i = 1 To tane + 1
                        For j = 1 To say
                        
                          If (i * uz2 + List2.List(j - 1)) > CDbl(Text2.Text) And (i * uz2 + List2.List(j - 1)) < CDbl(Text3.Text) Then
                            List2.AddItem (CDbl(List2.List(j - 1)) + i * uz2)
                          Else
                            Check1.Value = 0
                            Exit Sub
                          End If
                          
                          If (i * uz2 + List3.List(j - 1)) > CDbl(Text3.Text) Then
                             List3.AddItem (CDbl(Text3.Text))
                             List1.AddItem (CDbl(List1.List(j - 1)))
                             Check1.Value = 0
                             ciz
                             pot_ciz
                             Exit Sub
                          Else
                            List3.AddItem (CDbl(List3.List(j - 1)) + i * uz2)
                            List1.AddItem (CDbl(List1.List(j - 1)))
                          End If
                          
                        Next
                    Next
                End If
        Else
        Check1.Value = 0
        ciz
        pot_ciz
        End If
    End If
Else
  Check1.Value = 0
  ciz
  pot_ciz
End If
Else
Check1.Value = 0
ciz
pot_ciz
End If
hata:
ciz
pot_ciz
Command2_Click
oynandi = True
Exit Sub
hata_:
  Check1.Value = 0
  ciz
  pot_ciz
End Sub
Private Sub Command1_Click()
On Error GoTo hata
Dim pt, ilk, son As Double
If Not IsNumeric(Text1.Text) Or Not IsNumeric(Text2.Text) Or Not IsNumeric(Text3.Text) Then Exit Sub
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" Then
    MsgBox "Önce Enerji ve aralýðý belirleyin", vbOKOnly, "HATA"
    ciz
    pot_ciz
    Exit Sub
End If
pt_gir:
pt = InputBox("potansiyel deðerini girin", "V")

        If Not IsNumeric(pt) Then
            MsgBox "Numerik deðerler girin", , "HATA"
            GoTo hata
            'Potansiyel deðerini sýnýrla
        ElseIf pt > 999 Or pt < 0 Then
            MsgBox "Potansiyel deðerlerini 0 ile 1000 arasýnda girin", vbOKOnly, "HATA"
            ciz
            pot_ciz
            GoTo pt_gir
        End If
ciz
pot_ciz
If pt <> "" Then
ilk_gir:
    ilk = InputBox("baþlangýç noktasýný girin", "Xo")
   
        If Not IsNumeric(ilk) Then
           MsgBox "Numerik deðerler girin", , "HATA"
            GoTo hata
            'Ýlk deðeri belirle
        ElseIf ilk < CDbl(Text2.Text) Or ilk > CDbl(Text3.Text) Then
           MsgBox "Ýlk deðeri " & Text2.Text & " ve " & Text3.Text & " arasýnda girin", vbOKOnly, "HATA"
           ciz
           pot_ciz
           GoTo ilk_gir
        End If
  ciz
  pot_ciz
    If ilk <> "" Then
geri:
ciz
pot_ciz
        son = InputBox("bitiþ noktasýný girin", "Xs")
      
         If Not IsNumeric(son) Then
           MsgBox "Numerik deðerler girin", , "HATA"
            GoTo hata
        ElseIf son > CDbl(Text3.Text) Or son < CDbl(Text2.Text) Or son <= ilk Then
            MsgBox "Son deðeri " & Text2.Text & " ve " & Text3.Text & " arasýnda ve " & ilk & " deðerinden büyük girin", vbOKOnly, "HATA"
           ciz
           pot_ciz
           GoTo geri:
        End If
        
        On Error GoTo hata
        
    Else
        GoTo hata
    End If
   
       List1.AddItem (pt)
       List2.AddItem (ilk)
       List3.AddItem (son)
       Command10_Click
       Label6.Caption = ""
       Label7.Caption = ""
       Label12.Caption = ""
       Label13.Caption = ""
       Label16.Caption = ""
       Label15.Caption = "Çukur ve tepe sayýsý"
Else
 GoTo hata
End If

hata:
ciz
pot_ciz
Command2_Click
 Exit Sub
End Sub

Private Sub Command10_Click()
Dim i As Integer
For i = 1 To 10199
    phi(i) = 0
Next
Command2.Enabled = False
End Sub

Private Sub Command12_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = "0"

List1.Clear
List2.Clear
List3.Clear

phi_min = 0
phi_max = 0

Label9.Caption = ""
Label10.Caption = ""
Label12.Caption = ""
Label13.Caption = ""

Label16.Caption = ""
enerji_seviyesi = 0

Option3.Value = True
Option5.Value = True
Option8.Value = True
Option1.Value = True
Check1.Value = 0

oynandi = False

Command10_Click
ciz
pot_ciz

 Form1.Scale (0, 0)-(form_ilk_x, form_ilk_y)
Form1.Refresh

Form1.Caption = "Quantum Mechanics"

End Sub

Private Sub Command2_Click() 'Çizim buraya yazýlacak
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
Dim dx, x_ As Double
dx = (CDbl(Text3.Text) - CDbl(Text2.Text)) / 10000
x_ = CDbl(Text2.Text) 'baþlangýç deðeri
Dim i, j As Integer

ciz
pot_ciz

  On Error GoTo hata

                '____________Fonksiyon______________

               If kontrol = True Then
                      'Noktasal çiz
                       For i = 1 To 10198
                            'nokta varsa çiz
                            
                          If i <= 9998 Then
                            PSet (x_, phi(i + 1)), vbBlack
                          Else  '10000 den büyük deðerler için
                            PSet (x_, phi(i + 1)), vbRed
                          End If
                          
                         x_ = x_ + dx
                       Next

                Else
                      'Doðrusal çiz
                      For i = 1 To 10198
                        
                        If i <= 9998 Then
                          Line (x_ - dx, phi(i - 1))-(x_, phi(i)), vbBlack
                        Else
                            Line (x_ - dx, phi(i - 1))-(x_, phi(i)), vbRed
                        End If
                          x_ = x_ + dx
                      Next
               End If
enerji_seviyesi_say
Option1.Value = kontrol
hata:
End If
End Sub

Private Sub Command3_Click()
If List1.ListCount > 0 Then
    Dim i
    i = MsgBox("silmek istediðinizden emin misiniz?", vbOKCancel, "Sil")
    If i = vbOK Then
        List1.Clear
        List2.Clear
        List3.Clear
        UpDown1_Change
        Label15.Caption = "Enerji Seviyesi(n)"
        Label16.Caption = ""
        ciz
        pot_ciz
    End If
End If
End Sub

Private Sub Command4_Click()
    Dim i
    i = MsgBox("çýkmak istediðinizden emin misiniz?", vbOKCancel, "Çýkýþ")
    If i = vbOK Then
        f.Cls
         End
    End If
    ciz
    pot_ciz
    If Not (Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "") Then
      Command2_Click
     End If
End Sub

Private Sub Command5_Click()
 If List1.ListCount > 0 Then
    If List1.ListIndex > -1 Then
        List1.RemoveItem (List1.ListIndex)
        List2.RemoveItem (List2.ListIndex)
        List3.RemoveItem (List3.ListIndex)
        ciz
        pot_ciz
    End If
 End If
End Sub

Private Sub Command6_Click()
If Not IsNumeric(Text1.Text) Or Not IsNumeric(Text2.Text) Or Not IsNumeric(Text3.Text) Then Exit Sub
If List1.ListCount > 0 Then
MsgBox List1.ListCount, vbOKOnly, "Potansiyel Sayýsý"
Else
 MsgBox "Hiç bir potansiyel yok", vbOKOnly, "Potansiyel Sayýsý"
End If
End Sub

Private Sub Command7_Click()
   ciz
   pot_ciz
End Sub

Private Sub Command8_Click()
If Not IsNumeric(Text1.Text) Or Not IsNumeric(Text2.Text) Or Not IsNumeric(Text3.Text) Then Exit Sub
ciz
If CDbl(Text2.Text) * CDbl(Text3.Text) > 0 Then 'her ikisi ayný bölgede ise
    Line (CDbl(Text2.Text), 0)-(CDbl(Text3.Text), 0), vbBlack
End If
End Sub

Private Sub Command9_Click()
On Error GoTo hata_
 If Not IsNumeric(Text1.Text) Or Not IsNumeric(Text2.Text) Or Not IsNumeric(Text3.Text) Then Exit Sub
Command2.Enabled = False
If Not (Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "") Then
Dim sol_X, sag_X, ust_Y, alt_Y As Double

Label12.Caption = ""

Label16.Caption = ""

sol_X = CDbl(Text2.Text)
ust_Y = CDbl(Text1.Text)
sag_X = CDbl(Text3.Text)
alt_Y = -CDbl(Text1.Text)

Dim dx, x_ As Double
dx = (CDbl(Text3.Text) - CDbl(Text2.Text)) / 10000

Dim i, j, m, n As Integer
Dim E, E2 As Double
E = CDbl(Text1.Text)
E2 = E
phi(0) = 0
phi(1) = 1 / (10 ^ hassasiyet)

pot_max_hesapla


If Text1.Text = "" Then
    MsgBox "Enerji Deðerini girin", , "Hata"
Else
    If Text2.Text = "" Then
        MsgBox "Ýlk Deðeri girin", , "Hata"
    Else
        If Text3.Text = "" Then
             MsgBox "Son Deðeri girin", , "Hata"
        Else
                'text2 text3 den büyükse
            If CDbl(Text2.Text) >= CDbl(Text3.Text) Then
                MsgBox "Ls > Lo olmalý", vbOKOnly, "HATA"
                Exit Sub
            End If
            '_________________Yanlýþ deðerli potansiyelleri sil_____________________
             If List1.ListCount > 0 Then
               For m = 0 To List1.ListCount - 1
                 
                 If ((List2.List(m) > CDbl(Text3.Text)) Or (List2.List(m) < CDbl(Text2.Text))) Then
                    List1.RemoveItem (m)
                    List2.RemoveItem (m)
                    List3.RemoveItem (m)
                 End If
                 
                  If ((List3.List(m) > CDbl(Text3.Text)) Or (List3.List(m) < CDbl(Text2.Text))) Then
                    List1.RemoveItem (m)
                    List2.RemoveItem (m)
                    List3.RemoveItem (m)
                 End If
                 
               Next
             End If
            '_______________________________________________________________________
            x_ = sol_X 'baþlangýç deðeri
            max = 0
            min = 0
             
             '____________Fonksiyon______________
            Do
            For i = 1 To 10198
             If List1.ListCount = 0 Then  'hiç bir potansiyel yoksa
              phi(i + 1) = 2 * phi(i) - phi(i - 1) - 2 * (10 ^ (CInt(Text4.Text))) * (dx) ^ 2 * (E2) * phi(i)
               
              Else   'potansiyel varsa
               For j = 0 To List1.ListCount - 1
                    If x_ < List3.List(j) And x_ >= List2.List(j) Then
                    '***********************************
                      If E2 >= List1.List(j) Then   'Enerji Potansiyelden büyükse
                        phi(i + 1) = 2 * phi(i) - phi(i - 1) - 2 * (10 ^ (CInt(Text4.Text))) * (dx) ^ 2 * (E2 - List1.List(j)) * phi(i)
                      Else                           'Aksi halde
                        phi(i + 1) = 2 * phi(i) - phi(i - 1) - 2 * (10 ^ (CInt(Text4.Text))) * (dx) ^ 2 * (List1.List(j) - E2) * phi(i)
                      End If
                     '***********************************
                    Else
                      phi(i + 1) = 2 * phi(i) - phi(i - 1) - 2 * (10 ^ (CInt(Text4.Text))) * (dx) ^ 2 * (E2) * phi(i)
                    End If
                Next
              End If
              x_ = x_ + dx
            Next
                
                If ne_degis = 1 Then
                    E2 = E2 + az_art * E * 0.05
                Else
                    phi(1) = phi(1) + az_art * (1 / (10 ^ (hassasiyet + 1)))
                End If
            
             If ne_degis = 1 Then
                If E2 <= -2 * E Or E2 >= 2 * E Then
                  Label12.Caption = "Deðer Bulunamadý"
                  Label6.Caption = E2
                  Label7.Caption = phi(1)
                  Label13.Caption = phi(9999)
                  Command2.Enabled = True
                  ciz
                  pot_ciz
                  Exit Sub
                End If
             Else
                If phi(9999) > 1 / (10 ^ (hassasiyet - 2)) Or phi(9999) < -1 / (10 ^ (hassasiyet - 2)) Then
                    Label12.Caption = "Deðer yaklaþýk olarak bulundu"
                    Label6.Caption = E2
                    Label7.Caption = phi(1)
                    Label13.Caption = phi(9999)
                    Command2.Enabled = True
                    ciz
                    pot_ciz
                  Exit Sub
                End If
            End If
            
            If i > hassasiyet2 Then
               If phi(i) < 1 / (10 ^ (hassasiyet - 2)) And phi(i) > -1 / (10 ^ (hassasiyet - 2)) Then
                    Label12.Caption = "Deðer yaklaþýk olarak bulundu"
                    Label6.Caption = E2
                    Label7.Caption = phi(1)
                    Command2.Enabled = True
              'GERÝSÝNÝ HESAPLA VE ÇIK
             '__________________________________________________________________
                 For i = hassasiyet2 + 1 To 10198
                      If List1.ListCount = 0 Then  'hiç bir potansiyel yoksa
                        phi(i + 1) = 2 * phi(i) - phi(i - 1) - 2 * (10 ^ (CInt(Text4.Text))) * (dx) ^ 2 * (E2) * phi(i)
               
                      Else   'potansiyel varsa
                        For j = 0 To List1.ListCount - 1
                          If x_ < List3.List(j) And x_ >= List2.List(j) Then
                    '***********************************
                             If E2 >= List1.List(j) Then
                                  phi(i + 1) = 2 * phi(i) - phi(i - 1) - 2 * (10 ^ (CInt(Text4.Text))) * (dx) ^ 2 * (E2 - List1.List(j)) * phi(i)
                             Else
                                 phi(i + 1) = 2 * phi(i) - phi(i - 1) - 2 * (10 ^ (CInt(Text4.Text))) * (dx) ^ 2 * (List1.List(j) - E2) * phi(i)
                             End If
                     '***********************************
                          Else
                            phi(i + 1) = 2 * phi(i) - phi(i - 1) - 2 * (10 ^ (CInt(Text4.Text))) * (dx) ^ 2 * (E2) * phi(i)
                          End If
                        Next
                    End If
                    x_ = x_ + dx
                 Next
             '_________________________________________________________________
                 Label13.Caption = phi(9999)
                     ciz
                    pot_ciz
                 Exit Sub
               End If
            End If
      Loop
            '_________________________________________
        End If
    End If
End If

'Max ve Min deðerleri bul
    For i = 0 To 9999
        If phi(i) > max Then max = phi(i)
        If phi(i) < min Then min = phi(i)
     Next
Command2.Enabled = True
End If
        phi_max_bul
        phi_min_bul
        Label9.Caption = phi_max
        Label10.Caption = phi_min
          Exit Sub
hata_:
  MsgBox "Fonksiyon çok büyük deðerler alýyor,yada yanlýþ deðerler girdiniz!" + Chr(10) + "Potansiyel deðerleri ve m kütlesi için." _
  + Chr(10) + "Þimdi tekrar deneyin.", vbOKOnly, "HATA"
  Text4.Text = 0
  Command10_Click
  ciz
  pot_ciz
End Sub



Private Sub dosya_ac_Click()
On Error GoTo hata
Dim yer As String
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "Aç....Bir Quantum(qtm) dosyasý seçin."
CommonDialog1.Filter = "Quantum|*.qtm|Hepsi|*.*||"
CommonDialog1.DefaultExt = "qtm"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen

yer = CommonDialog1.FileName

'Açma iþlemi
    If yer <> "" Then

        If Right(yer, 4) = ".qtm" Then
            
            List1.Clear
            List2.Clear
            List3.Clear
            
            txt_ac (yer)
        Else
            MsgBox "Uzantýyý qtm olarak seçin", vbOKOnly, "HATA"
        End If
    End If
hata:
End Sub

Private Sub dosya_cikis_Click()
    Command4_Click
End Sub

Private Sub dosya_excel_Click()
On Error GoTo hata
CommonDialog1.CancelError = True
If Not (Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "") Then

CommonDialog1.DialogTitle = "Kaydet....Bir excel veya text dosyasý seçin"
CommonDialog1.Filter = "Excel Dosyasý|*.xls|Text Dosyasý|*.txt|Hepsi|*.*||"
CommonDialog1.DefaultExt = "xls"
CommonDialog1.FileName = ""
CommonDialog1.ShowSave

yer = CommonDialog1.FileName

'Kaydetme iþlemi
    If yer <> "" Then
     
     If Dir(yer) = "" Then
tamam:
            If Right(yer, 4) = ".xls" Or Right(yer, 4) = ".txt" Then
                excel_kaydet (yer)
            Else
                MsgBox "Uzantýyý xls veya txt olarak seçin", vbOKOnly, "HATA"
             End If
     ElseIf Dir(yer) <> "" Then
            
            Dim sor
            sor = MsgBox("Dosya zaten var üzerine yazmak ister misiniz?", vbYesNo, "Kaydet")
            
            If sor = vbYes Then
                Kill (yer)
                GoTo tamam
            End If
     Else
           Exit Sub
     End If
     
    End If
   Else
    MsgBox "Deðerler eksik", vbOKOnly, "HATA"
   End If
hata:
End Sub

Private Sub dosya_kaydet_Click()
On Error GoTo hata
CommonDialog1.CancelError = True
If Not (Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "") Then

CommonDialog1.DialogTitle = "Kaydet....Bir Quantum(qtm) dosyasý seçin"
CommonDialog1.Filter = "Quantum Dosyasý|*.qtm|Hepsi|*.*||"
CommonDialog1.DefaultExt = "qtm"
CommonDialog1.FileName = ""
CommonDialog1.ShowSave

yer = CommonDialog1.FileName

'Kaydetme iþlemi
    If yer <> "" Then
    
     If Dir(yer) = "" Then
tamam:
            If Right(yer, 4) = ".qtm" Then
                txt_kaydet (yer)
            Else
                MsgBox "Uzantýyý qtm olarak seçin", vbOKOnly, "HATA"
             End If
     ElseIf Dir(yer) <> "" Then
            
            Dim sor
            sor = MsgBox("Dosya zaten var üzerine yazmak ister misiniz?", vbYesNo, "Kaydet")
            
            If sor = vbYes Then
                Kill (yer)
                GoTo tamam
            End If
     Else
           Exit Sub
     End If
    End If
   Else
    MsgBox "Deðerler eksik", vbOKOnly, "HATA"
   End If
hata:
End Sub

Private Sub Form_Activate()
f.Hide
f3.Hide
f3.List1.Clear
End Sub

Private Sub Form_Click()
 Form_DblClick
End Sub

Private Sub Form_DblClick()
ciz
pot_ciz
  If Not (Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "") Then
    Command2_Click
  End If
End Sub
Private Sub Form_Load()

UpDown1.BuddyControl = Text4
UpDown1.BuddyProperty = "Text"
Text4.Text = 0

fark = 0
hassasiyet2 = 9600
Option1.Value = True
Option3.Value = True
Option5.Value = True
Option8.Value = True

Label18.Visible = False

az_art = 1
ne_degis = 1

kontrol = True
Command2.Enabled = False
oynandi = False

enerji_seviyesi = 0

max = 0
min = 0
phi_max = 0
phi_min = 0

form_ilk_x = Me.ScaleWidth
form_ilk_y = Me.ScaleHeight
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo hata_
    If Not (Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "") Then
     If IsNumeric(Text1.Text) And IsNumeric(Text2.Text) And IsNumeric(Text3.Text) Then
       If CDbl(Text2.Text) < CDbl(Text3.Text) Then
         Dim dx_ As Double
         dx_ = (CDbl(Text3.Text) - CDbl(Text2.Text)) / 10000
         Dim xx As Long
        
         
       If X <= CDbl(Text3.Text) And X >= CDbl(Text2.Text) Then
           xx = (Abs(X - CDbl(Text2.Text))) / dx_
         Form1.Caption = "Quantum Mechanics" & "   X=" & X & " Y=" & Y & "    " & xx & " inci adým  ----->  phi(" & xx & ")=" & phi(xx - 1)
          Label18.Visible = False
       Else
         If X > CDbl(Text3.Text) Then
           xx = (Abs(X - CDbl(Text3.Text))) / dx_
            If xx < 201 Then
              Label18.Left = X - Form1.ScaleWidth / (3.5)
              Label18.Top = Y
              
              If CDbl(Text3.Text) > CDbl(Text2.Text) Then
               If Command2.Enabled = True Then Label18.Visible = True
              End If
              
              Label18.Caption = " phi(" & xx + 9999 & ")=" & phi(xx + 9999)
            Else
              Label18.Visible = False
            End If
       Else
             Label18.Visible = False
         End If
           Form1.Caption = "Quantum Mechanics" & "   X=" & X & " Y=" & Y
       End If
        Else
          Form1.Caption = "Quantum Mechanics"
        End If
      Else
         Form1.Caption = "Quantum Mechanics" & "   X=" & X & " Y=" & Y
      End If
    Else
      Form1.Caption = "Quantum Mechanics" & "   X=" & X & " Y=" & Y
    End If

hata_:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If List1.ListCount > 0 Then
    If oynandi = True Then
     Dim i
        i = MsgBox("Verilerinizi kaydetmek ister misiniz? ", vbYesNo, "kaydet?")
        If i = vbYes Then
            dosya_kaydet_Click
        Else
          'Form_Unload (2332) 'Bu kýsmý anlamadým
        Exit Sub
        End If
      End If
    'End If
    f.Cls
    End
End Sub


Private Sub help_about_Click()
 f.Show
End Sub

Private Sub help_phi_Click()
Dim i As Long
On Error GoTo hata
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
      MsgBox "Deðerler eksik!", vbOKOnly, "HATA"
      Exit Sub
End If
i = InputBox("Kaçýncý(n) phi deðeri", "Phi(n)")

Dim fark, xx As Double
fark = (CDbl(Text3.Text) - CDbl(Text2.Text)) / 9999
xx = i * fark + CDbl(Text2.Text)

If i >= 0 And i <= 10199 Then
    MsgBox "phi(" & i & ")=" & phi(i), vbOKOnly, "X=" & xx & ", n=" & i
Else
hata:
    MsgBox "0 ile 10199 arasýnda bir sayý girin", vbOKOnly, "HATA"
    ciz
    pot_ciz
    Command2_Click
End If
End Sub

Private Sub help_x_Click()
Dim i As Double
On Error GoTo hata

If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
      MsgBox "Deðerler eksik!", vbOKOnly, "HATA"
      Exit Sub
End If

i = InputBox("X deðerini girin", "X")
If i >= CDbl(Text2.Text) And i <= CDbl(Text3.Text) Then
    Dim fark As Double
    Dim kac As Long
    Dim gercek_i As Double
    
    fark = (CDbl(Text3.Text) - CDbl(Text2.Text)) / 10000
    kac = (i - CDbl(Text2.Text)) / fark
    gercek_i = CDbl(Text2.Text) + kac * fark
    
    If kac = 0 Then
        MsgBox "Phi_X(" & i & ")= " & phi(0), vbOKOnly, "Phi(0)=Phi_X(" & i & ")"
        Exit Sub
    End If
        
    If i <> gercek_i Then
      MsgBox "Phi(" & kac & ") = " & "Phi_X(" & gercek_i & ")" & Chr(10) _
      & "Phi_X(" & gercek_i & ")=" & phi(kac), vbOKOnly, _
      "Phi_X(" & i & ") deðeri yaklaþýk olarak " & "Phi(" & kac & ")"
    Else
      MsgBox "Phi_X(" & gercek_i & ")=" & phi(kac - 1), vbOKOnly, _
      "Phi_X(" & i & ") = " & "Phi(" & kac - 1 & ")"
    End If
    
Else
hata:
    MsgBox CDbl(Text2.Text) & " ile " & CDbl(Text3.Text) & " arasýnda bir sayý girin", vbOKOnly, "HATA"
    ciz
    pot_ciz
    Command2_Click
End If
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
End Sub

Private Sub List1_DblClick()
MsgBox "V=" & List1.List(List1.ListIndex) & "  ,Xo=" & List2.List(List2.ListIndex) & "  ,Xs=" & List3.List(List3.ListIndex), vbOKOnly, CStr(List1.ListIndex + 1) + " . potansiyel"
End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
End Sub

Private Sub List2_DblClick()
MsgBox "V=" & List1.List(List1.ListIndex) & "  ,Xo=" & List2.List(List2.ListIndex) & "  ,Xs=" & List3.List(List3.ListIndex), vbOKOnly, CStr(List1.ListIndex + 1) + " . potansiyel"
End Sub

Private Sub List3_Click()
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
End Sub

Private Sub List3_DblClick()
MsgBox "V=" & List1.List(List1.ListIndex) & "  ,Xo=" & List2.List(List2.ListIndex) & "  ,Xs=" & List3.List(List3.ListIndex), vbOKOnly, CStr(List1.ListIndex + 1) + " . potansiyel"
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then kontrol = True
Command2_Click
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then kontrol = False
Command2_Click
End Sub

Private Sub Option3_Click()
az_art = 1
ciz
pot_ciz
Label12.Caption = ""
oynandi = True
Command10_Click
End Sub

Private Sub Option4_Click()
az_art = -1
ciz
pot_ciz
Label12.Caption = ""
oynandi = True
Command10_Click
End Sub

Private Sub Option5_Click()
ne_degis = 1
ciz
pot_ciz
Label12.Caption = ""
oynandi = True
Command10_Click
End Sub

Private Sub Option6_Click()
ne_degis = -1
ciz
pot_ciz
Label12.Caption = ""
oynandi = True
Command10_Click
End Sub

Private Sub Option7_Click()
 If Option7.Value = True Then
    hassasiyet = 3 - fark
    hassasiyet2 = 9500
 End If
 
 Label6.Caption = ""
 Label7.Caption = ""
 Label12.Caption = ""
 ciz
pot_ciz
oynandi = True
Command10_Click
End Sub

Private Sub Option8_Click()
 If Option8.Value = True Then
    hassasiyet = 4 - fark
    hassasiyet2 = 9600
End If

 Label6.Caption = ""
 Label7.Caption = ""
 Label12.Caption = ""
 ciz
pot_ciz
oynandi = True
Command10_Click
End Sub

Private Sub Option9_Click()
 If Option9.Value = True Then
    hassasiyet = 5 - fark
    hassasiyet2 = 9700
End If

 Label6.Caption = ""
 Label7.Caption = ""
 Label12.Caption = ""
 ciz
pot_ciz
oynandi = True
Command10_Click
End Sub

Private Sub Text1_Change()
On Error GoTo hata
        If Not IsNumeric(CDbl(Text1.Text)) Then
            Text1.Text = ""
        Else
            If CDbl(Text1.Text) < 0 Or CDbl(Text1.Text) > 999 Then
            MsgBox "Enerji deðerini 0 ile 1000 arasýnda seçin", vbOKOnly, "HATA"
hata:
                Text1.Text = ""
            End If
        End If

Label6.Caption = ""
Label7.Caption = ""
Label12.Caption = ""
Label13.Caption = ""
Label16.Caption = ""
ciz
pot_ciz

Command10_Click
Command2.Enabled = False
oynandi = True
End Sub

Private Sub Text2_Change()
On Error GoTo hata
If Len(Trim(Text2.Text)) > 1 Then
        If Not IsNumeric(CDbl(Text2.Text)) Then
            Text2.Text = ""
        Else
            If CDbl(Text2.Text) < -100 Or CDbl(Text2.Text) > 100 Then
            MsgBox "Baþlangýç deðerini -100 ile 100 arasýnda seçin", vbOKOnly, "HATA"
                Text2.Text = ""
            End If
        End If
End If

If Len(Trim(Text2.Text)) = 1 Then
    If Not (Text2.Text = "-" Or IsNumeric(Text2.Text)) Then
hata:
        Text2.Text = ""
    End If
End If

Label6.Caption = ""
Label7.Caption = ""
Label12.Caption = ""
Label13.Caption = ""
Label16.Caption = ""

ciz
pot_ciz

'Herhangi bir deðiþiklikte phi deðerlerini sýfýrla
Command10_Click
Command2.Enabled = False
oynandi = True
End Sub

Private Sub Text3_Change()
On Error GoTo hata
If Len(Trim(Text3.Text)) > 1 Then
        If Not IsNumeric(CDbl(Text3.Text)) Then
            Text3.Text = ""
        Else
            If CDbl(Text3.Text) < -100 Or CDbl(Text3.Text) > 100 Then
            MsgBox "Bitiþ deðerini -100 ile 100 arasýnda seçin", vbOKOnly, "HATA"
                Text3.Text = ""
            End If
        End If
End If

If Len(Trim(Text3.Text)) = 1 Then
    If Not (Text3.Text = "-" Or IsNumeric(Text3.Text)) Then
hata:
        Text3.Text = ""
    End If
End If

Label6.Caption = ""
Label7.Caption = ""
Label12.Caption = ""
Label13.Caption = phi(9999)
Label16.Caption = ""

ciz
pot_ciz

'Herhangi bir deðiþiklikte phi deðerlerini sýfýrla
Command10_Click
Command2.Enabled = False
oynandi = True
End Sub
Private Sub ciz()

  On Error GoTo hata_
  
    Dim i As Integer
   Form1.Cls
   If Not IsNumeric(Text1.Text) Or Not IsNumeric(Text2.Text) Or Not IsNumeric(Text3.Text) Then Exit Sub
   If Not (Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "") Then
        
        Dim sol_X, sol_Y, sag_X, sag_Y As Double
        
        bas_X = -(2 / 3) * (Abs(Text2.Text) + Abs(Text3.Text)) + CDbl(Text2.Text)
        sol_X = CDbl(Text2.Text)
        sag_X = CDbl(Text3.Text)
        ust_Y = CDbl(Text1.Text)
        ust_Y = -alt_Y
    If CDbl(Text2.Text) > CDbl(Text3.Text) Then Exit Sub
    
    '________________________________
     'Fonsiyonun max ve min  deðerlerine göre çizdir
     
     For i = 0 To 9999
      If phi(i) > max Then max = phi(i)
      If phi(i) < min Then min = phi(i)
     Next
     
     
        If max > ust_Y And max < 1000 Then
            ust_Y = max
        Else
            ust_Y = Abs(CDbl(Text1.Text))
        End If
        
        
        If min < alt_Y And min > -1000 Then
                alt_Y = min
        ElseIf min = 0 Then
               alt_Y = 0
        Else
            alt_Y = -Abs(CDbl(Text1.Text))
        End If
        phi_max_bul
        phi_min_bul
        Label9.Caption = phi_max
        Label10.Caption = phi_min
    '___________________________________________
       If min <> 0 Then
         Form1.Scale (bas_X, ust_Y * 1.05)-(sag_X, alt_Y * 1.05)
       Else    'alt tarafý yoksa
         Form1.Scale (bas_X, ust_Y * 1.05)-(sag_X, -ust_Y * 0.05)
       End If
        
        Form1.ScaleWidth = 1.05 * Form1.ScaleWidth
        'Form1.ScaleHeight = 1.05 * Form1.ScaleHeight
        Form1.Refresh
     'Koordinatlar
        'Line (-(2 / 3) * (Abs(Text2.Text) + Abs(Text3.Text)), 0)-(sag_X, 0)
        'Alttaki satýr yerine
        Line (sol_X, 0)-(sag_X, 0), vbBlack
        
    
    '___________________******^^^^^^^^^^********_____________________________
    If CDbl(Text2.Text) * CDbl(Text3.Text) < 0 Then 'her ikisi farklý iþarette ise
        Line (0, ust_Y)-(0, alt_Y), vbBlack
    Else
       bas_X = CDbl(Text2.Text) - (2 / 3) * (CDbl(Text3.Text) - CDbl(Text2.Text))
       Dim nn As Integer
       
         If min <> 0 Then
            Form1.Scale (bas_X, ust_Y * 1.05)-(sag_X, alt_Y * 1.05)
         Else    'alt tarafý yoksa
            Form1.Scale (bas_X, ust_Y * 1.05)-(sag_X, -ust_Y * 0.05)
         End If
         
         Form1.ScaleWidth = 1.05 * Form1.ScaleWidth
         Form1.Refresh
          If List1.ListCount > 0 Then
            For nn = 0 To List1.ListCount - 1
                Line (CDbl(List2.List(nn)), 0)-(CDbl(List3.List(nn)), 0), vbBlack
            Next
          End If
    End If
    '____________________******^^^^^^^^^^********___________________________
    
    
    'her ikisi pozitif veya negatif olursa arasý boþ kalmasýn diye
        'Line (0, 0)-(X_sol, 0)
        'Line (0, 0)-(X_sag, 0)
        
     'Sonsuz potansiyel kuyusu
        Line (sol_X, ust_Y)-(sol_X, 0), vbGreen
        Line (sag_X, ust_Y)-(sag_X, 0), vbGreen
     'Sonsuz potansiyel kuyusu duvarlarýndan aþaðý
      If min <> 0 Then
        Line (sol_X, 0)-(sol_X, -ust_Y), vbBlack
        Line (sag_X, 0)-(sag_X, -ust_Y), vbBlack
      End If
    End If
hata_:
End Sub
Private Sub pot_ciz()

 If Not IsNumeric(Text1.Text) Or Not IsNumeric(Text2.Text) Or Not IsNumeric(Text3.Text) Then Exit Sub
 If Not (Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "") Then
    Dim temp As Double
    temp = CDbl(Text2.Text)
    Dim i As Long

    For i = 0 To List1.ListCount - 1
        Line (temp, 0)-(List2.List(i), 0), vbGreen
        Line (List2.List(i), 0)-(List2.List(i), List1.List(i)), vbYellow
        Line (List2.List(i), List1.List(i))-(List3.List(i), List1.List(i)), vbRed
        Line (List3.List(i), List1.List(i))-(List3.List(i), 0), vbYellow
        temp = List3.List(i)
    Next
    Line (temp, 0)-(CDbl(Text3.Text), 0), vbGreen
 End If
End Sub
Private Sub txt_ac(yer As String)
On Error GoTo hata:

Open yer For Random As #1 Len = 10 'Dosyayý okumak için aç
Dim i As Integer
Dim t1, t2, t3 As Double
Dim op5, op6, op3, op4, op7, op8, op9 As Boolean
Dim pot_sayisi As Integer
Dim temp1, temp2, temp3 As Double
Dim txt4_ust As Integer
Dim s As Double

Get #1, 1, s
Get #1, 2, t1
Get #1, 3, t2
Get #1, 4, t3

Get #1, 5, op5
Get #1, 6, op6
Get #1, 7, op3
Get #1, 8, op4
Get #1, 9, op7
Get #1, 10, op8
Get #1, 11, op9
Get #1, 12, pot_sayisi
Get #1, 13, txt4_ust

If s = 571632 Then
    Text1.Text = t1
    Text2.Text = t2
    Text3.Text = t3
    
    
 Option5.Value = op5
 Option6.Value = op6
 Option3.Value = op3
 Option4.Value = op4
 Option7.Value = op7
 Option8.Value = op8
 Option9.Value = op9
 Text4.Text = txt4_ust

    If pot_sayisi > 0 Then
       For i = 0 To pot_sayisi - 1
  
        Get #1, i * 3 + 14, temp1
        Get #1, i * 3 + 15, temp2
        Get #1, i * 3 + 16, temp3
    
        List1.List(i) = temp1
        List2.List(i) = temp2
        List3.List(i) = temp3
     Next
   End If
    Label18.Visible = False
    Command9_Click  'Fonksiyonu hesapla
    Command2_Click   'Fonksiyonu çiz
    Command2.Enabled = True
    Check1.Value = 0
    oynandi = False
Else
 GoTo hata:
End If
Close #1
Exit Sub

hata:
 MsgBox "Dosya bozuk veya eksik!", vbOKOnly, "HATA"
 Close #1
End Sub
Private Sub txt_kaydet(yer As String)
Open yer For Random As #1 Len = 10
Dim i As Integer
Dim t1, t2, t3 As Double
Dim op5, op6, op3, op4, op7, op8, op9 As Boolean
Dim pot_sayisi As Integer
Dim temp1, temp2, temp3 As Double
Dim txt4_ust As Integer
Dim s As Double
s = 571632

t1 = Text1.Text
t2 = Text2.Text
t3 = Text3.Text

op5 = Option5.Value
op6 = Option6.Value
op3 = Option3.Value
op4 = Option4.Value
op7 = Option7.Value
op8 = Option8.Value
op9 = Option9.Value
pot_sayisi = List1.ListCount  'potansiyel sayýsýný verir
txt4_ust = Text4.Text        '10 nun derecesini verir

Put #1, 1, s
Put #1, 2, t1
Put #1, 3, t2
Put #1, 4, t3

Put #1, 5, op5
Put #1, 6, op6
Put #1, 7, op3
Put #1, 8, op4
Put #1, 9, op7
Put #1, 10, op8
Put #1, 11, op9
Put #1, 12, pot_sayisi   ' potansiyel sayýsý
Put #1, 13, txt4_ust

If pot_sayisi > 0 Then
  For i = 0 To pot_sayisi - 1
    
    temp1 = List1.List(i)
    temp2 = List2.List(i)
    temp3 = List3.List(i)
    
    Put #1, i * 3 + 14, temp1
    Put #1, i * 3 + 15, temp2
    Put #1, i * 3 + 16, temp3
  Next
End If

Close #1
End Sub
Private Sub phi_max_bul()
    Dim i As Integer
    Dim temp As Double
    temp = 0
    For i = 0 To 9999
        If temp < phi(i) Then temp = phi(i)
    Next
    phi_max = temp
End Sub

Private Sub phi_min_bul()
    Dim i As Integer
    Dim temp As Double
    temp = 0
    For i = 0 To 9999
        If temp > phi(i) Then temp = phi(i)
    Next
    phi_min = temp
End Sub
Private Sub pot_max_hesapla()
    Dim i As Integer

    If List1.ListCount > 0 Then
        For i = 0 To List1.ListCount - 1
            If List1.List(i) > pot_max Then pot_max = List1.List(i)
        Next
    End If
End Sub
 
 Private Sub enerji_seviyesi_say()
  Dim i As Integer
  enerji_seviyesi = 0
      'ENERJÝ SEVÝYESÝ SAY,yani fonk nun eðriyi kestiði noktalar
     For i = 1 To hassasiyet2
        If Math.Sgn(phi(i)) <> Math.Sgn(phi(i + 1)) Then
           enerji_seviyesi = enerji_seviyesi + 1
        End If
     Next
     If Label12.Caption = "Deðer Bulunamadý" Then
        Label15.Caption = "Enerji Seviyesi(n)"
        Label16.Caption = ""
     Else
       If List1.ListCount = 0 Then
        Label15.Caption = "Enerji Seviyesi(n)"
        Label16.Caption = enerji_seviyesi + 1
       Else
        Label15.Caption = "Çukur ve tepe sayýsý"
        Label16.Caption = enerji_seviyesi + 1
       End If
    End If
 End Sub
 
 Private Sub excel_kaydet(yer As String)
 
  Open yer For Output As #1
   Dim i As Integer
   Dim s As String

  For i = 1 To 10200
    Print #1, i, Chr(9), CDbl(phi(i - 1))
  Next
  oynandi = False
   Close #1
   
 End Sub

Private Sub UpDown1_Change()

List1.Clear
List2.Clear
List3.Clear

phi_min = 0
phi_max = 0

Label6.Caption = ""
Label7.Caption = ""

Label9.Caption = ""
Label10.Caption = ""
Label12.Caption = ""
Label13.Caption = ""

Label16.Caption = ""
enerji_seviyesi = 0

Option3.Value = True
Option5.Value = True
Option8.Value = True
Option1.Value = True
Check1.Value = 0

oynandi = False

Command10_Click
ciz
pot_ciz

 Form1.Scale (0, 0)-(form_ilk_x, form_ilk_y)
Form1.Refresh

Form1.Caption = "Quantum Mechanics"

End Sub

Private Sub yukarý_Click()
hangisi = False
f3.Show
End Sub
