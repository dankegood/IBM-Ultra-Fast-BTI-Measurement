VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form WGFMUVBInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fast_BTI (w/ WGFMU)"
   ClientHeight    =   10125
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   13860
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   13860
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame BottomFrame 
      Height          =   3975
      Left            =   0
      TabIndex        =   285
      Top             =   6048
      Width           =   13572
      Begin VB.TextBox YCoord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5184
         TabIndex        =   705
         Text            =   "Y"
         Top             =   1740
         Width           =   400
      End
      Begin VB.TextBox XCoord 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4752
         TabIndex        =   703
         Text            =   "X"
         Top             =   1740
         Width           =   400
      End
      Begin VB.TextBox ChipIDLabel 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4752
         TabIndex        =   702
         Text            =   "ChipID"
         Top             =   1320
         Width           =   800
      End
      Begin VB.TextBox NumDevWGFMU 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4320
         TabIndex        =   290
         Text            =   "5"
         Top             =   300
         Width           =   612
      End
      Begin VB.TextBox CommnGateNumDevWGFMU 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4800
         TabIndex        =   296
         Text            =   "2"
         Top             =   780
         Width           =   612
      End
      Begin VB.OptionButton DUTStressOptionCommonGate 
         Caption         =   "Common-gate device to stress (2<=n<=9) = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   120
         TabIndex        =   297
         Top             =   804
         Width           =   4935
      End
      Begin VB.CommandButton ChooseDir 
         Caption         =   "Select a working folder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   295
         Top             =   2220
         Width           =   2832
      End
      Begin VB.TextBox LotIDTextBox 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1260
         TabIndex        =   294
         Text            =   "LotID"
         Top             =   1320
         Width           =   2300
      End
      Begin VB.TextBox WaferIDTexBox 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1260
         TabIndex        =   293
         Text            =   "WaferID"
         Top             =   1740
         Width           =   2300
      End
      Begin VB.CommandButton SpecifyEXE 
         Caption         =   "Specify WGFMU stress program"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   292
         Top             =   2940
         Width           =   3975
      End
      Begin VB.OptionButton DUTStressOptionIndividual 
         Caption         =   "Individual devices to stress (max. 5) = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Left            =   120
         MaskColor       =   &H00000000&
         TabIndex        =   291
         Top             =   324
         Value           =   -1  'True
         Width           =   4275
      End
      Begin VB.CommandButton RunButton 
         BackColor       =   &H00000000&
         Caption         =   "Run"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   672
         Left            =   5640
         MaskColor       =   &H00000000&
         TabIndex        =   289
         Top             =   2880
         Width           =   2170
      End
      Begin VB.CommandButton SaveInputAsButton 
         Caption         =   "Save Input As"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   672
         Left            =   5640
         TabIndex        =   288
         Top             =   1200
         Width           =   2170
      End
      Begin VB.CommandButton RetrieveInputButton 
         Caption         =   "Retrieve Input From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   672
         Left            =   5640
         TabIndex        =   287
         Top             =   2016
         Width           =   2170
      End
      Begin VB.CommandButton ConfBTIButton 
         BackColor       =   &H0000C000&
         Caption         =   "Config. DUTs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   5640
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   286
         Top             =   360
         Width           =   2172
      End
      Begin MSComDlg.CommonDialog CommonDialogSaveInput 
         Left            =   11460
         Top             =   300
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialogCheckExstingFile 
         Left            =   12000
         Top             =   300
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialogOpenFile 
         Left            =   9120
         Top             =   300
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label XYLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "X / Y ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   3888
         TabIndex        =   704
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label ChipIPLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Chip ID = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   3696
         TabIndex        =   701
         Top             =   1392
         Width           =   1020
      End
      Begin VB.Label LotID 
         Alignment       =   1  'Right Justify
         Caption         =   "Lot ID = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   120
         TabIndex        =   301
         Top             =   1380
         Width           =   1152
      End
      Begin VB.Label WaferID 
         Alignment       =   1  'Right Justify
         Caption         =   "Wafer ID = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   120
         TabIndex        =   300
         Top             =   1800
         Width           =   1152
      End
      Begin VB.Label DataPathLabel 
         Caption         =   ": path to save data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   120
         TabIndex        =   299
         Top             =   2580
         Width           =   5892
         WordWrap        =   -1  'True
      End
      Begin VB.Label WhereaboutEXE 
         Caption         =   ": choose a program to run stress (e.g. WGFMU_BTI.exe)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   60
         TabIndex        =   298
         Top             =   3300
         Width           =   5115
         WordWrap        =   -1  'True
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6072
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13152
      _ExtentX        =   23204
      _ExtentY        =   10716
      _Version        =   393216
      Tabs            =   10
      TabsPerRow      =   5
      TabHeight       =   420
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "DUT #1"
      TabPicture(0)   =   "WGFMUVBInput.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Tab1MainFrame"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "DUT #2"
      TabPicture(1)   =   "WGFMUVBInput.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Tab2MainFrame"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "DUT #3"
      TabPicture(2)   =   "WGFMUVBInput.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Tab3MainFrame"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "DUT #4"
      TabPicture(3)   =   "WGFMUVBInput.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Tab4MainFrame"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "DUT #5"
      TabPicture(4)   =   "WGFMUVBInput.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Tab5MainFrame"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "TempTab1"
      TabPicture(5)   =   "WGFMUVBInput.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame1"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "TempTab2"
      TabPicture(6)   =   "WGFMUVBInput.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame2"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "TempTab3"
      TabPicture(7)   =   "WGFMUVBInput.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame4"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "TempTab4"
      TabPicture(8)   =   "WGFMUVBInput.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame5"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "TempTab5"
      TabPicture(9)   =   "WGFMUVBInput.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame6"
      Tab(9).ControlCount=   1
      Begin VB.Frame Tab2MainFrame 
         Height          =   5232
         Left            =   -74820
         TabIndex        =   309
         Top             =   600
         Width           =   12972
         Begin VB.TextBox CopyInputFromforDUT2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   8460
            TabIndex        =   311
            Text            =   "1"
            Top             =   3480
            Width           =   612
         End
         Begin VB.CommandButton ClearAllDut2 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   6900
            TabIndex        =   310
            Top             =   4140
            Width           =   1272
         End
         Begin VB.CommandButton Copy_Input_for_DUT2 
            Caption         =   "Copy input from DUT #"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   5880
            TabIndex        =   312
            Top             =   3480
            Width           =   2472
         End
         Begin VB.Frame StressOptionMainFrame 
            Caption         =   "Stress Options"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4752
            Index           =   1
            Left            =   240
            TabIndex        =   318
            Top             =   300
            Width           =   3072
            Begin VB.OptionButton StressOptionRadioButton 
               Caption         =   "AC"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   2
               Left            =   300
               TabIndex        =   320
               Top             =   360
               Width           =   672
            End
            Begin VB.OptionButton StressOptionRadioButton 
               Caption         =   "DC"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   3
               Left            =   1200
               TabIndex        =   319
               Top             =   420
               Width           =   672
            End
            Begin VB.Frame ACStressOptionFrame 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3912
               Index           =   1
               Left            =   60
               TabIndex        =   321
               Top             =   720
               Width           =   2652
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   50
                  Left            =   1200
                  TabIndex        =   424
                  Text            =   "Text11"
                  Top             =   3240
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   51
                  Left            =   1200
                  TabIndex        =   423
                  Text            =   "Text12"
                  Top             =   3540
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   40
                  Left            =   1200
                  TabIndex        =   422
                  Text            =   "Text1"
                  Top             =   120
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   41
                  Left            =   1200
                  TabIndex        =   421
                  Text            =   "Text2"
                  Top             =   420
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   42
                  Left            =   1200
                  TabIndex        =   420
                  Text            =   "Text3"
                  Top             =   720
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   43
                  Left            =   1200
                  TabIndex        =   419
                  Text            =   "Text4"
                  Top             =   1080
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   44
                  Left            =   1200
                  TabIndex        =   418
                  Text            =   "Text5"
                  Top             =   1380
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   45
                  Left            =   1200
                  TabIndex        =   417
                  Text            =   "Text6"
                  Top             =   1740
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   46
                  Left            =   1200
                  TabIndex        =   416
                  Text            =   "Text7"
                  Top             =   2040
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   47
                  Left            =   1200
                  TabIndex        =   415
                  Text            =   "Text8"
                  Top             =   2400
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   48
                  Left            =   1200
                  TabIndex        =   414
                  Text            =   "Text9"
                  Top             =   2640
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   49
                  Left            =   1260
                  TabIndex        =   413
                  Text            =   "Text10"
                  Top             =   2940
                  Width           =   1212
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   51
                  Left            =   780
                  TabIndex        =   412
                  Top             =   3300
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   40
                  Left            =   840
                  TabIndex        =   411
                  Top             =   180
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   41
                  Left            =   780
                  TabIndex        =   410
                  Top             =   480
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   42
                  Left            =   780
                  TabIndex        =   409
                  Top             =   660
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   43
                  Left            =   780
                  TabIndex        =   408
                  Top             =   960
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   44
                  Left            =   780
                  TabIndex        =   407
                  Top             =   1320
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   45
                  Left            =   780
                  TabIndex        =   406
                  Top             =   1620
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   46
                  Left            =   780
                  TabIndex        =   405
                  Top             =   1920
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   47
                  Left            =   780
                  TabIndex        =   404
                  Top             =   2220
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   48
                  Left            =   780
                  TabIndex        =   403
                  Top             =   2460
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   49
                  Left            =   780
                  TabIndex        =   402
                  Top             =   2760
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   50
                  Left            =   780
                  TabIndex        =   401
                  Top             =   3000
                  Width           =   250
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm11"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   50
                  Left            =   120
                  TabIndex        =   390
                  Top             =   3360
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm12"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   51
                  Left            =   60
                  TabIndex        =   389
                  Top             =   3600
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm1"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   40
                  Left            =   180
                  TabIndex        =   400
                  Top             =   180
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm2"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   41
                  Left            =   120
                  TabIndex        =   399
                  Top             =   540
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm3"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   42
                  Left            =   120
                  TabIndex        =   398
                  Top             =   900
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm4"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   43
                  Left            =   120
                  TabIndex        =   397
                  Top             =   1260
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm5"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   44
                  Left            =   120
                  TabIndex        =   396
                  Top             =   1620
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm6"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   45
                  Left            =   120
                  TabIndex        =   395
                  Top             =   1980
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm7"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   46
                  Left            =   120
                  TabIndex        =   394
                  Top             =   2340
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm8"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   47
                  Left            =   120
                  TabIndex        =   393
                  Top             =   2700
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm9"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   48
                  Left            =   120
                  TabIndex        =   392
                  Top             =   2940
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm10"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   49
                  Left            =   60
                  TabIndex        =   391
                  Top             =   3120
                  Width           =   1572
               End
            End
         End
         Begin VB.Frame MOSMainFrame 
            Caption         =   "Measurement Option"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4632
            Index           =   1
            Left            =   3420
            TabIndex        =   313
            Top             =   420
            Width           =   5832
            Begin VB.OptionButton InitIVSweepOption 
               Caption         =   "Initial IVSweep Only"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   312
               Index           =   1
               Left            =   3168
               TabIndex        =   707
               Top             =   192
               Width           =   2076
            End
            Begin VB.OptionButton MeasOptionRadioOption 
               Caption         =   "IVSweep"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   3
               Left            =   1800
               TabIndex        =   323
               Top             =   240
               Width           =   1572
            End
            Begin VB.OptionButton MeasOptionRadioOption 
               Caption         =   "Spot"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   2
               Left            =   180
               TabIndex        =   322
               Top             =   240
               Width           =   1572
            End
            Begin VB.Frame MeasTypeFrame 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3732
               Index           =   1
               Left            =   240
               TabIndex        =   317
               Top             =   600
               Width           =   2112
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   52
                  Left            =   -60
                  TabIndex        =   460
                  Top             =   420
                  Width           =   250
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   52
                  Left            =   1200
                  TabIndex        =   447
                  Text            =   "Text13"
                  Top             =   360
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   53
                  Left            =   1200
                  TabIndex        =   446
                  Text            =   "Text14"
                  Top             =   600
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   54
                  Left            =   1200
                  TabIndex        =   445
                  Text            =   "Text15"
                  Top             =   840
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   55
                  Left            =   1200
                  TabIndex        =   444
                  Text            =   "Text16"
                  Top             =   1080
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   56
                  Left            =   1200
                  TabIndex        =   443
                  Text            =   "Text17"
                  Top             =   1320
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   57
                  Left            =   1200
                  TabIndex        =   442
                  Text            =   "Text18"
                  Top             =   1560
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   58
                  Left            =   1200
                  TabIndex        =   441
                  Text            =   "Text19"
                  Top             =   1860
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   59
                  Left            =   1200
                  TabIndex        =   440
                  Text            =   "Text20"
                  Top             =   2100
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   60
                  Left            =   1200
                  TabIndex        =   439
                  Text            =   "Text21"
                  Top             =   2340
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   61
                  Left            =   1200
                  TabIndex        =   438
                  Text            =   "Text22"
                  Top             =   2580
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   62
                  Left            =   1200
                  TabIndex        =   437
                  Text            =   "Text23"
                  Top             =   2820
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   63
                  Left            =   1260
                  TabIndex        =   436
                  Text            =   "Text24"
                  Top             =   3060
                  Width           =   1212
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   53
                  Left            =   -60
                  TabIndex        =   435
                  Top             =   600
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   54
                  Left            =   -60
                  TabIndex        =   434
                  Top             =   780
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   55
                  Left            =   -60
                  TabIndex        =   433
                  Top             =   960
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   56
                  Left            =   -60
                  TabIndex        =   432
                  Top             =   1140
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   57
                  Left            =   -60
                  TabIndex        =   431
                  Top             =   1320
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   58
                  Left            =   -60
                  TabIndex        =   430
                  Top             =   1560
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   59
                  Left            =   -60
                  TabIndex        =   429
                  Top             =   1740
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   60
                  Left            =   -60
                  TabIndex        =   428
                  Top             =   1920
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   61
                  Left            =   -60
                  TabIndex        =   427
                  Top             =   2100
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   62
                  Left            =   -60
                  TabIndex        =   426
                  Top             =   2280
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   63
                  Left            =   -60
                  TabIndex        =   425
                  Top             =   2460
                  Width           =   250
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm13"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   52
                  Left            =   540
                  TabIndex        =   459
                  Top             =   420
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm14"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   53
                  Left            =   540
                  TabIndex        =   458
                  Top             =   660
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm15"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   54
                  Left            =   480
                  TabIndex        =   457
                  Top             =   900
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm16"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   55
                  Left            =   480
                  TabIndex        =   456
                  Top             =   1140
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm17"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   56
                  Left            =   480
                  TabIndex        =   455
                  Top             =   1380
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm18"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   57
                  Left            =   420
                  TabIndex        =   454
                  Top             =   1560
                  Width           =   1452
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm19"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   58
                  Left            =   360
                  TabIndex        =   453
                  Top             =   1800
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm20"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   59
                  Left            =   420
                  TabIndex        =   452
                  Top             =   2040
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm21"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   60
                  Left            =   420
                  TabIndex        =   451
                  Top             =   2280
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm22"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   61
                  Left            =   420
                  TabIndex        =   450
                  Top             =   2520
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm23"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   62
                  Left            =   360
                  TabIndex        =   449
                  Top             =   2760
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm24"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   63
                  Left            =   420
                  TabIndex        =   448
                  Top             =   2880
                  Width           =   1572
               End
            End
            Begin VB.Frame SetMeasTimeIntervalFrame 
               Caption         =   "Meas. Time Interval"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2052
               Index           =   1
               Left            =   2580
               TabIndex        =   314
               Top             =   780
               Width           =   2832
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   64
                  Left            =   1200
                  TabIndex        =   464
                  Text            =   "Text25"
                  Top             =   780
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   65
                  Left            =   1200
                  TabIndex        =   463
                  Text            =   "Text26"
                  Top             =   1080
                  Width           =   1212
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   64
                  Left            =   300
                  TabIndex        =   462
                  Top             =   780
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   65
                  Left            =   300
                  TabIndex        =   461
                  Top             =   1020
                  Width           =   250
               End
               Begin VB.OptionButton MeasScaleRaioButton 
                  Caption         =   "Log"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   432
                  Index           =   2
                  Left            =   180
                  TabIndex        =   316
                  Top             =   300
                  Width           =   972
               End
               Begin VB.OptionButton MeasScaleRaioButton 
                  Caption         =   "Linear"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   432
                  Index           =   3
                  Left            =   1500
                  TabIndex        =   315
                  Top             =   300
                  Width           =   1272
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm25"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   64
                  Left            =   600
                  TabIndex        =   466
                  Top             =   780
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm26"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   65
                  Left            =   600
                  TabIndex        =   465
                  Top             =   1020
                  Width           =   1572
               End
            End
         End
      End
      Begin VB.Frame Tab5MainFrame 
         Height          =   5232
         Left            =   -74880
         TabIndex        =   248
         Top             =   600
         Width           =   13152
         Begin VB.CommandButton ClearAllDut5 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   6540
            TabIndex        =   260
            Top             =   4020
            Width           =   1272
         End
         Begin VB.CommandButton Copy_Input_for_DUT5 
            Caption         =   "Copy input from DUT #"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   5760
            TabIndex        =   259
            Top             =   3240
            Width           =   2472
         End
         Begin VB.TextBox CopyInputFromforDUT5 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   8400
            TabIndex        =   258
            Text            =   "1"
            Top             =   3240
            Width           =   612
         End
         Begin VB.Frame StressOptionMainFrame 
            Caption         =   "Stress Options"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4872
            Index           =   4
            Left            =   360
            TabIndex        =   254
            Top             =   240
            Width           =   3072
            Begin VB.OptionButton StressOptionRadioButton 
               Caption         =   "DC"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   9
               Left            =   1620
               TabIndex        =   256
               Top             =   420
               Width           =   672
            End
            Begin VB.OptionButton StressOptionRadioButton 
               Caption         =   "AC"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   8
               Left            =   180
               TabIndex        =   255
               Top             =   420
               Width           =   672
            End
            Begin VB.Frame ACStressOptionFrame 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3912
               Index           =   4
               Left            =   60
               TabIndex        =   257
               Top             =   840
               Width           =   2892
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   160
                  Left            =   420
                  TabIndex        =   646
                  Top             =   960
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   161
                  Left            =   420
                  TabIndex        =   645
                  Top             =   1140
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   162
                  Left            =   420
                  TabIndex        =   644
                  Top             =   1320
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   163
                  Left            =   420
                  TabIndex        =   643
                  Top             =   1500
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   164
                  Left            =   420
                  TabIndex        =   642
                  Top             =   1680
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   165
                  Left            =   420
                  TabIndex        =   641
                  Top             =   1860
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   166
                  Left            =   420
                  TabIndex        =   640
                  Top             =   2040
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   167
                  Left            =   420
                  TabIndex        =   639
                  Top             =   2220
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   168
                  Left            =   420
                  TabIndex        =   638
                  Top             =   2400
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   169
                  Left            =   420
                  TabIndex        =   637
                  Top             =   2580
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   170
                  Left            =   420
                  TabIndex        =   636
                  Top             =   2760
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   171
                  Left            =   420
                  TabIndex        =   635
                  Top             =   2940
                  Width           =   250
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   160
                  Left            =   1380
                  TabIndex        =   634
                  Text            =   "Text1"
                  Top             =   960
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   161
                  Left            =   1380
                  TabIndex        =   633
                  Text            =   "Text2"
                  Top             =   1140
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   162
                  Left            =   1380
                  TabIndex        =   632
                  Text            =   "Text3"
                  Top             =   1320
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   163
                  Left            =   1380
                  TabIndex        =   631
                  Text            =   "Text4"
                  Top             =   1560
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   164
                  Left            =   1380
                  TabIndex        =   630
                  Text            =   "Text5"
                  Top             =   1740
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   165
                  Left            =   1380
                  TabIndex        =   629
                  Text            =   "Text6"
                  Top             =   1920
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   166
                  Left            =   1380
                  TabIndex        =   628
                  Text            =   "Text7"
                  Top             =   2100
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   167
                  Left            =   1380
                  TabIndex        =   627
                  Text            =   "Text8"
                  Top             =   2280
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   168
                  Left            =   1380
                  TabIndex        =   626
                  Text            =   "Text9"
                  Top             =   2460
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   169
                  Left            =   1380
                  TabIndex        =   625
                  Text            =   "Text10"
                  Top             =   2640
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   170
                  Left            =   1380
                  TabIndex        =   624
                  Text            =   "Text11"
                  Top             =   2820
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   171
                  Left            =   1380
                  TabIndex        =   623
                  Text            =   "Text12"
                  Top             =   3000
                  Width           =   1212
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm1"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   160
                  Left            =   840
                  TabIndex        =   658
                  Top             =   960
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm2"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   161
                  Left            =   840
                  TabIndex        =   657
                  Top             =   1140
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm3"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   162
                  Left            =   840
                  TabIndex        =   656
                  Top             =   1320
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm4"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   163
                  Left            =   840
                  TabIndex        =   655
                  Top             =   1500
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm5"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   164
                  Left            =   840
                  TabIndex        =   654
                  Top             =   1680
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm6"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   165
                  Left            =   840
                  TabIndex        =   653
                  Top             =   1860
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm7"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   166
                  Left            =   840
                  TabIndex        =   652
                  Top             =   2040
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm8"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   167
                  Left            =   840
                  TabIndex        =   651
                  Top             =   2220
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm9"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   168
                  Left            =   840
                  TabIndex        =   650
                  Top             =   2400
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm10"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   169
                  Left            =   840
                  TabIndex        =   649
                  Top             =   2580
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm11"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   170
                  Left            =   840
                  TabIndex        =   648
                  Top             =   2760
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm12"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   171
                  Left            =   840
                  TabIndex        =   647
                  Top             =   2940
                  Width           =   1572
               End
            End
         End
         Begin VB.Frame MOSMainFrame 
            Caption         =   "Measurement Option"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4632
            Index           =   4
            Left            =   3480
            TabIndex        =   249
            Top             =   360
            Width           =   5652
            Begin VB.OptionButton InitIVSweepOption 
               Caption         =   "Initial IVSweep Only"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   312
               Index           =   4
               Left            =   3312
               TabIndex        =   710
               Top             =   240
               Width           =   2076
            End
            Begin VB.OptionButton MeasOptionRadioOption 
               Caption         =   "Spot"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   8
               Left            =   180
               TabIndex        =   329
               Top             =   240
               Width           =   972
            End
            Begin VB.OptionButton MeasOptionRadioOption 
               Caption         =   "IVSweep"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   9
               Left            =   1560
               TabIndex        =   328
               Top             =   240
               Width           =   1572
            End
            Begin VB.Frame MeasTypeFrame 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3852
               Index           =   4
               Left            =   60
               TabIndex        =   253
               Top             =   600
               Width           =   2712
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   172
                  Left            =   0
                  TabIndex        =   682
                  Top             =   1020
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   173
                  Left            =   0
                  TabIndex        =   681
                  Top             =   1200
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   174
                  Left            =   0
                  TabIndex        =   680
                  Top             =   1380
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   175
                  Left            =   0
                  TabIndex        =   679
                  Top             =   1560
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   176
                  Left            =   0
                  TabIndex        =   678
                  Top             =   1740
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   177
                  Left            =   0
                  TabIndex        =   677
                  Top             =   1920
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   178
                  Left            =   0
                  TabIndex        =   676
                  Top             =   2100
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   179
                  Left            =   0
                  TabIndex        =   675
                  Top             =   2280
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   180
                  Left            =   0
                  TabIndex        =   674
                  Top             =   2460
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   181
                  Left            =   0
                  TabIndex        =   673
                  Top             =   2640
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   182
                  Left            =   0
                  TabIndex        =   672
                  Top             =   2820
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   183
                  Left            =   0
                  TabIndex        =   671
                  Top             =   3000
                  Width           =   250
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   172
                  Left            =   1200
                  TabIndex        =   670
                  Text            =   "Text13"
                  Top             =   1020
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   173
                  Left            =   1200
                  TabIndex        =   669
                  Text            =   "Text14"
                  Top             =   1200
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   174
                  Left            =   1200
                  TabIndex        =   668
                  Text            =   "Text15"
                  Top             =   1440
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   175
                  Left            =   1200
                  TabIndex        =   667
                  Text            =   "Text16"
                  Top             =   1620
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   176
                  Left            =   1200
                  TabIndex        =   666
                  Text            =   "Text17"
                  Top             =   1800
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   177
                  Left            =   1200
                  TabIndex        =   665
                  Text            =   "Text18"
                  Top             =   1980
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   178
                  Left            =   1200
                  TabIndex        =   664
                  Text            =   "Text19"
                  Top             =   2160
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   179
                  Left            =   1200
                  TabIndex        =   663
                  Text            =   "Text20"
                  Top             =   2340
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   180
                  Left            =   1200
                  TabIndex        =   662
                  Text            =   "Text21"
                  Top             =   2520
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   181
                  Left            =   1200
                  TabIndex        =   661
                  Text            =   "Text22"
                  Top             =   2700
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   182
                  Left            =   1200
                  TabIndex        =   660
                  Text            =   "Text23"
                  Top             =   2880
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   183
                  Left            =   1200
                  TabIndex        =   659
                  Text            =   "Text24"
                  Top             =   3060
                  Width           =   1212
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm13"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   172
                  Left            =   420
                  TabIndex        =   694
                  Top             =   1020
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm14"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   173
                  Left            =   420
                  TabIndex        =   693
                  Top             =   1200
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm15"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   174
                  Left            =   420
                  TabIndex        =   692
                  Top             =   1380
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm16"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   175
                  Left            =   420
                  TabIndex        =   691
                  Top             =   1560
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm17"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   176
                  Left            =   420
                  TabIndex        =   690
                  Top             =   1740
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm18"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   177
                  Left            =   420
                  TabIndex        =   689
                  Top             =   1920
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm19"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   178
                  Left            =   420
                  TabIndex        =   688
                  Top             =   2100
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm20"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   179
                  Left            =   420
                  TabIndex        =   687
                  Top             =   2280
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm21"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   180
                  Left            =   420
                  TabIndex        =   686
                  Top             =   2460
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm22"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   181
                  Left            =   420
                  TabIndex        =   685
                  Top             =   2640
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm23"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   182
                  Left            =   420
                  TabIndex        =   684
                  Top             =   2820
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm24"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   183
                  Left            =   420
                  TabIndex        =   683
                  Top             =   3000
                  Width           =   1572
               End
            End
            Begin VB.Frame SetMeasTimeIntervalFrame 
               Caption         =   "Meas. Time Interval"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1752
               Index           =   4
               Left            =   2520
               TabIndex        =   250
               Top             =   900
               Width           =   2832
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   184
                  Left            =   360
                  TabIndex        =   698
                  Top             =   840
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   185
                  Left            =   360
                  TabIndex        =   697
                  Top             =   1140
                  Width           =   250
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   184
                  Left            =   1260
                  TabIndex        =   696
                  Text            =   "Text25"
                  Top             =   720
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   185
                  Left            =   1320
                  TabIndex        =   695
                  Text            =   "Text26"
                  Top             =   1080
                  Width           =   1212
               End
               Begin VB.OptionButton MeasScaleRaioButton 
                  Caption         =   "Linear"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   432
                  Index           =   9
                  Left            =   1440
                  TabIndex        =   252
                  Top             =   360
                  Width           =   1272
               End
               Begin VB.OptionButton MeasScaleRaioButton 
                  Caption         =   "Log"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   432
                  Index           =   8
                  Left            =   180
                  TabIndex        =   251
                  Top             =   300
                  Width           =   1272
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm25"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   184
                  Left            =   660
                  TabIndex        =   700
                  Top             =   840
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm26"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   185
                  Left            =   660
                  TabIndex        =   699
                  Top             =   1140
                  Width           =   1572
               End
            End
         End
      End
      Begin VB.Frame Tab4MainFrame 
         Height          =   5172
         Left            =   -74700
         TabIndex        =   235
         Top             =   660
         Width           =   12972
         Begin VB.CommandButton ClearAllDut4 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   6840
            TabIndex        =   247
            Top             =   3720
            Width           =   1272
         End
         Begin VB.CommandButton Copy_Input_for_DUT4 
            Caption         =   "Copy input from DUT #"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   5700
            TabIndex        =   246
            Top             =   2700
            Width           =   2472
         End
         Begin VB.TextBox CopyInputFromforDUT4 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   8280
            TabIndex        =   245
            Text            =   "1"
            Top             =   2700
            Width           =   612
         End
         Begin VB.Frame MOSMainFrame 
            Caption         =   "Measurement Option"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4812
            Index           =   3
            Left            =   3240
            TabIndex        =   240
            Top             =   180
            Width           =   5652
            Begin VB.OptionButton InitIVSweepOption 
               Caption         =   "Initial IVSweep Only"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   312
               Index           =   3
               Left            =   3024
               TabIndex        =   709
               Top             =   240
               Width           =   2076
            End
            Begin VB.OptionButton MeasOptionRadioOption 
               Caption         =   "Spot"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   6
               Left            =   240
               TabIndex        =   327
               Top             =   360
               Width           =   972
            End
            Begin VB.OptionButton MeasOptionRadioOption 
               Caption         =   "IVSweep"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   7
               Left            =   1260
               TabIndex        =   326
               Top             =   360
               Width           =   1572
            End
            Begin VB.Frame SetMeasTimeIntervalFrame 
               Caption         =   "Meas. Time Interval"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1692
               Index           =   3
               Left            =   2580
               TabIndex        =   242
               Top             =   660
               Width           =   2832
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   144
                  Left            =   1140
                  TabIndex        =   620
                  Text            =   "Text25"
                  Top             =   840
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   145
                  Left            =   1140
                  TabIndex        =   619
                  Text            =   "Text26"
                  Top             =   1200
                  Width           =   1212
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   144
                  Left            =   180
                  TabIndex        =   618
                  Top             =   900
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   145
                  Left            =   180
                  TabIndex        =   617
                  Top             =   1140
                  Width           =   250
               End
               Begin VB.OptionButton MeasScaleRaioButton 
                  Caption         =   "Linear"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   432
                  Index           =   7
                  Left            =   1440
                  TabIndex        =   244
                  Top             =   420
                  Width           =   1272
               End
               Begin VB.OptionButton MeasScaleRaioButton 
                  Caption         =   "Log"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   432
                  Index           =   6
                  Left            =   240
                  TabIndex        =   243
                  Top             =   360
                  Width           =   1272
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm25"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   144
                  Left            =   600
                  TabIndex        =   622
                  Top             =   900
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm26"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   145
                  Left            =   600
                  TabIndex        =   621
                  Top             =   1140
                  Width           =   1572
               End
            End
            Begin VB.Frame MeasTypeFrame 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4152
               Index           =   3
               Left            =   60
               TabIndex        =   241
               Top             =   600
               Width           =   2472
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   132
                  Left            =   1200
                  TabIndex        =   604
                  Text            =   "Text13"
                  Top             =   540
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   133
                  Left            =   1200
                  TabIndex        =   603
                  Text            =   "Text14"
                  Top             =   780
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   134
                  Left            =   1200
                  TabIndex        =   602
                  Text            =   "Text15"
                  Top             =   1020
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   135
                  Left            =   1200
                  TabIndex        =   601
                  Text            =   "Text16"
                  Top             =   1260
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   136
                  Left            =   1200
                  TabIndex        =   600
                  Text            =   "Text17"
                  Top             =   1500
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   137
                  Left            =   1200
                  TabIndex        =   599
                  Text            =   "Text18"
                  Top             =   1740
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   138
                  Left            =   1200
                  TabIndex        =   598
                  Text            =   "Text19"
                  Top             =   1920
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   139
                  Left            =   1200
                  TabIndex        =   597
                  Text            =   "Text20"
                  Top             =   2160
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   140
                  Left            =   1200
                  TabIndex        =   596
                  Text            =   "Text21"
                  Top             =   2400
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   141
                  Left            =   1200
                  TabIndex        =   595
                  Text            =   "Text22"
                  Top             =   2580
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   142
                  Left            =   1200
                  TabIndex        =   594
                  Text            =   "Text23"
                  Top             =   2820
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   143
                  Left            =   1200
                  TabIndex        =   593
                  Text            =   "Text24"
                  Top             =   3060
                  Width           =   1212
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   132
                  Left            =   180
                  TabIndex        =   592
                  Top             =   540
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   133
                  Left            =   180
                  TabIndex        =   591
                  Top             =   720
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   134
                  Left            =   180
                  TabIndex        =   590
                  Top             =   900
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   135
                  Left            =   180
                  TabIndex        =   589
                  Top             =   1080
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   136
                  Left            =   180
                  TabIndex        =   588
                  Top             =   1260
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   137
                  Left            =   180
                  TabIndex        =   587
                  Top             =   1440
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   138
                  Left            =   180
                  TabIndex        =   586
                  Top             =   1620
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   139
                  Left            =   180
                  TabIndex        =   585
                  Top             =   1800
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   140
                  Left            =   180
                  TabIndex        =   584
                  Top             =   1980
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   141
                  Left            =   180
                  TabIndex        =   583
                  Top             =   2160
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   142
                  Left            =   180
                  TabIndex        =   582
                  Top             =   2340
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   143
                  Left            =   180
                  TabIndex        =   581
                  Top             =   2520
                  Width           =   250
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm13"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   132
                  Left            =   600
                  TabIndex        =   616
                  Top             =   540
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm14"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   133
                  Left            =   600
                  TabIndex        =   615
                  Top             =   720
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm15"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   134
                  Left            =   600
                  TabIndex        =   614
                  Top             =   900
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm16"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   135
                  Left            =   600
                  TabIndex        =   613
                  Top             =   1080
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm17"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   136
                  Left            =   600
                  TabIndex        =   612
                  Top             =   1260
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm18"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   137
                  Left            =   600
                  TabIndex        =   611
                  Top             =   1440
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm19"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   138
                  Left            =   600
                  TabIndex        =   610
                  Top             =   1620
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm20"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   139
                  Left            =   600
                  TabIndex        =   609
                  Top             =   1800
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm21"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   140
                  Left            =   600
                  TabIndex        =   608
                  Top             =   1980
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm22"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   141
                  Left            =   600
                  TabIndex        =   607
                  Top             =   2160
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm23"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   142
                  Left            =   600
                  TabIndex        =   606
                  Top             =   2340
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm24"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   143
                  Left            =   600
                  TabIndex        =   605
                  Top             =   2520
                  Width           =   1572
               End
            End
         End
         Begin VB.Frame StressOptionMainFrame 
            Caption         =   "Stress Options"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4752
            Index           =   3
            Left            =   120
            TabIndex        =   236
            Top             =   300
            Width           =   3072
            Begin VB.OptionButton StressOptionRadioButton 
               Caption         =   "DC"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   7
               Left            =   1500
               TabIndex        =   239
               Top             =   420
               Width           =   672
            End
            Begin VB.OptionButton StressOptionRadioButton 
               Caption         =   "AC"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   6
               Left            =   240
               TabIndex        =   238
               Top             =   420
               Width           =   672
            End
            Begin VB.Frame ACStressOptionFrame 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3912
               Index           =   3
               Left            =   120
               TabIndex        =   237
               Top             =   780
               Width           =   2892
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   120
                  Left            =   120
                  TabIndex        =   568
                  Top             =   360
                  Width           =   250
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   120
                  Left            =   1260
                  TabIndex        =   567
                  Text            =   "Text1"
                  Top             =   360
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   121
                  Left            =   1260
                  TabIndex        =   566
                  Text            =   "Text2"
                  Top             =   600
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   122
                  Left            =   1260
                  TabIndex        =   565
                  Text            =   "Text3"
                  Top             =   840
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   123
                  Left            =   1260
                  TabIndex        =   564
                  Text            =   "Text4"
                  Top             =   1020
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   124
                  Left            =   1260
                  TabIndex        =   563
                  Text            =   "Text5"
                  Top             =   1260
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   125
                  Left            =   1260
                  TabIndex        =   562
                  Text            =   "Text6"
                  Top             =   1500
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   126
                  Left            =   1260
                  TabIndex        =   561
                  Text            =   "Text7"
                  Top             =   1740
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   127
                  Left            =   1260
                  TabIndex        =   560
                  Text            =   "Text8"
                  Top             =   1980
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   128
                  Left            =   1260
                  TabIndex        =   559
                  Text            =   "Text9"
                  Top             =   2220
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   129
                  Left            =   1260
                  TabIndex        =   558
                  Text            =   "Text10"
                  Top             =   2460
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   130
                  Left            =   1260
                  TabIndex        =   557
                  Text            =   "Text11"
                  Top             =   2700
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   131
                  Left            =   1260
                  TabIndex        =   556
                  Text            =   "Text12"
                  Top             =   2940
                  Width           =   1212
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   121
                  Left            =   120
                  TabIndex        =   555
                  Top             =   540
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   122
                  Left            =   120
                  TabIndex        =   554
                  Top             =   780
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   123
                  Left            =   120
                  TabIndex        =   553
                  Top             =   960
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   124
                  Left            =   120
                  TabIndex        =   552
                  Top             =   1200
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   125
                  Left            =   120
                  TabIndex        =   551
                  Top             =   1320
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   126
                  Left            =   120
                  TabIndex        =   550
                  Top             =   1500
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   127
                  Left            =   120
                  TabIndex        =   549
                  Top             =   1680
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   128
                  Left            =   120
                  TabIndex        =   548
                  Top             =   1860
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   129
                  Left            =   120
                  TabIndex        =   547
                  Top             =   2040
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   130
                  Left            =   120
                  TabIndex        =   546
                  Top             =   2220
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   131
                  Left            =   120
                  TabIndex        =   545
                  Top             =   2340
                  Width           =   250
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm1"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   120
                  Left            =   420
                  TabIndex        =   580
                  Top             =   360
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm2"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   121
                  Left            =   420
                  TabIndex        =   579
                  Top             =   540
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm3"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   122
                  Left            =   420
                  TabIndex        =   578
                  Top             =   720
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm4"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   123
                  Left            =   420
                  TabIndex        =   577
                  Top             =   900
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm5"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   124
                  Left            =   420
                  TabIndex        =   576
                  Top             =   1080
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm6"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   125
                  Left            =   420
                  TabIndex        =   575
                  Top             =   1260
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm7"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   126
                  Left            =   420
                  TabIndex        =   574
                  Top             =   1440
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm8"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   127
                  Left            =   420
                  TabIndex        =   573
                  Top             =   1620
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm9"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   128
                  Left            =   420
                  TabIndex        =   572
                  Top             =   1800
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm10"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   129
                  Left            =   420
                  TabIndex        =   571
                  Top             =   1980
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm11"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   130
                  Left            =   420
                  TabIndex        =   570
                  Top             =   2160
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm12"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   131
                  Left            =   420
                  TabIndex        =   569
                  Top             =   2340
                  Width           =   1572
               End
            End
         End
      End
      Begin VB.Frame Tab3MainFrame 
         Height          =   5112
         Left            =   -74760
         TabIndex        =   222
         Top             =   660
         Width           =   13032
         Begin VB.TextBox CopyInputFromforDUT3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   8280
            TabIndex        =   232
            Text            =   "1"
            Top             =   2700
            Width           =   612
         End
         Begin VB.CommandButton ClearAllDut3 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   6840
            TabIndex        =   234
            Top             =   3600
            Width           =   1272
         End
         Begin VB.CommandButton Copy_Input_for_DUT3 
            Caption         =   "Copy input from DUT #"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   5820
            TabIndex        =   233
            Top             =   2700
            Width           =   2472
         End
         Begin VB.Frame StressOptionMainFrame 
            Caption         =   "Stress Options"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4512
            Index           =   2
            Left            =   120
            TabIndex        =   228
            Top             =   300
            Width           =   3072
            Begin VB.OptionButton StressOptionRadioButton 
               Caption         =   "DC"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   5
               Left            =   1620
               TabIndex        =   230
               Top             =   420
               Width           =   672
            End
            Begin VB.OptionButton StressOptionRadioButton 
               Caption         =   "AC"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   4
               Left            =   180
               TabIndex        =   229
               Top             =   420
               Width           =   672
            End
            Begin VB.Frame ACStressOptionFrame 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3552
               Index           =   2
               Left            =   240
               TabIndex        =   231
               Top             =   780
               Width           =   2652
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   91
                  Left            =   1020
                  TabIndex        =   502
                  Text            =   "Text12"
                  Top             =   2580
                  Width           =   1212
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   80
                  Left            =   180
                  TabIndex        =   489
                  Top             =   480
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   81
                  Left            =   180
                  TabIndex        =   488
                  Top             =   720
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   82
                  Left            =   180
                  TabIndex        =   487
                  Top             =   900
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   83
                  Left            =   180
                  TabIndex        =   486
                  Top             =   1140
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   84
                  Left            =   180
                  TabIndex        =   485
                  Top             =   1380
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   85
                  Left            =   180
                  TabIndex        =   484
                  Top             =   1560
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   86
                  Left            =   180
                  TabIndex        =   483
                  Top             =   1740
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   87
                  Left            =   180
                  TabIndex        =   482
                  Top             =   1920
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   88
                  Left            =   180
                  TabIndex        =   481
                  Top             =   2100
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   89
                  Left            =   180
                  TabIndex        =   480
                  Top             =   2280
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   90
                  Left            =   180
                  TabIndex        =   479
                  Top             =   2460
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   91
                  Left            =   240
                  TabIndex        =   478
                  Top             =   2580
                  Width           =   250
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   80
                  Left            =   1020
                  TabIndex        =   477
                  Text            =   "Text1"
                  Top             =   420
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   81
                  Left            =   1020
                  TabIndex        =   476
                  Text            =   "Text2"
                  Top             =   660
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   82
                  Left            =   1020
                  TabIndex        =   475
                  Text            =   "Text3"
                  Top             =   840
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   83
                  Left            =   1020
                  TabIndex        =   474
                  Text            =   "Text4"
                  Top             =   1020
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   84
                  Left            =   1020
                  TabIndex        =   473
                  Text            =   "Text5"
                  Top             =   1200
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   85
                  Left            =   1020
                  TabIndex        =   472
                  Text            =   "Text6"
                  Top             =   1380
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   86
                  Left            =   1020
                  TabIndex        =   471
                  Text            =   "Text7"
                  Top             =   1560
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   87
                  Left            =   1020
                  TabIndex        =   470
                  Text            =   "Text8"
                  Top             =   1740
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   88
                  Left            =   1020
                  TabIndex        =   469
                  Text            =   "Text9"
                  Top             =   1920
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   89
                  Left            =   1020
                  TabIndex        =   468
                  Text            =   "Text10"
                  Top             =   2100
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   90
                  Left            =   1020
                  TabIndex        =   467
                  Text            =   "Text11"
                  Top             =   2280
                  Width           =   1212
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm1"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   80
                  Left            =   540
                  TabIndex        =   501
                  Top             =   480
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm2"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   81
                  Left            =   540
                  TabIndex        =   500
                  Top             =   660
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm3"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   82
                  Left            =   540
                  TabIndex        =   499
                  Top             =   780
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm4"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   83
                  Left            =   540
                  TabIndex        =   498
                  Top             =   960
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm5"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   84
                  Left            =   540
                  TabIndex        =   497
                  Top             =   1140
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm6"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   85
                  Left            =   540
                  TabIndex        =   496
                  Top             =   1320
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm7"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   86
                  Left            =   540
                  TabIndex        =   495
                  Top             =   1500
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm8"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   87
                  Left            =   540
                  TabIndex        =   494
                  Top             =   1680
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm9"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   88
                  Left            =   540
                  TabIndex        =   493
                  Top             =   1860
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm10"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   89
                  Left            =   540
                  TabIndex        =   492
                  Top             =   2040
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm11"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   90
                  Left            =   540
                  TabIndex        =   491
                  Top             =   2220
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm12"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   91
                  Left            =   540
                  TabIndex        =   490
                  Top             =   2460
                  Width           =   1572
               End
            End
         End
         Begin VB.Frame MOSMainFrame 
            Caption         =   "Measurement Option"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4632
            Index           =   2
            Left            =   3300
            TabIndex        =   223
            Top             =   240
            Width           =   5652
            Begin VB.OptionButton InitIVSweepOption 
               Caption         =   "Initial IVSweep Only"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   312
               Index           =   2
               Left            =   2784
               TabIndex        =   708
               Top             =   192
               Width           =   2076
            End
            Begin VB.OptionButton MeasOptionRadioOption 
               Caption         =   "Spot"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   4
               Left            =   180
               TabIndex        =   325
               Top             =   360
               Width           =   972
            End
            Begin VB.OptionButton MeasOptionRadioOption 
               Caption         =   "IVSweep"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   5
               Left            =   1260
               TabIndex        =   324
               Top             =   360
               Width           =   1572
            End
            Begin VB.Frame MeasTypeFrame 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3612
               Index           =   2
               Left            =   120
               TabIndex        =   227
               Top             =   780
               Width           =   2352
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   92
                  Left            =   120
                  TabIndex        =   526
                  Top             =   300
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   93
                  Left            =   120
                  TabIndex        =   525
                  Top             =   480
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   94
                  Left            =   120
                  TabIndex        =   524
                  Top             =   660
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   95
                  Left            =   120
                  TabIndex        =   523
                  Top             =   840
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   96
                  Left            =   120
                  TabIndex        =   522
                  Top             =   1020
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   97
                  Left            =   120
                  TabIndex        =   521
                  Top             =   1200
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   98
                  Left            =   120
                  TabIndex        =   520
                  Top             =   1380
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   99
                  Left            =   120
                  TabIndex        =   519
                  Top             =   1560
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   100
                  Left            =   120
                  TabIndex        =   518
                  Top             =   1680
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   101
                  Left            =   120
                  TabIndex        =   517
                  Top             =   1860
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   102
                  Left            =   120
                  TabIndex        =   516
                  Top             =   2040
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   103
                  Left            =   120
                  TabIndex        =   515
                  Top             =   2220
                  Width           =   250
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   92
                  Left            =   360
                  TabIndex        =   514
                  Text            =   "Text13"
                  Top             =   300
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   93
                  Left            =   360
                  TabIndex        =   513
                  Text            =   "Text14"
                  Top             =   600
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   94
                  Left            =   360
                  TabIndex        =   512
                  Text            =   "Text15"
                  Top             =   840
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   95
                  Left            =   360
                  TabIndex        =   511
                  Text            =   "Text16"
                  Top             =   1080
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   96
                  Left            =   360
                  TabIndex        =   510
                  Text            =   "Text17"
                  Top             =   1320
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   97
                  Left            =   360
                  TabIndex        =   509
                  Text            =   "Text18"
                  Top             =   1560
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   98
                  Left            =   360
                  TabIndex        =   508
                  Text            =   "Text19"
                  Top             =   1800
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   99
                  Left            =   360
                  TabIndex        =   507
                  Text            =   "Text20"
                  Top             =   2040
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   100
                  Left            =   360
                  TabIndex        =   506
                  Text            =   "Text21"
                  Top             =   2280
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   101
                  Left            =   360
                  TabIndex        =   505
                  Text            =   "Text22"
                  Top             =   2520
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   102
                  Left            =   360
                  TabIndex        =   504
                  Text            =   "Text23"
                  Top             =   2820
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   103
                  Left            =   360
                  TabIndex        =   503
                  Text            =   "Text24"
                  Top             =   3120
                  Width           =   1212
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm13"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   92
                  Left            =   1560
                  TabIndex        =   538
                  Top             =   300
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm14"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   93
                  Left            =   1560
                  TabIndex        =   537
                  Top             =   480
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm15"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   94
                  Left            =   1560
                  TabIndex        =   536
                  Top             =   660
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm25"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   95
                  Left            =   1560
                  TabIndex        =   535
                  Top             =   840
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm16"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   96
                  Left            =   1620
                  TabIndex        =   534
                  Top             =   1020
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm17"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   97
                  Left            =   1560
                  TabIndex        =   533
                  Top             =   1200
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm18"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   98
                  Left            =   1620
                  TabIndex        =   532
                  Top             =   1380
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm19"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   99
                  Left            =   1620
                  TabIndex        =   531
                  Top             =   1560
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm20"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   100
                  Left            =   1620
                  TabIndex        =   530
                  Top             =   1740
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm21"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   101
                  Left            =   1620
                  TabIndex        =   529
                  Top             =   1920
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm22"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   102
                  Left            =   1560
                  TabIndex        =   528
                  Top             =   2160
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm23"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   103
                  Left            =   1620
                  TabIndex        =   527
                  Top             =   2400
                  Width           =   1572
               End
            End
            Begin VB.Frame SetMeasTimeIntervalFrame 
               Caption         =   "Meas. Time Interval"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1392
               Index           =   2
               Left            =   2520
               TabIndex        =   224
               Top             =   660
               Width           =   2832
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   104
                  Left            =   180
                  TabIndex        =   542
                  Top             =   720
                  Width           =   250
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
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
                  Index           =   105
                  Left            =   180
                  TabIndex        =   541
                  Top             =   1020
                  Width           =   250
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   104
                  Left            =   480
                  TabIndex        =   540
                  Text            =   "Text25"
                  Top             =   660
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   105
                  Left            =   480
                  TabIndex        =   539
                  Text            =   "Text26"
                  Top             =   960
                  Width           =   1212
               End
               Begin VB.OptionButton MeasScaleRaioButton 
                  Caption         =   "Linear"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   432
                  Index           =   5
                  Left            =   1440
                  TabIndex        =   226
                  Top             =   300
                  Width           =   1272
               End
               Begin VB.OptionButton MeasScaleRaioButton 
                  Caption         =   "Log"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   432
                  Index           =   4
                  Left            =   240
                  TabIndex        =   225
                  Top             =   300
                  Width           =   1272
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm24"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   104
                  Left            =   1680
                  TabIndex        =   544
                  Top             =   720
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm25"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   105
                  Left            =   1740
                  TabIndex        =   543
                  Top             =   960
                  Width           =   1572
               End
            End
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4392
         Left            =   -74580
         TabIndex        =   179
         Top             =   720
         Width           =   12792
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   199
            Left            =   12180
            TabIndex        =   207
            Top             =   4020
            Width           =   250
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   199
            Left            =   1140
            TabIndex        =   206
            Text            =   "Text40"
            Top             =   360
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   198
            Left            =   1140
            TabIndex        =   205
            Text            =   "Text39"
            Top             =   720
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   197
            Left            =   1140
            TabIndex        =   204
            Text            =   "Text38"
            Top             =   1080
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   196
            Left            =   1140
            TabIndex        =   203
            Text            =   "Text37"
            Top             =   1440
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   195
            Left            =   1140
            TabIndex        =   202
            Text            =   "Text36"
            Top             =   1800
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   194
            Left            =   1140
            TabIndex        =   201
            Text            =   "Text35"
            Top             =   2160
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   193
            Left            =   1140
            TabIndex        =   200
            Text            =   "Text34"
            Top             =   2520
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   192
            Left            =   1140
            TabIndex        =   199
            Text            =   "Text33"
            Top             =   2880
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   191
            Left            =   1140
            TabIndex        =   198
            Text            =   "Text32"
            Top             =   3240
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   190
            Left            =   1140
            TabIndex        =   197
            Text            =   "Text31"
            Top             =   3600
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   189
            Left            =   3900
            TabIndex        =   196
            Text            =   "Text30"
            Top             =   360
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   188
            Left            =   3900
            TabIndex        =   195
            Text            =   "Text29"
            Top             =   720
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   187
            Left            =   3900
            TabIndex        =   194
            Text            =   "Text28"
            Top             =   1080
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   186
            Left            =   3900
            TabIndex        =   193
            Text            =   "Text27"
            Top             =   1440
            Width           =   1212
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   198
            Left            =   2400
            TabIndex        =   192
            Top             =   360
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   197
            Left            =   2400
            TabIndex        =   191
            Top             =   600
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   196
            Left            =   2400
            TabIndex        =   190
            Top             =   900
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   195
            Left            =   2400
            TabIndex        =   189
            Top             =   1260
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   194
            Left            =   2400
            TabIndex        =   188
            Top             =   1560
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   193
            Left            =   2400
            TabIndex        =   187
            Top             =   1860
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   192
            Left            =   2400
            TabIndex        =   186
            Top             =   2160
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   191
            Left            =   2400
            TabIndex        =   185
            Top             =   2400
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   190
            Left            =   2400
            TabIndex        =   184
            Top             =   2700
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   189
            Left            =   2400
            TabIndex        =   183
            Top             =   2940
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   188
            Left            =   5100
            TabIndex        =   182
            Top             =   360
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   187
            Left            =   5100
            TabIndex        =   181
            Top             =   840
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   186
            Left            =   5160
            TabIndex        =   180
            Top             =   1200
            Width           =   250
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm40"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   199
            Left            =   60
            TabIndex        =   221
            Top             =   360
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm39"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   198
            Left            =   0
            TabIndex        =   220
            Top             =   720
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm38"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   197
            Left            =   0
            TabIndex        =   219
            Top             =   1080
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm37"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   196
            Left            =   0
            TabIndex        =   218
            Top             =   1440
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm36"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   195
            Left            =   0
            TabIndex        =   217
            Top             =   1800
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm35"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   194
            Left            =   0
            TabIndex        =   216
            Top             =   2160
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm34"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   193
            Left            =   0
            TabIndex        =   215
            Top             =   2520
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm33"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   192
            Left            =   0
            TabIndex        =   214
            Top             =   2880
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm32"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   191
            Left            =   0
            TabIndex        =   213
            Top             =   3240
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm31"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   190
            Left            =   0
            TabIndex        =   212
            Top             =   3600
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm30"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   189
            Left            =   2700
            TabIndex        =   211
            Top             =   3600
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm29"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   188
            Left            =   2700
            TabIndex        =   210
            Top             =   3240
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm28"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   187
            Left            =   2700
            TabIndex        =   209
            Top             =   2880
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm27"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   186
            Left            =   2700
            TabIndex        =   208
            Top             =   2520
            Width           =   1572
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4392
         Left            =   -74580
         TabIndex        =   136
         Top             =   660
         Width           =   12792
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   159
            Left            =   11700
            TabIndex        =   164
            Top             =   3300
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   158
            Left            =   11760
            TabIndex        =   163
            Top             =   2940
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   157
            Left            =   11700
            TabIndex        =   162
            Top             =   2580
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   156
            Left            =   11700
            TabIndex        =   161
            Top             =   2220
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   155
            Left            =   11760
            TabIndex        =   160
            Top             =   1860
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   154
            Left            =   11760
            TabIndex        =   159
            Top             =   1500
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   153
            Left            =   11820
            TabIndex        =   158
            Top             =   1200
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   152
            Left            =   11760
            TabIndex        =   157
            Top             =   780
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   151
            Left            =   11760
            TabIndex        =   156
            Top             =   420
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   150
            Left            =   8760
            TabIndex        =   155
            Top             =   3660
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   149
            Left            =   8760
            TabIndex        =   154
            Top             =   3300
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   148
            Left            =   8760
            TabIndex        =   153
            Top             =   2940
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   147
            Left            =   8760
            TabIndex        =   152
            Top             =   2580
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   146
            Left            =   8760
            TabIndex        =   151
            Top             =   2220
            Width           =   250
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   159
            Left            =   10440
            TabIndex        =   150
            Text            =   "Text40"
            Top             =   3600
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   158
            Left            =   10440
            TabIndex        =   149
            Text            =   "Text39"
            Top             =   3240
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   157
            Left            =   10440
            TabIndex        =   148
            Text            =   "Text38"
            Top             =   2880
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   156
            Left            =   10440
            TabIndex        =   147
            Text            =   "Text37"
            Top             =   2520
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   155
            Left            =   10440
            TabIndex        =   146
            Text            =   "Text36"
            Top             =   2160
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   154
            Left            =   10440
            TabIndex        =   145
            Text            =   "Text35"
            Top             =   1800
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   153
            Left            =   10440
            TabIndex        =   144
            Text            =   "Text34"
            Top             =   1440
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   152
            Left            =   10440
            TabIndex        =   143
            Text            =   "Text33"
            Top             =   1080
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   151
            Left            =   10440
            TabIndex        =   142
            Text            =   "Text32"
            Top             =   720
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   150
            Left            =   10440
            TabIndex        =   141
            Text            =   "Text31"
            Top             =   360
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   149
            Left            =   7500
            TabIndex        =   140
            Text            =   "Text30"
            Top             =   3600
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   148
            Left            =   7500
            TabIndex        =   139
            Text            =   "Text29"
            Top             =   3240
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   147
            Left            =   7500
            TabIndex        =   138
            Text            =   "Text28"
            Top             =   2880
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   146
            Left            =   7500
            TabIndex        =   137
            Text            =   "Text27"
            Top             =   2520
            Width           =   1212
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm40"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   159
            Left            =   9420
            TabIndex        =   178
            Top             =   300
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm39"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   158
            Left            =   9420
            TabIndex        =   177
            Top             =   660
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm38"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   157
            Left            =   9420
            TabIndex        =   176
            Top             =   1020
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm37"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   156
            Left            =   9420
            TabIndex        =   175
            Top             =   1380
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm36"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   155
            Left            =   9420
            TabIndex        =   174
            Top             =   1740
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm35"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   154
            Left            =   9420
            TabIndex        =   173
            Top             =   2100
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm34"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   153
            Left            =   9420
            TabIndex        =   172
            Top             =   2460
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm33"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   152
            Left            =   9420
            TabIndex        =   171
            Top             =   2820
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm32"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   151
            Left            =   9420
            TabIndex        =   170
            Top             =   3180
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm31"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   150
            Left            =   9420
            TabIndex        =   169
            Top             =   3540
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm30"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   149
            Left            =   6300
            TabIndex        =   168
            Top             =   360
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm29"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   148
            Left            =   6300
            TabIndex        =   167
            Top             =   720
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm28"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   147
            Left            =   6300
            TabIndex        =   166
            Top             =   1080
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm27"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   146
            Left            =   6300
            TabIndex        =   165
            Top             =   1440
            Width           =   1572
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4392
         Left            =   -74580
         TabIndex        =   93
         Top             =   720
         Width           =   12792
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   119
            Left            =   9960
            TabIndex        =   121
            Top             =   3720
            Width           =   250
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   119
            Left            =   7560
            TabIndex        =   120
            Text            =   "Text40"
            Top             =   240
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   118
            Left            =   7560
            TabIndex        =   119
            Text            =   "Text39"
            Top             =   600
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   117
            Left            =   7560
            TabIndex        =   118
            Text            =   "Text38"
            Top             =   960
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   116
            Left            =   7560
            TabIndex        =   117
            Text            =   "Text37"
            Top             =   1320
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   115
            Left            =   7560
            TabIndex        =   116
            Text            =   "Text36"
            Top             =   1680
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   114
            Left            =   7560
            TabIndex        =   115
            Text            =   "Text35"
            Top             =   2040
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   113
            Left            =   7560
            TabIndex        =   114
            Text            =   "Text34"
            Top             =   2400
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   112
            Left            =   7560
            TabIndex        =   113
            Text            =   "Text33"
            Top             =   2760
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   111
            Left            =   7560
            TabIndex        =   112
            Text            =   "Text32"
            Top             =   3120
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   110
            Left            =   7560
            TabIndex        =   111
            Text            =   "Text31"
            Top             =   3480
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   109
            Left            =   4800
            TabIndex        =   110
            Text            =   "Text30"
            Top             =   360
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   108
            Left            =   4800
            TabIndex        =   109
            Text            =   "Text29"
            Top             =   720
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   107
            Left            =   4800
            TabIndex        =   108
            Text            =   "Text28"
            Top             =   1080
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   106
            Left            =   4800
            TabIndex        =   107
            Text            =   "Text27"
            Top             =   1440
            Width           =   1212
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   118
            Left            =   3300
            TabIndex        =   106
            Top             =   360
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   117
            Left            =   3300
            TabIndex        =   105
            Top             =   600
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   116
            Left            =   3300
            TabIndex        =   104
            Top             =   900
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   115
            Left            =   3300
            TabIndex        =   103
            Top             =   1260
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   114
            Left            =   3300
            TabIndex        =   102
            Top             =   1560
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   113
            Left            =   3300
            TabIndex        =   101
            Top             =   1860
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   112
            Left            =   3300
            TabIndex        =   100
            Top             =   2160
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   111
            Left            =   3300
            TabIndex        =   99
            Top             =   2400
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   110
            Left            =   3300
            TabIndex        =   98
            Top             =   2700
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   109
            Left            =   3300
            TabIndex        =   97
            Top             =   2940
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   108
            Left            =   6000
            TabIndex        =   96
            Top             =   360
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   107
            Left            =   6000
            TabIndex        =   95
            Top             =   840
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   106
            Left            =   6060
            TabIndex        =   94
            Top             =   1200
            Width           =   250
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm40"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   119
            Left            =   10440
            TabIndex        =   135
            Top             =   480
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm39"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   118
            Left            =   10440
            TabIndex        =   134
            Top             =   720
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm38"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   117
            Left            =   10440
            TabIndex        =   133
            Top             =   960
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm37"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   116
            Left            =   10440
            TabIndex        =   132
            Top             =   1200
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm36"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   115
            Left            =   10440
            TabIndex        =   131
            Top             =   1440
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm35"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   114
            Left            =   10440
            TabIndex        =   130
            Top             =   1680
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm34"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   113
            Left            =   10440
            TabIndex        =   129
            Top             =   1920
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm33"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   112
            Left            =   10440
            TabIndex        =   128
            Top             =   2160
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm32"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   111
            Left            =   10440
            TabIndex        =   127
            Top             =   2400
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm31"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   110
            Left            =   10440
            TabIndex        =   126
            Top             =   2640
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm30"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   109
            Left            =   10440
            TabIndex        =   125
            Top             =   2880
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm29"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   108
            Left            =   10440
            TabIndex        =   124
            Top             =   3120
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm28"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   107
            Left            =   3600
            TabIndex        =   123
            Top             =   2880
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm27"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   106
            Left            =   3600
            TabIndex        =   122
            Top             =   2520
            Width           =   1572
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4392
         Left            =   -74580
         TabIndex        =   50
         Top             =   780
         Width           =   12792
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   79
            Left            =   11700
            TabIndex        =   78
            Top             =   3300
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   78
            Left            =   11760
            TabIndex        =   77
            Top             =   2940
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   77
            Left            =   11700
            TabIndex        =   76
            Top             =   2580
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   76
            Left            =   11700
            TabIndex        =   75
            Top             =   2220
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   75
            Left            =   11760
            TabIndex        =   74
            Top             =   1860
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   74
            Left            =   11760
            TabIndex        =   73
            Top             =   1500
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   73
            Left            =   11820
            TabIndex        =   72
            Top             =   1200
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   72
            Left            =   11760
            TabIndex        =   71
            Top             =   780
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   71
            Left            =   11760
            TabIndex        =   70
            Top             =   420
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   70
            Left            =   8760
            TabIndex        =   69
            Top             =   3660
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   69
            Left            =   8760
            TabIndex        =   68
            Top             =   3300
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   68
            Left            =   8760
            TabIndex        =   67
            Top             =   2940
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   67
            Left            =   8760
            TabIndex        =   66
            Top             =   2580
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   66
            Left            =   8760
            TabIndex        =   65
            Top             =   2220
            Width           =   250
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   79
            Left            =   10440
            TabIndex        =   64
            Text            =   "Text40"
            Top             =   3600
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   78
            Left            =   10440
            TabIndex        =   63
            Text            =   "Text39"
            Top             =   3240
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   77
            Left            =   10440
            TabIndex        =   62
            Text            =   "Text38"
            Top             =   2880
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   76
            Left            =   10440
            TabIndex        =   61
            Text            =   "Text37"
            Top             =   2520
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   75
            Left            =   10440
            TabIndex        =   60
            Text            =   "Text36"
            Top             =   2160
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   74
            Left            =   10440
            TabIndex        =   59
            Text            =   "Text35"
            Top             =   1800
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   73
            Left            =   10440
            TabIndex        =   58
            Text            =   "Text34"
            Top             =   1440
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   72
            Left            =   10440
            TabIndex        =   57
            Text            =   "Text33"
            Top             =   1080
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   71
            Left            =   10440
            TabIndex        =   56
            Text            =   "Text32"
            Top             =   720
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   70
            Left            =   10440
            TabIndex        =   55
            Text            =   "Text31"
            Top             =   360
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   69
            Left            =   7500
            TabIndex        =   54
            Text            =   "Text30"
            Top             =   3600
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   68
            Left            =   7500
            TabIndex        =   53
            Text            =   "Text29"
            Top             =   3240
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   67
            Left            =   7500
            TabIndex        =   52
            Text            =   "Text28"
            Top             =   2880
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   66
            Left            =   7500
            TabIndex        =   51
            Text            =   "Text27"
            Top             =   2520
            Width           =   1212
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm40"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   79
            Left            =   9420
            TabIndex        =   92
            Top             =   300
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm39"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   78
            Left            =   9420
            TabIndex        =   91
            Top             =   660
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm38"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   77
            Left            =   9420
            TabIndex        =   90
            Top             =   1020
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm37"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   76
            Left            =   9420
            TabIndex        =   89
            Top             =   1380
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm36"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   75
            Left            =   9420
            TabIndex        =   88
            Top             =   1740
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm35"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   74
            Left            =   9420
            TabIndex        =   87
            Top             =   2100
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm34"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   73
            Left            =   9420
            TabIndex        =   86
            Top             =   2460
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm33"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   72
            Left            =   9420
            TabIndex        =   85
            Top             =   2820
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm32"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   71
            Left            =   9420
            TabIndex        =   84
            Top             =   3180
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm31"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   70
            Left            =   9420
            TabIndex        =   83
            Top             =   3540
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm30"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   69
            Left            =   6300
            TabIndex        =   82
            Top             =   360
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm29"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   68
            Left            =   6300
            TabIndex        =   81
            Top             =   720
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm28"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   67
            Left            =   6300
            TabIndex        =   80
            Top             =   1080
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm27"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   66
            Left            =   6300
            TabIndex        =   79
            Top             =   1440
            Width           =   1572
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5352
         Left            =   -74640
         TabIndex        =   11
         Top             =   660
         Width           =   12792
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   39
            Left            =   7020
            TabIndex        =   39
            Top             =   3780
            Width           =   250
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   26
            Left            =   1740
            TabIndex        =   38
            Text            =   "Text27"
            Top             =   600
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   27
            Left            =   1740
            TabIndex        =   37
            Text            =   "Text28"
            Top             =   960
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   28
            Left            =   1800
            TabIndex        =   36
            Text            =   "Text29"
            Top             =   1320
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   29
            Left            =   1920
            TabIndex        =   35
            Text            =   "Text30"
            Top             =   1740
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   30
            Left            =   5700
            TabIndex        =   34
            Text            =   "Text31"
            Top             =   480
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   31
            Left            =   5700
            TabIndex        =   33
            Text            =   "Text32"
            Top             =   840
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   32
            Left            =   5700
            TabIndex        =   32
            Text            =   "Text33"
            Top             =   1200
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   33
            Left            =   5700
            TabIndex        =   31
            Text            =   "Text34"
            Top             =   1560
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   34
            Left            =   5700
            TabIndex        =   30
            Text            =   "Text35"
            Top             =   1920
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   35
            Left            =   5700
            TabIndex        =   29
            Text            =   "Text36"
            Top             =   2280
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   36
            Left            =   5700
            TabIndex        =   28
            Text            =   "Text37"
            Top             =   2640
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   37
            Left            =   5700
            TabIndex        =   27
            Text            =   "Text38"
            Top             =   3000
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   38
            Left            =   5700
            TabIndex        =   26
            Text            =   "Text39"
            Top             =   3360
            Width           =   1212
         End
         Begin VB.TextBox DUT1_ParmBox 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   39
            Left            =   5700
            TabIndex        =   25
            Text            =   "Text40"
            Top             =   3720
            Width           =   1212
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   26
            Left            =   660
            TabIndex        =   24
            Top             =   660
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   27
            Left            =   660
            TabIndex        =   23
            Top             =   1020
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   28
            Left            =   660
            TabIndex        =   22
            Top             =   1320
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   29
            Left            =   660
            TabIndex        =   21
            Top             =   1740
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   30
            Left            =   7020
            TabIndex        =   20
            Top             =   540
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   31
            Left            =   7020
            TabIndex        =   19
            Top             =   900
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   32
            Left            =   7080
            TabIndex        =   18
            Top             =   1320
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   33
            Left            =   7020
            TabIndex        =   17
            Top             =   1620
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   34
            Left            =   7020
            TabIndex        =   16
            Top             =   1980
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   35
            Left            =   6960
            TabIndex        =   15
            Top             =   2340
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   36
            Left            =   6960
            TabIndex        =   14
            Top             =   2700
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   37
            Left            =   7020
            TabIndex        =   13
            Top             =   3060
            Width           =   250
         End
         Begin VB.CommandButton HelpMeParmDUT1 
            Caption         =   "..."
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
            Index           =   38
            Left            =   6960
            TabIndex        =   12
            Top             =   3420
            Width           =   250
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm27"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   26
            Left            =   1080
            TabIndex        =   383
            Top             =   600
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm28"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   27
            Left            =   1080
            TabIndex        =   382
            Top             =   960
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm29"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   28
            Left            =   1080
            TabIndex        =   381
            Top             =   1260
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm30"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   29
            Left            =   1080
            TabIndex        =   380
            Top             =   1680
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm31"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   30
            Left            =   3840
            TabIndex        =   49
            Top             =   540
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm32"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   31
            Left            =   3840
            TabIndex        =   48
            Top             =   900
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm33"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   32
            Left            =   3720
            TabIndex        =   47
            Top             =   1320
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm34"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   33
            Left            =   3660
            TabIndex        =   46
            Top             =   1680
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm35"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   34
            Left            =   3540
            TabIndex        =   45
            Top             =   1980
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm36"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   35
            Left            =   3360
            TabIndex        =   44
            Top             =   2220
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm37"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   36
            Left            =   3600
            TabIndex        =   43
            Top             =   2640
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm38"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   37
            Left            =   3600
            TabIndex        =   42
            Top             =   3060
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm39"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   38
            Left            =   3840
            TabIndex        =   41
            Top             =   3420
            Width           =   1572
         End
         Begin VB.Label DUT1_Label 
            Caption         =   "Parm40"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   39
            Left            =   3960
            TabIndex        =   40
            Top             =   3780
            Width           =   1572
         End
      End
      Begin VB.Frame Tab1MainFrame 
         Height          =   5412
         Left            =   180
         TabIndex        =   1
         Top             =   540
         Width           =   9975
         Begin VB.CommandButton ClearAllDut1 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   7260
            TabIndex        =   388
            Top             =   3600
            Width           =   1272
         End
         Begin VB.Frame MOSMainFrame 
            Caption         =   "Measurement Option"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5052
            Index           =   0
            Left            =   3480
            TabIndex        =   6
            Top             =   240
            Width           =   5652
            Begin VB.OptionButton InitIVSweepOption 
               Caption         =   "Initial IVSweep Only"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   312
               Index           =   0
               Left            =   2640
               TabIndex        =   706
               Top             =   192
               Width           =   2076
            End
            Begin VB.Frame SetMeasTimeIntervalFrame 
               Caption         =   "Meas. Time Interval"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1632
               Index           =   0
               Left            =   2880
               TabIndex        =   7
               Top             =   720
               Width           =   2472
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   25
                  Left            =   120
                  TabIndex        =   379
                  Top             =   960
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   24
                  Left            =   120
                  TabIndex        =   378
                  Top             =   660
                  Width           =   150
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   25
                  Left            =   1020
                  TabIndex        =   368
                  Text            =   "Text26"
                  Top             =   960
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   24
                  Left            =   1200
                  TabIndex        =   367
                  Text            =   "Text25"
                  Top             =   660
                  Width           =   1212
               End
               Begin VB.OptionButton MeasScaleRaioButton 
                  Caption         =   "Log"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   432
                  Index           =   0
                  Left            =   240
                  TabIndex        =   9
                  Top             =   240
                  Width           =   732
               End
               Begin VB.OptionButton MeasScaleRaioButton 
                  Caption         =   "Linear"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   432
                  Index           =   1
                  Left            =   1260
                  TabIndex        =   8
                  Top             =   240
                  Width           =   1272
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm25"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   24
                  Left            =   480
                  TabIndex        =   369
                  Top             =   660
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm26"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   25
                  Left            =   420
                  TabIndex        =   366
                  Top             =   960
                  Width           =   1572
               End
            End
            Begin VB.OptionButton MeasOptionRadioOption 
               Caption         =   "Spot"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   0
               Left            =   60
               TabIndex        =   308
               Top             =   300
               Width           =   912
            End
            Begin VB.OptionButton MeasOptionRadioOption 
               Caption         =   "IVSweep"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   1
               Left            =   960
               TabIndex        =   307
               Top             =   300
               Width           =   1212
            End
            Begin VB.Frame MeasTypeFrame 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4392
               Index           =   0
               Left            =   180
               TabIndex        =   10
               Top             =   540
               Width           =   2592
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   23
                  Left            =   60
                  TabIndex        =   377
                  Top             =   3900
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   22
                  Left            =   60
                  TabIndex        =   376
                  Top             =   3540
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   21
                  Left            =   60
                  TabIndex        =   375
                  Top             =   3180
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   20
                  Left            =   60
                  TabIndex        =   374
                  Top             =   2880
                  Width           =   150
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   23
                  Left            =   1200
                  TabIndex        =   373
                  Text            =   "Text24"
                  Top             =   3960
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   22
                  Left            =   1200
                  TabIndex        =   372
                  Text            =   "Text23"
                  Top             =   3600
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   21
                  Left            =   1200
                  TabIndex        =   371
                  Text            =   "Text22"
                  Top             =   3240
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   20
                  Left            =   1200
                  TabIndex        =   370
                  Text            =   "Text21"
                  Top             =   2880
                  Width           =   1212
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   19
                  Left            =   60
                  TabIndex        =   357
                  Top             =   2580
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   18
                  Left            =   60
                  TabIndex        =   356
                  Top             =   2160
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   17
                  Left            =   60
                  TabIndex        =   355
                  Top             =   1740
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   16
                  Left            =   60
                  TabIndex        =   354
                  Top             =   1380
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   15
                  Left            =   60
                  TabIndex        =   353
                  Top             =   960
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   14
                  Left            =   60
                  TabIndex        =   352
                  Top             =   720
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   13
                  Left            =   120
                  TabIndex        =   351
                  Top             =   540
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   12
                  Left            =   60
                  TabIndex        =   350
                  Top             =   300
                  Width           =   150
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   19
                  Left            =   1200
                  TabIndex        =   349
                  Text            =   "Text20"
                  Top             =   2640
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   18
                  Left            =   1200
                  TabIndex        =   348
                  Text            =   "Text19"
                  Top             =   2340
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   17
                  Left            =   1200
                  TabIndex        =   347
                  Text            =   "Text18"
                  Top             =   1980
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   16
                  Left            =   1200
                  TabIndex        =   346
                  Text            =   "Text17"
                  Top             =   1680
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   15
                  Left            =   1200
                  TabIndex        =   345
                  Text            =   "Text16"
                  Top             =   1380
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   14
                  Left            =   1200
                  TabIndex        =   344
                  Text            =   "Text15"
                  Top             =   1020
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   13
                  Left            =   1200
                  TabIndex        =   343
                  Text            =   "Text14"
                  Top             =   660
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   288
                  Index           =   12
                  Left            =   1200
                  TabIndex        =   342
                  Text            =   "Text13"
                  Top             =   300
                  Width           =   1212
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm24"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   23
                  Left            =   420
                  TabIndex        =   387
                  Top             =   3900
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm23"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   22
                  Left            =   480
                  TabIndex        =   386
                  Top             =   3600
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm22"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   21
                  Left            =   480
                  TabIndex        =   385
                  Top             =   3240
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm21"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   20
                  Left            =   420
                  TabIndex        =   384
                  Top             =   2940
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm20"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   19
                  Left            =   420
                  TabIndex        =   365
                  Top             =   2700
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm19"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   18
                  Left            =   360
                  TabIndex        =   364
                  Top             =   2280
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm18"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   17
                  Left            =   300
                  TabIndex        =   363
                  Top             =   1800
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm17"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   16
                  Left            =   300
                  TabIndex        =   362
                  Top             =   1380
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm16"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   15
                  Left            =   420
                  TabIndex        =   361
                  Top             =   1020
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm15"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   14
                  Left            =   420
                  TabIndex        =   360
                  Top             =   780
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm14"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   13
                  Left            =   420
                  TabIndex        =   359
                  Top             =   540
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm13"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   12
                  Left            =   240
                  TabIndex        =   358
                  Top             =   300
                  Width           =   1572
               End
            End
         End
         Begin VB.Frame StressOptionMainFrame 
            Caption         =   "Stress Options"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5172
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   240
            Width           =   3072
            Begin VB.OptionButton StressOptionRadioButton 
               Caption         =   "DC"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   1
               Left            =   1620
               TabIndex        =   4
               Top             =   300
               Width           =   672
            End
            Begin VB.OptionButton StressOptionRadioButton 
               Caption         =   "AC"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               Index           =   0
               Left            =   180
               TabIndex        =   3
               Top             =   300
               Width           =   672
            End
            Begin VB.Frame ACStressOptionFrame 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4452
               Index           =   0
               Left            =   120
               TabIndex        =   5
               Top             =   480
               Width           =   2832
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   11
                  Left            =   120
                  TabIndex        =   341
                  Top             =   4020
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   10
                  Left            =   120
                  TabIndex        =   340
                  Top             =   3720
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   9
                  Left            =   60
                  TabIndex        =   339
                  Top             =   3420
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   8
                  Left            =   60
                  TabIndex        =   338
                  Top             =   3000
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   7
                  Left            =   60
                  TabIndex        =   337
                  Top             =   2700
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   6
                  Left            =   60
                  TabIndex        =   336
                  Top             =   2400
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   5
                  Left            =   60
                  TabIndex        =   335
                  Top             =   2100
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   4
                  Left            =   60
                  TabIndex        =   334
                  Top             =   1740
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   3
                  Left            =   120
                  TabIndex        =   333
                  Top             =   1260
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   2
                  Left            =   60
                  TabIndex        =   332
                  Top             =   900
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   1
                  Left            =   120
                  TabIndex        =   331
                  Top             =   600
                  Width           =   150
               End
               Begin VB.CommandButton HelpMeParmDUT1 
                  Caption         =   "..."
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   150
                  Index           =   0
                  Left            =   60
                  TabIndex        =   330
                  Top             =   300
                  Width           =   150
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   11
                  Left            =   1440
                  TabIndex        =   284
                  Text            =   "Text12"
                  Top             =   4176
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   10
                  Left            =   1440
                  TabIndex        =   283
                  Text            =   "Text11"
                  Top             =   3780
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   9
                  Left            =   1440
                  TabIndex        =   282
                  Text            =   "Text10"
                  Top             =   3420
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   8
                  Left            =   1440
                  TabIndex        =   281
                  Text            =   "Text9"
                  Top             =   3060
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   7
                  Left            =   1440
                  TabIndex        =   280
                  Text            =   "Text8"
                  Top             =   2700
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   6
                  Left            =   1440
                  TabIndex        =   279
                  Text            =   "Text7"
                  Top             =   2340
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   5
                  Left            =   1440
                  TabIndex        =   278
                  Text            =   "Text6"
                  Top             =   2040
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   4
                  Left            =   1440
                  TabIndex        =   277
                  Text            =   "Text5"
                  Top             =   1680
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   3
                  Left            =   1440
                  TabIndex        =   276
                  Text            =   "Text4"
                  Top             =   1320
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   2
                  Left            =   1440
                  TabIndex        =   275
                  Text            =   "Text3"
                  Top             =   960
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   1
                  Left            =   1440
                  TabIndex        =   274
                  Text            =   "Text2"
                  Top             =   600
                  Width           =   1212
               End
               Begin VB.TextBox DUT1_ParmBox 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   0
                  Left            =   1440
                  TabIndex        =   273
                  Text            =   "Text1"
                  Top             =   240
                  Width           =   1212
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm2"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   1
                  Left            =   600
                  TabIndex        =   262
                  Top             =   660
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm1"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   156
                  Index           =   0
                  Left            =   240
                  TabIndex        =   261
                  Top             =   240
                  Width           =   432
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm12"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   11
                  Left            =   540
                  TabIndex        =   272
                  Top             =   3960
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm11"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   10
                  Left            =   540
                  TabIndex        =   271
                  Top             =   3780
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm10"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   9
                  Left            =   600
                  TabIndex        =   270
                  Top             =   3420
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm9"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   8
                  Left            =   600
                  TabIndex        =   269
                  Top             =   3060
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm8"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   7
                  Left            =   600
                  TabIndex        =   268
                  Top             =   2700
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm7"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   6
                  Left            =   600
                  TabIndex        =   267
                  Top             =   2400
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm6"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   5
                  Left            =   660
                  TabIndex        =   266
                  Top             =   2100
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm5"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   4
                  Left            =   660
                  TabIndex        =   265
                  Top             =   1740
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm4"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   3
                  Left            =   720
                  TabIndex        =   264
                  Top             =   1320
                  Width           =   1572
               End
               Begin VB.Label DUT1_Label 
                  Caption         =   "Parm3"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   252
                  Index           =   2
                  Left            =   720
                  TabIndex        =   263
                  Top             =   960
                  Width           =   1572
               End
            End
         End
      End
   End
   Begin VB.TextBox HiddenInputString 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   5640
      TabIndex        =   302
      Text            =   "DUT1:This is invisible during run time"
      Top             =   6780
      Width           =   7872
   End
   Begin VB.TextBox HiddenInputString 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   5640
      TabIndex        =   303
      Text            =   "DUT2:This is invisible during run time"
      Top             =   7140
      Width           =   7872
   End
   Begin VB.TextBox HiddenInputString 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   5640
      TabIndex        =   304
      Text            =   "DUT3:This is invisible during run time"
      Top             =   7500
      Width           =   7872
   End
   Begin VB.TextBox HiddenInputString 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   5640
      TabIndex        =   305
      Text            =   "DUT4:This is invisible during run time"
      Top             =   7860
      Width           =   7872
   End
   Begin VB.TextBox HiddenInputString 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   5640
      TabIndex        =   306
      Text            =   "DUT5:This is invisible during run time"
      Top             =   8220
      Width           =   7872
   End
End
Attribute VB_Name = "WGFMUVBInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Const BIF_USENEWUI = &H40

Private Declare Function SHBrowseForFolder Lib _
"shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib _
"shell32" (ByVal pidList As Long, ByVal lpBuffer _
As String) As Long

Private Declare Function lstrcat Lib "kernel32" _
Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Private Sub ChooseDir_Click()
'Opens a Browse Folders Dialog Box that displays the directories in your computer
'Declare Varibles
Dim lpIDList As Long
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

' Display information
szTitle = "Select a working directory where data will be saved:"

With tBrowseInfo
   .hWndOwner = Me.hWnd ' Owner Form
   .lpszTitle = lstrcat(szTitle, "")
   .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_USENEWUI
End With

lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
   sBuffer = Space(MAX_PATH)
   SHGetPathFromIDList lpIDList, sBuffer
   sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
   'MsgBox sBuffer
   DataPathLabel.Caption = sBuffer
End If
End Sub
Private Sub ClearAllDut1_Click()
Dim Index_of_SavePath As Integer
Index_of_SavePath = 35
iStart = 0
iEnd = iStart + 39
For i = iStart To iEnd
 DUT1_ParmBox(i) = ""
Next
DUT1_ParmBox(iStart + Index_of_SavePath) = DataPathLabel.Caption
End Sub

Private Sub ClearAllDut2_Click()
Dim Index_of_SavePath As Integer
Index_of_SavePath = 35
iStart = 40
iEnd = iStart + 39
For i = iStart To iEnd
 DUT1_ParmBox(i) = ""
Next
DUT1_ParmBox(iStart + Index_of_SavePath) = DataPathLabel.Caption
End Sub

Private Sub ClearAllDut3_Click()
Dim Index_of_SavePath As Integer
Index_of_SavePath = 35
iStart = 80
iEnd = iStart + 39
For i = iStart To iEnd
 DUT1_ParmBox(i) = ""
Next
DUT1_ParmBox(iStart + Index_of_SavePath) = DataPathLabel.Caption
End Sub

Private Sub ClearAllDut4_Click()
Dim Index_of_SavePath As Integer
Index_of_SavePath = 35
iStart = 120
iEnd = iStart + 39
For i = iStart To iEnd
 DUT1_ParmBox(i) = ""
Next
DUT1_ParmBox(iStart + Index_of_SavePath) = DataPathLabel.Caption
End Sub

Private Sub ClearAllDut5_Click()
Dim Index_of_SavePath As Integer
Index_of_SavePath = 35
iStart = 160
iEnd = iStart + 39
For i = iStart To iEnd
 DUT1_ParmBox(i) = ""
Next
DUT1_ParmBox(iStart + Index_of_SavePath) = DataPathLabel.Caption
End Sub
Private Sub Command3_Click()
MsgBox (Channel_1_DUT1)
MsgBox (Channel_2_DUT1)
MsgBox (Channel_3_DUT1)
MsgBox (Channel_4_DUT1)

MsgBox (Channel_1_DUT2)
MsgBox (Channel_2_DUT2)
MsgBox (Channel_3_DUT2)
MsgBox (Channel_4_DUT2)

End Sub



Private Sub CommnGateNumDevWGFMU_Change()
'---------------------------------------------------
'If user makes this blank or place 0 then force the value to be 1
'so that Tab for DUT #1 is always displayed

MaxNumDevicesCG = 9

If (CommnGateNumDevWGFMU.Text = "") Then
    MsgBox ("Value must be between 2 and 9. Use above individual option to stress only one device.")
End If

If (Val(CommnGateNumDevWGFMU.Text) = 0 Or Val(CommnGateNumDevWGFMU.Text) = 1) Then
    MsgBox ("Value must be between 2 and 9. Use above individual option to stress only one device.")
    CommnGateNumDevWGFMU.Text = 2
End If
If (Val(CommnGateNumDevWGFMU.Text) > MaxNumDevicesCG) Then
    MsgBox ("Maximum allowed common-gate devices to stress = " & Str$(MaxNumDevicesCG))
    CommnGateNumDevWGFMU.Text = 9
End If
End Sub
Private Sub ConfBTIButton_Click()
ConfigFastBTIStress.Show
End Sub

Private Sub Copy_Input_for_DUT2_Click()
Dim InputCopyfromDutNum As Integer
InputCopyfromDutNum = Val(CopyInputFromforDUT2.Text)
istartThisTab = 40

Select Case InputCopyfromDutNum

Case 1
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(2).value = StressOptionRadioButton(0).value
 StressOptionRadioButton(3).value = StressOptionRadioButton(1).value
 MeasOptionRadioOption(2).value = MeasOptionRadioOption(0).value
 MeasOptionRadioOption(3).value = MeasOptionRadioOption(1).value
 MeasScaleRaioButton(2).value = MeasScaleRaioButton(0).value
 MeasScaleRaioButton(3).value = MeasScaleRaioButton(1).value

Case 3
iStart = 80
iEnd = iStart + 39
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(2).value = StressOptionRadioButton(4).value
 StressOptionRadioButton(3).value = StressOptionRadioButton(5).value
 MeasOptionRadioOption(2).value = MeasOptionRadioOption(4).value
 MeasOptionRadioOption(3).value = MeasOptionRadioOption(5).value
 MeasScaleRaioButton(2).value = MeasScaleRaioButton(4).value
 MeasScaleRaioButton(3).value = MeasScaleRaioButton(5).value

Case 4
iStart = 120
iEnd = iStart + 39
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(2).value = StressOptionRadioButton(6).value
 StressOptionRadioButton(3).value = StressOptionRadioButton(7).value
 MeasOptionRadioOption(2).value = MeasOptionRadioOption(6).value
 MeasOptionRadioOption(3).value = MeasOptionRadioOption(7).value
 MeasScaleRaioButton(2).value = MeasScaleRaioButton(6).value
 MeasScaleRaioButton(3).value = MeasScaleRaioButton(7).value

Case 5
iStart = 160
iEnd = iStart + 39
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(2).value = StressOptionRadioButton(8).value
 StressOptionRadioButton(3).value = StressOptionRadioButton(9).value
 MeasOptionRadioOption(2).value = MeasOptionRadioOption(8).value
 MeasOptionRadioOption(3).value = MeasOptionRadioOption(9).value
 MeasScaleRaioButton(2).value = MeasScaleRaioButton(8).value
 MeasScaleRaioButton(3).value = MeasScaleRaioButton(9).value

End Select

End Sub

Private Sub Copy_Input_for_DUT3_Click()
Dim InputCopyfromDutNum As Integer
InputCopyfromDutNum = Val(CopyInputFromforDUT3.Text)
istartThisTab = 80

Select Case InputCopyfromDutNum

Case 1
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(4).value = StressOptionRadioButton(0).value
 StressOptionRadioButton(5).value = StressOptionRadioButton(1).value
 MeasOptionRadioOption(4).value = MeasOptionRadioOption(0).value
 MeasOptionRadioOption(5).value = MeasOptionRadioOption(1).value
 MeasScaleRaioButton(4).value = MeasScaleRaioButton(0).value
 MeasScaleRaioButton(5).value = MeasScaleRaioButton(1).value

Case 2
iStart = 40
iEnd = iStart + 39
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(4).value = StressOptionRadioButton(2).value
 StressOptionRadioButton(5).value = StressOptionRadioButton(3).value
 MeasOptionRadioOption(4).value = MeasOptionRadioOption(2).value
 MeasOptionRadioOption(5).value = MeasOptionRadioOption(3).value
 MeasScaleRaioButton(4).value = MeasScaleRaioButton(2).value
 MeasScaleRaioButton(5).value = MeasScaleRaioButton(3).value

Case 4
iStart = 120
iEnd = iStart + 39
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(4).value = StressOptionRadioButton(6).value
 StressOptionRadioButton(5).value = StressOptionRadioButton(7).value
 MeasOptionRadioOption(4).value = MeasOptionRadioOption(6).value
 MeasOptionRadioOption(5).value = MeasOptionRadioOption(7).value
 MeasScaleRaioButton(4).value = MeasScaleRaioButton(6).value
 MeasScaleRaioButton(5).value = MeasScaleRaioButton(7).value

Case 5
iStart = 160
iEnd = iStart + 39
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(4).value = StressOptionRadioButton(8).value
 StressOptionRadioButton(5).value = StressOptionRadioButton(9).value
 MeasOptionRadioOption(4).value = MeasOptionRadioOption(8).value
 MeasOptionRadioOption(5).value = MeasOptionRadioOption(9).value
 MeasScaleRaioButton(4).value = MeasScaleRaioButton(8).value
 MeasScaleRaioButton(5).value = MeasScaleRaioButton(9).value

End Select
End Sub

Private Sub Copy_Input_for_DUT4_Click()
Dim InputCopyfromDutNum As Integer
InputCopyfromDutNum = Val(CopyInputFromforDUT3.Text)
istartThisTab = 120

Select Case InputCopyfromDutNum

Case 1
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(6).value = StressOptionRadioButton(0).value
 StressOptionRadioButton(7).value = StressOptionRadioButton(1).value
 MeasOptionRadioOption(6).value = MeasOptionRadioOption(0).value
 MeasOptionRadioOption(7).value = MeasOptionRadioOption(1).value
 MeasScaleRaioButton(6).value = MeasScaleRaioButton(0).value
 MeasScaleRaioButton(7).value = MeasScaleRaioButton(1).value

Case 2
iStart = 40
iEnd = iStart + 39
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(6).value = StressOptionRadioButton(2).value
 StressOptionRadioButton(7).value = StressOptionRadioButton(3).value
 MeasOptionRadioOption(6).value = MeasOptionRadioOption(2).value
 MeasOptionRadioOption(7).value = MeasOptionRadioOption(3).value
 MeasScaleRaioButton(6).value = MeasScaleRaioButton(2).value
 MeasScaleRaioButton(7).value = MeasScaleRaioButton(3).value

Case 3
iStart = 80
iEnd = iStart + 39
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(6).value = StressOptionRadioButton(4).value
 StressOptionRadioButton(7).value = StressOptionRadioButton(5).value
 MeasOptionRadioOption(6).value = MeasOptionRadioOption(4).value
 MeasOptionRadioOption(7).value = MeasOptionRadioOption(5).value
 MeasScaleRaioButton(6).value = MeasScaleRaioButton(4).value
 MeasScaleRaioButton(7).value = MeasScaleRaioButton(5).value

Case 5
iStart = 160
iEnd = iStart + 39
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(6).value = StressOptionRadioButton(8).value
 StressOptionRadioButton(7).value = StressOptionRadioButton(9).value
 MeasOptionRadioOption(6).value = MeasOptionRadioOption(8).value
 MeasOptionRadioOption(7).value = MeasOptionRadioOption(9).value
 MeasScaleRaioButton(6).value = MeasScaleRaioButton(8).value
 MeasScaleRaioButton(7).value = MeasScaleRaioButton(9).value

End Select
End Sub

Private Sub Copy_Input_for_DUT5_Click()
Dim InputCopyfromDutNum As Integer
InputCopyfromDutNum = Val(CopyInputFromforDUT3.Text)
istartThisTab = 160

Select Case InputCopyfromDutNum

Case 1
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(8).value = StressOptionRadioButton(0).value
 StressOptionRadioButton(9).value = StressOptionRadioButton(1).value
 MeasOptionRadioOption(8).value = MeasOptionRadioOption(0).value
 MeasOptionRadioOption(9).value = MeasOptionRadioOption(1).value
 MeasScaleRaioButton(8).value = MeasScaleRaioButton(0).value
 MeasScaleRaioButton(9).value = MeasScaleRaioButton(1).value

Case 2
iStart = 40
iEnd = iStart + 39
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(8).value = StressOptionRadioButton(2).value
 StressOptionRadioButton(9).value = StressOptionRadioButton(3).value
 MeasOptionRadioOption(8).value = MeasOptionRadioOption(2).value
 MeasOptionRadioOption(9).value = MeasOptionRadioOption(3).value
 MeasScaleRaioButton(8).value = MeasScaleRaioButton(2).value
 MeasScaleRaioButton(9).value = MeasScaleRaioButton(3).value

Case 3
iStart = 80
iEnd = iStart + 39
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(8).value = StressOptionRadioButton(4).value
 StressOptionRadioButton(9).value = StressOptionRadioButton(5).value
 MeasOptionRadioOption(8).value = MeasOptionRadioOption(4).value
 MeasOptionRadioOption(9).value = MeasOptionRadioOption(5).value
 MeasScaleRaioButton(8).value = MeasScaleRaioButton(4).value
 MeasScaleRaioButton(9).value = MeasScaleRaioButton(5).value

Case 4
iStart = 120
iEnd = iStart + 39
 iStart = 0
 iEnd = iStart + 39
 For i = iStart To iEnd
  DUT1_ParmBox(istartThisTab).Text = DUT1_ParmBox(i).Text
  istartThisTab = istartThisTab + 1
 Next
 StressOptionRadioButton(8).value = StressOptionRadioButton(6).value
 StressOptionRadioButton(9).value = StressOptionRadioButton(7).value
 MeasOptionRadioOption(8).value = MeasOptionRadioOption(6).value
 MeasOptionRadioOption(9).value = MeasOptionRadioOption(7).value
 MeasScaleRaioButton(8).value = MeasScaleRaioButton(6).value
 MeasScaleRaioButton(9).value = MeasScaleRaioButton(7).value

End Select
End Sub

Private Sub CopyInputFromforDUT2_Change()
NumValMax = Val(NumDevWGFMU.Text)
'Display or hide tabs depending on user's choice on number of devices to stress
If (Val(CopyInputFromforDUT2.Text) > NumValMax Or Val(CopyInputFromforDUT2.Text) < 1) Then
    MsgBox ("Value must be between 1 and " & NumValMax & " except 2.")
    CopyInputFromforDUT2.Text = 1
ElseIf (Val(CopyInputFromforDUT2.Text) = 2) Then
    MsgBox ("Value must be between 1 and " & NumValMax & " except 2.")
    CopyInputFromforDUT2.Text = 1
End If
End Sub

Private Sub CopyInputFromforDUT3_Change()
NumValMax = Val(NumDevWGFMU.Text)
'Display or hide tabs depending on user's choice on number of devices to stress
If (Val(CopyInputFromforDUT3.Text) > NumValMax Or Val(CopyInputFromforDUT3.Text) < 1) Then
    MsgBox ("Value must be between 1 and " & NumValMax & " except 3.")
    CopyInputFromforDUT3.Text = 1
ElseIf (Val(CopyInputFromforDUT3.Text) = 3) Then
    MsgBox ("Value must be between 1 and " & NumValMax & " except 3.")
    CopyInputFromforDUT3.Text = 1
End If
End Sub


Private Sub CopyInputFromforDUT4_Change()
NumValMax = Val(NumDevWGFMU.Text)
'Display or hide tabs depending on user's choice on number of devices to stress
If (Val(CopyInputFromforDUT4.Text) > NumValMax Or Val(CopyInputFromforDUT4.Text) < 1) Then
    MsgBox ("Value must be between 1 and " & NumValMax & " except 4.")
    CopyInputFromforDUT4.Text = 1
ElseIf (Val(CopyInputFromforDUT4.Text) = 4) Then
    MsgBox ("Value must be between 1 and " & NumValMax & " except 4.")
    CopyInputFromforDUT4.Text = 1
End If
End Sub

Private Sub CopyInputFromforDUT5_Change()
NumValMax = Val(NumDevWGFMU.Text)
'Display or hide tabs depending on user's choice on number of devices to stress
If (Val(CopyInputFromforDUT5.Text) > NumValMax Or Val(CopyInputFromforDUT5.Text) < 1) Then
    MsgBox ("Value must be between 1 and " & NumValMax & " except 5.")
    CopyInputFromforDUT5.Text = 1
ElseIf (Val(CopyInputFromforDUT5.Text) = 5) Then
    MsgBox ("Value must be between 1 and " & NumValMax & " except 5.")
    CopyInputFromforDUT5.Text = 1
End If
End Sub

Private Sub DataPathLabel_Change()
Dim CycleArray(4) As Integer
Dim Index_of_SavePath As Integer
Index_of_SavePath = 35
CycleArray(0) = Index_of_SavePath + 0 * 40
CycleArray(1) = Index_of_SavePath + 1 * 40
CycleArray(2) = Index_of_SavePath + 2 * 40
CycleArray(3) = Index_of_SavePath + 3 * 40
CycleArray(4) = Index_of_SavePath + 4 * 40
For i = 0 To 4
 DUT1_ParmBox(CycleArray(i)).Text = DataPathLabel.Caption
Next
End Sub

Private Sub DUT1_ParmBox_Change(index As Integer)
'    If Index = 0 Then
'        If IsNumeric(DUT1_ParmBox(Index).Text) = false Then
'            MsgBox ("Please enter number")
'        End If
'    End If
End Sub

Private Sub DUTStressOptionCommonGate_Click()
'---------------------------------------------------
'Change font color selected option to blue
 DUTStressOptionIndividual.ForeColor = &H80000012
 DUTStressOptionCommonGate.ForeColor = &HFF0000
 '
 MsgBox ("Common-gate Stress option - Under construction!")
 DUTStressOptionIndividual.ForeColor = &HFF0000
 DUTStressOptionCommonGate.ForeColor = &H80000012
 DUTStressOptionCommonGate.value = False
 DUTStressOptionIndividual.value = True
 Exit Sub
Exit Sub
'---------------------------------------------------
'Change font color selected option to blue
DUTStressOptionIndividual.ForeColor = &H80000012
DUTStressOptionCommonGate.ForeColor = &HFF0000
'---------------------------------------------------
For i = 1 To 4
 SSTab1.TabEnabled(i) = False
 SSTab1.TabVisible(i) = False
Next
SSTab1.Caption = "CG DUT"
NumDevWGFMU.Enabled = False
CommnGateNumDevWGFMU.Enabled = True
End Sub

Private Sub DUTStressOptionIndividual_Click()
'---------------------------------------------------
'Change font color selected option to blue
DUTStressOptionIndividual.ForeColor = &HFF0000
DUTStressOptionCommonGate.ForeColor = &H80000012
'---------------------------------------------------
'Display or hide tabs depending on user's choice on number of devices to stress
 For i = 0 To Val(NumDevWGFMU.Text) - 1
  SSTab1.TabEnabled(i) = True
  SSTab1.TabVisible(i) = True
 Next
'---------------------------------------------------
'Set tab names
SSTab1.TabCaption(0) = "DUT #1"
SSTab1.TabCaption(1) = "DUT #2"
SSTab1.TabCaption(2) = "DUT #3"
SSTab1.TabCaption(3) = "DUT #4"
SSTab1.TabCaption(4) = "DUT #5"
'---------------------------------------------------
'Disable common-gate stress option
'and only enable individual device stress option
NumDevWGFMU.Enabled = True
CommnGateNumDevWGFMU.Enabled = False
End Sub
Private Sub Form_Load()
'---------------------------------------------------
NumPopup_ConfigFastBTIStress = 0
'---------------------------------------------------
LotIDTextBox.Width = 2300
WaferIDTexBox.Width = LotIDTextBox.Width
'---------------------------------------------------
ConfBTIButton.Left = 6250
ConfBTIButton.Top = 500
ConfBTIButton.Height = 700
ConfBTIButton.Width = 2170

SaveInputAsButton.Left = ConfBTIButton.Left
SaveInputAsButton.Height = ConfBTIButton.Height
SaveInputAsButton.Width = ConfBTIButton.Width
SaveInputAsButton.Top = ConfBTIButton.Top + ConfBTIButton.Height + 100

RetrieveInputButton.Left = ConfBTIButton.Left
RetrieveInputButton.Height = ConfBTIButton.Height
RetrieveInputButton.Width = ConfBTIButton.Width
RetrieveInputButton.Top = SaveInputAsButton.Top + SaveInputAsButton.Height + 100

RunButton.Left = ConfBTIButton.Left
RunButton.Height = ConfBTIButton.Height
RunButton.Width = ConfBTIButton.Width
RunButton.Top = RetrieveInputButton.Top + RetrieveInputButton.Height + 100
'---------------------------------------------------
ChooseDir.Left = 120
ChooseDir.Top = 2200
ChooseDir.Width = 2600
'---------------------------------------------------
SpecifyEXE.Left = ChooseDir.Left
SpecifyEXE.Top = 2940
SpecifyEXE.Width = 3500
'---------------------------------------------------
'Disable run button until Channel Config is done
If (NumPopup_ConfigFastBTIStress = 0) Then
 RunButton.Enabled = False
End If
'---------------------------------------------------
WGFMUVBInput.Width = 8250 + 500
WGFMUVBInput.Height = 10545
'---------------------------------------------------
'Hide all temporary tabs that are not to be shown to users
For i = 5 To 9
 SSTab1.TabVisible(i) = False
Next
'---------------------------------------------------
'Make 5 text boxes invisible --> string to C#
For i = 0 To 4
 HiddenInputString(i).Visible = False
Next
'---------------------------------------------------
'Number of input boxes per tab
NumInput = 40
'---------------------------------------------------
'Set default radio option buttons
For i = 0 To 8 Step 2
 StressOptionRadioButton(i).value = True
 MeasOptionRadioOption(i).value = True
 MeasScaleRaioButton(i).value = True
Next
'---------------------------------------------------
'Set Tab width and height
SSTab1.Width = 8000 + 600
SSTab1.Height = 6000
'-------------------------------
'Bottom frame width
BottomFrame.Width = SSTab1.Width
BottomFrame.Height = 3975
'-------------------------------
'Size and Position of MainFrame and SubFrames for each DUT tab
'MainFrame
WidthMainFrame = 7750
HeightMainFrame = 5600
Tab1MainFrame.Left = 180
Tab1MainFrame.Top = 300
Tab1MainFrame.Width = WidthMainFrame + 575
Tab1MainFrame.Height = HeightMainFrame
'
Tab2MainFrame.Left = Tab1MainFrame.Left
Tab2MainFrame.Top = Tab1MainFrame.Top
Tab2MainFrame.Width = Tab1MainFrame.Width
Tab2MainFrame.Height = Tab1MainFrame.Height
'
Tab3MainFrame.Left = Tab1MainFrame.Left
Tab3MainFrame.Top = Tab1MainFrame.Top
Tab3MainFrame.Width = Tab1MainFrame.Width
Tab3MainFrame.Height = Tab1MainFrame.Height
'
Tab4MainFrame.Left = Tab1MainFrame.Left
Tab4MainFrame.Top = Tab1MainFrame.Top
Tab4MainFrame.Width = Tab1MainFrame.Width
Tab4MainFrame.Height = Tab1MainFrame.Height
'
Tab5MainFrame.Left = Tab1MainFrame.Left
Tab5MainFrame.Top = Tab1MainFrame.Top
Tab5MainFrame.Width = Tab1MainFrame.Width
Tab5MainFrame.Height = Tab1MainFrame.Height
'-------------------------------
'Stress Option Frame
For i = 0 To 4
 StressOptionMainFrame(i).Left = 120
 StressOptionMainFrame(i).Top = 240
 StressOptionMainFrame(i).Width = 2520 + 200
 StressOptionMainFrame(i).Height = 4850
Next
 'AC radio button
 For i = 0 To 8 Step 2
  StressOptionRadioButton(i).Left = 180
  StressOptionRadioButton(i).Top = 300
 Next
'DC radio button
 For i = 1 To 9 Step 2
  StressOptionRadioButton(i).Left = 1620
  StressOptionRadioButton(i).Top = 300
 Next
 'Frame for stress parameters
For i = 0 To 4
 ACStressOptionFrame(i).Left = 60
 ACStressOptionFrame(i).Top = 500
 ACStressOptionFrame(i).Width = 2400 + 200
 ACStressOptionFrame(i).Height = 4300
Next
'-------------------------------
'Measurement Option Frame
For i = 0 To 4
 MOSMainFrame(i).Left = StressOptionMainFrame(i).Left + StressOptionMainFrame(i).Width + 50
 MOSMainFrame(i).Top = StressOptionMainFrame(i).Top
 MOSMainFrame(i).Width = 2# * StressOptionMainFrame(i).Width - 100
 MOSMainFrame(i).Height = StressOptionMainFrame(i).Height
Next
'Measurement type frame
For i = 0 To 4
 MeasTypeFrame(i).Left = ACStressOptionFrame(i).Left
 MeasTypeFrame(i).Top = ACStressOptionFrame(i).Top
 MeasTypeFrame(i).Width = ACStressOptionFrame(i).Width
 MeasTypeFrame(i).Height = ACStressOptionFrame(i).Height
Next
 'Spot measure radio button
 For i = 0 To 8 Step 2
  MeasOptionRadioOption(i).Left = 120
  MeasOptionRadioOption(i).Top = StressOptionRadioButton(i).Top
  MeasOptionRadioOption(i).Width = 800
 Next
 'IVSweep radio button
 For i = 1 To 9 Step 2
  MeasOptionRadioOption(i).Left = 1000
  MeasOptionRadioOption(i).Top = StressOptionRadioButton(i).Top
  MeasOptionRadioOption(i).Width = 1210
 Next
'Meas. time interval frame
For i = 0 To 4
 SetMeasTimeIntervalFrame(i).Left = MeasTypeFrame(i).Left + MeasTypeFrame(i).Width + 25
 SetMeasTimeIntervalFrame(i).Top = MeasTypeFrame(i).Top
 SetMeasTimeIntervalFrame(i).Width = MeasTypeFrame(i).Width
 SetMeasTimeIntervalFrame(i).Height = 50 + (1 / 4) * MeasTypeFrame(i).Height
Next
 'Meas. time Log interval radio buttons
 For i = 0 To 8 Step 2
  MeasScaleRaioButton(i).Left = 220
  MeasScaleRaioButton(i).Top = 340
  MeasScaleRaioButton(i).Width = 800
 Next
 'Meas. time Linear interval
 For i = 1 To 9 Step 2
  MeasScaleRaioButton(i).Left = 1100
  MeasScaleRaioButton(i).Top = 340
  MeasScaleRaioButton(i).Width = 1210
 Next
'---------------------------------------------------
'Default channel numbers
'Values will be updated from user input
Channel_1_DUT1 = 901
Channel_2_DUT1 = 902
Channel_3_DUT1 = 801
Channel_4_DUT1 = 802

Channel_1_DUT2 = 701
Channel_2_DUT2 = 702
Channel_3_DUT2 = 601
Channel_4_DUT2 = 602

Channel_1_DUT3 = 501
Channel_2_DUT3 = 502
Channel_3_DUT3 = -1
Channel_4_DUT3 = -1

Channel_1_DUT4 = -1
Channel_2_DUT4 = -1
Channel_3_DUT4 = -1
Channel_4_DUT4 = -1

Channel_1_DUT5 = -1
Channel_2_DUT5 = -1
Channel_3_DUT5 = -1
Channel_4_DUT5 = -1
'---------------------------------------------------
'Set the number of ConfigFastBTIStress form showing up
'equal to zero
NumPopup_ConfigFastBTIStress = 0
'---------------------------------------------------
'Default location for data saving
DataPathLabel.Caption = "C:\"
'---------------------------------------------------
'Location of copy_input_from_button and text box
'Clear buttons (to make input text boxes empty)
ClearAllDut1.Left = 120
ClearAllDut1.Top = 5150
ClearAllDut2.Left = ClearAllDut1.Left
ClearAllDut2.Top = ClearAllDut1.Top
ClearAllDut3.Left = ClearAllDut1.Left
ClearAllDut3.Top = ClearAllDut1.Top
ClearAllDut4.Left = ClearAllDut1.Left
ClearAllDut4.Top = ClearAllDut1.Top
ClearAllDut5.Left = ClearAllDut1.Left
ClearAllDut5.Top = ClearAllDut1.Top
'
Copy_Input_for_DUT2.Left = ClearAllDut2.Width + 500
Copy_Input_for_DUT2.Top = ClearAllDut2.Top
Copy_Input_for_DUT3.Left = Copy_Input_for_DUT2.Left
Copy_Input_for_DUT3.Top = ClearAllDut3.Top
Copy_Input_for_DUT4.Left = Copy_Input_for_DUT2.Left
Copy_Input_for_DUT4.Top = ClearAllDut4.Top
Copy_Input_for_DUT5.Left = Copy_Input_for_DUT2.Left
Copy_Input_for_DUT5.Top = ClearAllDut5.Top
'
CopyInputFromforDUT2.Left = ClearAllDut2.Width + Copy_Input_for_DUT2.Width + 600
CopyInputFromforDUT2.Top = ClearAllDut2.Top
CopyInputFromforDUT3.Left = CopyInputFromforDUT2.Left
CopyInputFromforDUT3.Top = ClearAllDut3.Top
CopyInputFromforDUT4.Left = CopyInputFromforDUT2.Left
CopyInputFromforDUT4.Top = ClearAllDut4.Top
CopyInputFromforDUT5.Left = CopyInputFromforDUT2.Left
CopyInputFromforDUT5.Top = ClearAllDut5.Top
'---------------------------------------------------
'Fill dummy parameter values in input text boxes
'This is only for debugging purposes
iStart = 0
iEnd = iStart + 39
For i = iStart To iEnd
 DUT1_ParmBox(i).Text = "D1_P" & i
Next
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
 DUT1_ParmBox(i).Text = "D2_P" & i
Next
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
 DUT1_ParmBox(i).Text = "D3_P" & i
Next
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
 DUT1_ParmBox(i).Text = "D4_P" & i
Next
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
 DUT1_ParmBox(i).Text = "D5_P" & i
Next
'---------------------------------------------------
'Define default values
Dim default_parm_values(39)
'> AC and DC
default_parm_values(0) = 10#
default_parm_values(1) = 1#
default_parm_values(2) = 0#
default_parm_values(3) = 5#
default_parm_values(4) = 0#
default_parm_values(5) = 0#
default_parm_values(6) = Format(0.00000001, "0.00E+0")
default_parm_values(7) = Format(0.00000001, "0.00E+0")
'> AC Only
default_parm_values(8) = 0#
default_parm_values(9) = 0#
default_parm_values(10) = Format(1000, "0.00E+0")
default_parm_values(11) = 50#
'> Measurement - General
default_parm_values(12) = 0#
default_parm_values(13) = 0.5
default_parm_values(14) = Format(0.000002, "0.00E+0")
default_parm_values(15) = Format(0.000001, "0.00E+0")
default_parm_values(16) = 0#
default_parm_values(17) = 50#
default_parm_values(18) = 10#
default_parm_values(19) = 40#
default_parm_values(20) = Format(0.0000001, "0.00E+0")
'> IV Sweep
default_parm_values(21) = 0#
default_parm_values(22) = -0.5
default_parm_values(23) = -0.05
'> Log. Meas. Interval
default_parm_values(24) = 3#
'> Linear Meas. Interval
default_parm_values(25) = 10#
'> Hidden and values are assigned from other controls
default_parm_values(34) = "true"    'isLog
default_parm_values(35) = DataPathLabel.Caption
default_parm_values(36) = "true"   'acStress
default_parm_values(38) = "false"  'measIV
'> Not used at the moment
default_parm_values(26) = 0#    'skew
default_parm_values(27) = "false"   'measAfterHigh
default_parm_values(28) = Format(1000000, "0.00E+0")    'cpFreqStart
default_parm_values(29) = Format(4000000, "0.00E+0")    'cpFreqStep
default_parm_values(30) = 1#    'cpNumSteps
default_parm_values(31) = 1#    'cpHigh
default_parm_values(32) = -1#   'cpLow
default_parm_values(33) = Format(0.00000005, "0.00E+0") 'cpTrans
default_parm_values(37) = "false"   'invStress
default_parm_values(39) = "false"   'measCP
'---------------------------------------------------
'Populate default values for all Tabs
'Tab #1
iEnd = -1
iStart = iEnd + 1
iEnd = iStart + 39
ICount = 0
For i = iStart To iEnd
 DUT1_ParmBox(i).Text = default_parm_values(ICount)
 ICount = ICount + 1
Next
'Tab #2
iStart = iEnd + 1
iEnd = iStart + 39
ICount = 0
For i = iStart To iEnd
 DUT1_ParmBox(i).Text = default_parm_values(ICount)
 ICount = ICount + 1
Next
'Tab #3
iStart = iEnd + 1
iEnd = iStart + 39
ICount = 0
For i = iStart To iEnd
 DUT1_ParmBox(i).Text = default_parm_values(ICount)
 ICount = ICount + 1
Next
'Tab #4
iStart = iEnd + 1
iEnd = iStart + 39
ICount = 0
For i = iStart To iEnd
 DUT1_ParmBox(i).Text = default_parm_values(ICount)
 ICount = ICount + 1
Next
'Tab #5
iStart = iEnd + 1
iEnd = iStart + 39
ICount = 0
For i = iStart To iEnd
 DUT1_ParmBox(i).Text = default_parm_values(ICount)
 ICount = ICount + 1
Next
'---------------------------------------------------
'disable text boxes for path to prevent user from typing characters
iStart = 35
iEnd = 40 * 5
For i = iStart To iEnd Step 40
  DUT1_ParmBox(i).Enabled = False
Next
'---------------------------------------------------
'By default, make controls visible only on DUT#1 Tab
Tab1MainFrame.Visible = True
Tab2MainFrame.Visible = False
Tab3MainFrame.Visible = False
Tab4MainFrame.Visible = False
Tab5MainFrame.Visible = False
'---------------------------------------------------
'Set tab names
SSTab1.TabCaption(0) = "DUT #1"
SSTab1.TabCaption(1) = "DUT #2"
SSTab1.TabCaption(2) = "DUT #3"
SSTab1.TabCaption(3) = "DUT #4"
SSTab1.TabCaption(4) = "DUT #5"
'---------------------------------------------------
'By default, set stress mode as individual device stress (max = 5 devices)
DUTStressOptionIndividual.value = True
DUTStressOptionIndividual.ForeColor = &HFF0000
DUTStressOptionCommonGate = False
CommnGateNumDevWGFMU.Enabled = False
'---------------------------------------------------
'By default, make only DUT #1 active by showing only one tab for DUT #1
NumDevWGFMU.Text = 1            'for stressing individual devices
CommnGateNumDevWGFMU = 2        'for common-gate devices
SSTab1.TabEnabled(0) = True
SSTab1.TabVisible(0) = True
For i = 1 To 4
    SSTab1.TabEnabled(i) = True
    SSTab1.TabVisible(i) = False
Next
'---------------------------------------------------
'Define parameter names for Tab #1
'> AC and DC
DUT1_Label(0).Caption = "stressTime="
DUT1_Label(1).Caption = "VGateStress="
DUT1_Label(2).Caption = "VDrainStress="
DUT1_Label(3).Caption = "relaxTime="
DUT1_Label(4).Caption = "VGateRelax="
DUT1_Label(5).Caption = "VDrainRelax="
DUT1_Label(6).Caption = "gateTransTime="
DUT1_Label(7).Caption = "drainTransTime="
'> AC Only
DUT1_Label(8).Caption = "VGateACLow="
DUT1_Label(9).Caption = "VDrainACLow="
DUT1_Label(10).Caption = "freq="
DUT1_Label(11).Caption = "dutyCycle="
'> Measurement - General
DUT1_Label(12).Caption = "VGateSense="
DUT1_Label(13).Caption = "VDrainSense="
DUT1_Label(14).Caption = "measIRange="
DUT1_Label(15).Caption = "initialSenseTime="
DUT1_Label(16).Caption = "measDelay="
DUT1_Label(17).Caption = "measPoints="
DUT1_Label(18).Caption = "startAvgPoint="
DUT1_Label(19).Caption = "stopAvgPoint="
DUT1_Label(20).Caption = "sampleInterval="
'> IV Sweep
DUT1_Label(21).Caption = "IVGateStart="
DUT1_Label(22).Caption = "IVGateStop="
DUT1_Label(23).Caption = "IVGateStep="
'> Log. Meas. Interval
DUT1_Label(24).Caption = "ppd="
'> Linear Meas. Interval
DUT1_Label(25).Caption = "stepTime="
'> Not used at the moment
DUT1_Label(26).Caption = "skew="
DUT1_Label(27).Caption = "measAfterHigh="
DUT1_Label(28).Caption = "cpFreqStart="
DUT1_Label(29).Caption = "cpFreqStep="
DUT1_Label(30).Caption = "cpNumSteps ="
DUT1_Label(31).Caption = "cpHigh="
DUT1_Label(32).Caption = "cpLow="
DUT1_Label(33).Caption = "cpTrans="
'> Hidden
DUT1_Label(34).Caption = "isLog="
DUT1_Label(35).Caption = "savePath="
DUT1_Label(36).Caption = "acStress="
DUT1_Label(37).Caption = "invStress="
DUT1_Label(38).Caption = "measIV="
DUT1_Label(39).Caption = "measCP="

'Store parameter labels in an array
Dim ParmName(39) As String
For i = 0 To 39
 ParmName(i) = DUT1_Label(i).Caption
Next
'Define parameter names for Tab #2
iEnd = 39
iStart = iEnd + 1
iEnd = iStart + 39
ICount = 0
For i = iStart To iEnd
 DUT1_Label(i).Caption = ParmName(ICount)
 ICount = ICount + 1
Next
'Define parameter names for Tab #3
iStart = iEnd + 1
iEnd = iStart + 39
ICount = 0
For i = iStart To iEnd
 DUT1_Label(i).Caption = ParmName(ICount)
 ICount = ICount + 1
 Next
'Define parameter names for Tab #4
iStart = iEnd + 1
iEnd = iStart + 39
ICount = 0
For i = iStart To iEnd
 DUT1_Label(i).Caption = ParmName(ICount)
 ICount = ICount + 1
Next
'Define parameter names for Tab #5
iStart = iEnd + 1
iEnd = iStart + 39
ICount = 0
For i = iStart To iEnd
 DUT1_Label(i).Caption = ParmName(ICount)
 ICount = ICount + 1
Next
'---------------------------------------------------
'Set aligment of label to be right-justified for Tab #1
'0 for left-justisfied
'1 for right-justified
'2 for center
'For Tab #1
iEnd = -1
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_Label(i).Alignment = 1
Next
'For Tab #2
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_Label(i).Alignment = 1
Next
'For Tab #3
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_Label(i).Alignment = 1
Next
'For Tab #4
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_Label(i).Alignment = 1
Next
'For Tab #5
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_Label(i).Alignment = 1
Next
'---------------------------------------------------
'Set width of parameter names for Tab #1
WidthDUT1_Label = 1600
iEnd = -1
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_Label(i).Width = WidthDUT1_Label
Next
'for Tab #2
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_Label(i).Width = WidthDUT1_Label
Next
'for Tab #3
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_Label(i).Width = WidthDUT1_Label
Next
'for Tab #4
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_Label(i).Width = WidthDUT1_Label
Next
'for Tab #5
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_Label(i).Width = WidthDUT1_Label
Next
'---------------------------------------------------
'Set the width of text box where input parameters are specified
'and the size of help buttonsfor Tab #1
Width_DUT_ParmBox = 800
iEnd = -1
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_ParmBox(i).Width = Width_DUT_ParmBox
    HelpMeParmDUT1(i).Caption = ""
    HelpMeParmDUT1(i).Width = 150
    HelpMeParmDUT1(i).Height = 150
Next
'for Tab #2
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_ParmBox(i).Width = Width_DUT_ParmBox
    HelpMeParmDUT1(i).Caption = ""
    HelpMeParmDUT1(i).Width = 150
    HelpMeParmDUT1(i).Height = 150
Next
'for Tab #3
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_ParmBox(i).Width = Width_DUT_ParmBox
    HelpMeParmDUT1(i).Caption = ""
    HelpMeParmDUT1(i).Width = 150
    HelpMeParmDUT1(i).Height = 150
Next
'for Tab #4
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_ParmBox(i).Width = Width_DUT_ParmBox
    HelpMeParmDUT1(i).Caption = ""
    HelpMeParmDUT1(i).Width = 150
    HelpMeParmDUT1(i).Height = 150
Next
'for Tab #5
iStart = iEnd + 1
iEnd = iStart + 39
For i = iStart To iEnd
    DUT1_ParmBox(i).Width = Width_DUT_ParmBox
    HelpMeParmDUT1(i).Caption = ""
    HelpMeParmDUT1(i).Width = 150
    HelpMeParmDUT1(i).Height = 150
Next
'---------------------------------------------------
'Adjust loations of DUT1_ParmBox
'------------------------
'For Tab #1
XLoc = 1750
InitY = 240
DelY = 340
HeightVal = 250
'Controls in Stress Option Frame
iStart = 40 * 0
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    DUT1_ParmBox(i).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    DUT1_ParmBox(i).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.25 * 240
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    'DUT1_ParmBox(I).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Top = InitY
    DUT1_ParmBox(i).Height = HeightVal
Next
'------------------------
'For Tab #2
XLoc = XLoc
InitY = 240
DelY = 340
HeightVal = 250
'Controls in Stress Option Frame
iStart = 40 * 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    DUT1_ParmBox(i).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    DUT1_ParmBox(i).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.25 * 240
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    'DUT1_ParmBox(I).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Top = InitY
    DUT1_ParmBox(i).Height = HeightVal
Next
'------------------------
'For Tab #3
XLoc = XLoc
InitY = 240
DelY = 340
HeightVal = 250
'Controls in Stress Option Frame
iStart = 40 * 2
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    DUT1_ParmBox(i).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    DUT1_ParmBox(i).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.25 * 240
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    'DUT1_ParmBox(I).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Top = InitY
    DUT1_ParmBox(i).Height = HeightVal
Next
'------------------------
'For Tab #4
XLoc = XLoc
InitY = 240
DelY = 340
HeightVal = 250
'Controls in Stress Option Frame
iStart = 40 * 3
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    DUT1_ParmBox(i).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    DUT1_ParmBox(i).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.25 * 240
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    'DUT1_ParmBox(I).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Top = InitY
    DUT1_ParmBox(i).Height = HeightVal
Next
'------------------------
'For Tab #5
XLoc = XLoc
InitY = 240
DelY = 340
HeightVal = 250
'Controls in Stress Option Frame
iStart = 40 * 4
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    DUT1_ParmBox(i).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    DUT1_ParmBox(i).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.25 * 240
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_ParmBox(i).Left = XLoc
    'DUT1_ParmBox(I).Top = InitY + (ICount * DelY)
    DUT1_ParmBox(i).Top = InitY
    DUT1_ParmBox(i).Height = HeightVal
Next
'---------------------------------------------------
'Adjust loations of parameter names
'------------------------
'For Tab #1
XLoc = 150
InitY = 250
DelY = 340
HeightVal = 250
'Controls in Stress Option Frame
iStart = 40 * 0
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    DUT1_Label(i).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    DUT1_Label(i).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.25 * 240
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    'DUT1_Label(I).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Top = InitY
    DUT1_Label(i).Height = HeightVal
Next
'------------------------
'For Tab #2
XLoc = XLoc
InitY = 250
DelY = 340
HeightVal = 250
'Controls in Stress Option Frame
iStart = 40 * 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    DUT1_Label(i).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    DUT1_Label(i).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.25 * 240
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    'DUT1_Label(I).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Top = InitY
    DUT1_Label(i).Height = HeightVal
Next
'------------------------
'For Tab #3
XLoc = XLoc
InitY = 250
DelY = 340
HeightVal = 250
'Controls in Stress Option Frame
iStart = 40 * 2
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    DUT1_Label(i).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    DUT1_Label(i).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.25 * 240
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    'DUT1_Label(I).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Top = InitY
    DUT1_Label(i).Height = HeightVal
Next
'------------------------
'For Tab #4
XLoc = XLoc
InitY = 250
DelY = 340
HeightVal = 250
'Controls in Stress Option Frame
iStart = 40 * 3
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    DUT1_Label(i).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    DUT1_Label(i).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.25 * 240
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    'DUT1_Label(I).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Top = InitY
    DUT1_Label(i).Height = HeightVal
Next
'------------------------
'For Tab #5
XLoc = XLoc
InitY = 250
DelY = 340
HeightVal = 250
'Controls in Stress Option Frame
iStart = 40 * 4
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    DUT1_Label(i).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    DUT1_Label(i).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.25 * 240
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    DUT1_Label(i).Left = XLoc
    'DUT1_Label(I).Top = InitY + (ICount * DelY)
    DUT1_Label(i).Top = InitY
    DUT1_Label(i).Height = HeightVal
Next
'---------------------------------------------------
'Adjust loations of parameter description butons
'------------------------
'For Tab #1
XLoc = 50
InitY = 250
DelY = 340
HeightVal = 190
'Controls in Stress Option Frame
iStart = 40 * 0
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    HelpMeParmDUT1(i).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    HelpMeParmDUT1(i).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.1 * 250
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    'HelpMeParmDUT1(I).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Top = InitY
    HelpMeParmDUT1(i).Height = HeightVal
Next
'------------------------
'For Tab #2
XLoc = 50
InitY = 250
DelY = 340
HeightVal = 190
'Controls in Stress Option Frame
iStart = 40 * 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    HelpMeParmDUT1(i).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    HelpMeParmDUT1(i).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.1 * 250
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    'HelpMeParmDUT1(I).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Top = InitY
    HelpMeParmDUT1(i).Height = HeightVal
Next
'------------------------
'For Tab #3
XLoc = 50
InitY = 250
DelY = 340
HeightVal = 190
'Controls in Stress Option Frame
iStart = 40 * 2
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    HelpMeParmDUT1(i).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    HelpMeParmDUT1(i).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.1 * 250
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    'HelpMeParmDUT1(I).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Top = InitY
    HelpMeParmDUT1(i).Height = HeightVal
Next
'------------------------
'For Tab #4
XLoc = 50
InitY = 250
DelY = 340
HeightVal = 190
'Controls in Stress Option Frame
iStart = 40 * 3
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    HelpMeParmDUT1(i).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    HelpMeParmDUT1(i).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.1 * 250
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    'HelpMeParmDUT1(I).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Top = InitY
    HelpMeParmDUT1(i).Height = HeightVal
Next
'------------------------
'For Tab #5
XLoc = 50
InitY = 250
DelY = 340
HeightVal = 190
'Controls in Stress Option Frame
iStart = 40 * 4
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    HelpMeParmDUT1(i).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Height = HeightVal
Next
'Controls in Measurement Option Frame
iStart = iEnd + 1
iEnd = iStart + 11
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    HelpMeParmDUT1(i).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Height = HeightVal
Next
'Controls in Measurement Time Interval Frame
'Make the two overlapped during run time (only one to be shown)
InitY = 3.1 * 250
iStart = iEnd + 1
iEnd = iStart + 1
ICount = -1
For i = iStart To iEnd
    ICount = ICount + 1
    HelpMeParmDUT1(i).Left = XLoc
    'HelpMeParmDUT1(I).Top = InitY + (ICount * DelY)
    HelpMeParmDUT1(i).Top = InitY
    HelpMeParmDUT1(i).Height = HeightVal
Next
'---------------------------------------------------
'Location of InitiIVSweep only option radio buttons
For i = 0 To 4
InitIVSweepOption(i).Top = 180
InitIVSweepOption(i).Left = 3000
Next
'---------------------------------------------------
End Sub
Private Sub HelpMeParmDUT1_Click(index As Integer)
'Display description of parameters
Dim IndexNumber As Integer
Dim NumIncrement As Integer
NumIncrement = 40
IndexNumber = index

Select Case IndexNumber
    Case 0, 40, 80, 120, 160
        MsgBoxAnswer = MsgBox("How long to stress (sec)", vbOKOnly, "")
        
    Case 1, 41, 81, 121, 161
        MsgBoxAnswer = MsgBox("Gate Stress Voltage (V)", vbOKOnly, "vGateStress")
     
    Case 2, 42, 82, 122, 162
        MsgBoxAnswer = MsgBox("Drain Stress voltage (V)", vbOKOnly, "vDrainStress")

    Case 3, 43, 83, 123, 163
        MsgBoxAnswer = MsgBox("How long to perform the relaxation part of the test (sec)", vbOKOnly, "relaxTime")
 
    Case 4, 44, 84, 124, 164
        MsgBoxAnswer = MsgBox("The gate voltage during relaxation (V)", vbOKOnly, "vGateRelax")

    Case 5, 45, 85, 125, 165
        MsgBoxAnswer = MsgBox("The drain voltage during relaxation (V)", vbOKOnly, "vDrainRelax")

    Case 6, 46, 86, 126, 166
        MsgBoxAnswer = MsgBox("Rising and falling edge of the gate pulse (sec)", vbOKOnly, "gateTransTime")

    Case 7, 47, 87, 127, 167
        MsgBoxAnswer = MsgBox("Rising and falling edge of the drain pulse (sec)", vbOKOnly, "drainTransTime")

    Case 8, 48, 88, 128, 168
        MsgBoxAnswer = MsgBox("The gate voltage (vGateStress) when the AC pulse is on the low cycle (V)", vbOKOnly, "vGateACLow")

    Case 9, 49, 89, 129, 169
        MsgBoxAnswer = MsgBox("The drain voltage (vDrainStress) when the AC pulse is on the low cycle (V)", vbOKOnly, "vDrainACLow")
    
    Case 10, 50, 90, 130, 170
        MsgBoxAnswer = MsgBox("The AC stress frequency (Hz)", vbOKOnly, "freq")
    
    Case 11, 51, 91, 131, 171
        MsgBoxAnswer = MsgBox("Duty cyle of the pulse in percent (%)", vbOKOnly, "dutyCycle")

    Case 12, 52, 92, 132, 172
        MsgBoxAnswer = MsgBox("Gate Sense Voltage (V)", vbOKOnly, "vGateSense")

    Case 13, 53, 93, 133, 173
        MsgBoxAnswer = MsgBox("Drain Sense voltage (V)", vbOKOnly, "")

    Case 14, 54, 94, 134, 174
        MsgBoxAnswer = MsgBox("The current measurement range to use on the drain (A)", vbOKOnly, "measIRange")

    Case 15, 55, 95, 135, 175
        MsgBoxAnswer = MsgBox("How long to wait before the first measurement point (sec)", vbOKOnly, "initialSenseTime")

    Case 16, 56, 96, 136, 176
        MsgBoxAnswer = MsgBox("Time to wait before each measure (sec)", vbOKOnly, "")
        
    Case 17, 57, 97, 137, 177
        MsgBoxAnswer = MsgBox("Number of points to measure (#)", vbOKOnly, "")
        
    Case 18, 58, 98, 138, 178
        MsgBoxAnswer = MsgBox("Start point to begin averaging (#)", vbOKOnly, "")
        
    Case 19, 59, 99, 139, 179
        MsgBoxAnswer = MsgBox("End point to end averaging  (#)", vbOKOnly, "")
        
    Case 20, 60, 100, 140, 180
        MsgBoxAnswer = MsgBox("Time between measurement points (sec)", vbOKOnly, "")
        
    Case 21, 61, 101, 141, 181
        MsgBoxAnswer = MsgBox("Gate start voltage for the sweep (V)", vbOKOnly, "ivGateStart")

    Case 22, 62, 102, 142, 182
        MsgBoxAnswer = MsgBox("Gate end voltage for the sweep (V)", vbOKOnly, "ivGateStop")
        
    Case 23, 63, 103, 143, 183
        MsgBoxAnswer = MsgBox("Gate step voltage (V)", vbOKOnly, "ivGateStep")
        
    Case 24, 64, 104, 144, 184
        MsgBoxAnswer = MsgBox("How many points to measure in a decade (# of points per decade)", vbOKOnly, "")
        
    Case 25, 65, 105, 145, 185
        MsgBoxAnswer = MsgBox("How often (every X secconds) to sample in linear measurement space (sec)", vbOKOnly, "stepTime")

    Case 26, 66, 106, 146, 186
        MsgBoxAnswer = MsgBox("Controls the skew of the drain relative to the gate symmetrically skews on both sides of the pulse (e.g., 0 = overlapping pulses, -10e-9 = drain transisiton 10ns before gate, 10e-9 = drain 10ns after gate)", vbOKOnly, "skew")

    Case 27, 67, 107, 147, 187
        MsgBoxAnswer = MsgBox("Measure the current after the high side of the pulse instead of the low (true or false)", vbOKOnly, "measAfterHigh")
        
    Case 28, 68, 108, 148, 188
        MsgBoxAnswer = MsgBox("The start frequency of the CP freq sweep (Hz)", vbOKOnly, "cpFreqStart")
    
    Case 29, 69, 109, 149, 189
        MsgBoxAnswer = MsgBox("The step frequency of the CP freq sweep (Hz)", vbOKOnly, "cpFreqStep")
    
    Case 30, 70, 110, 150, 190
        MsgBoxAnswer = MsgBox("The number of the CP freq steps to take (#)", vbOKOnly, "cpNumSteps")
        
    Case 31, 71, 111, 151, 191
        MsgBoxAnswer = MsgBox("The high voltage of the cp (V)", vbOKOnly, "cpHigh")
    
    Case 32, 72, 112, 152, 192
        MsgBoxAnswer = MsgBox("The low voltage of the cp (V)", vbOKOnly, "cpLow")
    
    Case 33, 73, 113, 153, 193
        MsgBoxAnswer = MsgBox("The transistion time of cp pulse (sec)", vbOKOnly, "cpTrans")
        
    Case 34, 74, 114, 154, 194
        MsgBoxAnswer = MsgBox("Measure log or linearly (true for log and false for linear)", vbOKOnly, "")
        
    Case 35, 75, 115, 155, 195
        MsgBoxAnswer = MsgBox("Where to save the data (= working folder)", vbOKOnly, "savePath")
        
    Case 36, 76, 116, 156, 196
        MsgBoxAnswer = MsgBox("Flag to perform AC stress or not (true for AC or false for DC)", vbOKOnly, "acStress")
        
    Case 37, 77, 117, 157, 197
        MsgBoxAnswer = MsgBox("Set Drain low when gate high (true or false)", vbOKOnly, "invStress")
        
    Case 38, 78, 118, 158, 198
        MsgBoxAnswer = MsgBox("Flag to enable IV measurement during the measurement part of the test (true or false)", vbOKOnly, "measIV")
        
    Case 39, 79, 119, 159, 199
        MsgBoxAnswer = MsgBox("Flag to tell the program to perform chargepumping during the sense phase of the test (true or false)", vbOKOnly, "measCP")

End Select

End Sub

Private Sub InitIVSweepOption_Click(index As Integer)
Dim IndexNumber As Integer
Dim NumParmBoxPerTab As Integer
NumParmBoxPerTab = 40
IndexNumber = index

'Measurement options for the selected tab
'Only "Which_Tab_am_I_On = 0" works during loading the form
'and fully works when user clicks a tab
Hidefrom = 21
NumHide = 3
If (Which_Tab_am_I_On = 0) Then 'Tab #1 (DUT #1)
 iStart = Hidefrom + (0 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
 HideOne = 12 + (0 * NumParmBoxPerTab)
ElseIf (Which_Tab_am_I_On = 1) Then  'Tab #2 (DUT #2)
 iStart = Hidefrom + (1 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
 HideOne = 12 + (1 * NumParmBoxPerTab)
ElseIf (Which_Tab_am_I_On = 2) Then 'Tab #3 (DUT #3)
 iStart = Hidefrom + (2 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
 HideOne = 12 + (2 * NumParmBoxPerTab)
ElseIf (Which_Tab_am_I_On = 3) Then 'Tab #4 (DUT #4)
 iStart = Hidefrom + (3 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
 HideOne = 12 + (3 * NumParmBoxPerTab)
ElseIf (Which_Tab_am_I_On = 4) Then 'Tab #5 (DUT #5)
 iStart = Hidefrom + (4 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
 HideOne = 12 + (4 * NumParmBoxPerTab)
End If

Select Case IndexNumber
    'IV measuremnt option selected
    Case 0, 1, 2, 3, 4
     For i = iStart To iEnd
      DUT1_Label(i).Visible = True
      DUT1_ParmBox(i).Visible = True
      HelpMeParmDUT1(i).Visible = True
      
      If IndexNumber = 0 Then
       DUT1_ParmBox(38).Text = "true"  'Turn on IVSweep
       'Set DC option and set stressTime=relaxTime=0
       StressOptionRadioButton(0).value = False
       StressOptionRadioButton(1).value = True
       MeasScaleRaioButton(0).value = False
       MeasScaleRaioButton(1).value = True
       DUT1_ParmBox(0).Text = 0
       DUT1_ParmBox(3).Text = 0
       DUT1_ParmBox(25).Text = 0
       DUT1_ParmBox(34).Text = "false"
       StressOptionRadioButton(0).Enabled = False
       
      ElseIf IndexNumber = 1 Then
       DUT1_ParmBox(78).Text = "true"  'Turn on IVSweep
       StressOptionRadioButton(2).value = False
       StressOptionRadioButton(3).value = True
       MeasScaleRaioButton(2).value = False
       MeasScaleRaioButton(3).value = True
       DUT1_ParmBox(40).Text = 0
       DUT1_ParmBox(43).Text = 0
       DUT1_ParmBox(65).Text = 0
       DUT1_ParmBox(74).Text = "false"
       StressOptionRadioButton(2).Enabled = False
      
      ElseIf IndexNumber = 2 Then
       DUT1_ParmBox(118).Text = "true"  'Turn on IVSweep
       StressOptionRadioButton(4).value = False
       StressOptionRadioButton(5).value = True
       MeasScaleRaioButton(4).value = False
       MeasScaleRaioButton(5).value = True
       DUT1_ParmBox(80).Text = 0
       DUT1_ParmBox(83).Text = 0
       DUT1_ParmBox(105).Text = 0
       DUT1_ParmBox(114).Text = "false"
       StressOptionRadioButton(4).Enabled = False
      
      ElseIf IndexNumber = 3 Then
       DUT1_ParmBox(158).Text = "true"  'Turn on IVSweep
       StressOptionRadioButton(6).value = False
       StressOptionRadioButton(7).value = True
       MeasScaleRaioButton(6).value = False
       MeasScaleRaioButton(7).value = True
       DUT1_ParmBox(120).Text = 0
       DUT1_ParmBox(123).Text = 0
       DUT1_ParmBox(145).Text = 0
       DUT1_ParmBox(154).Text = "false"
       StressOptionRadioButton(6).Enabled = False
      
      ElseIf IndexNumber = 4 Then
       DUT1_ParmBox(198).Text = "true"  'Turn on IVSweep
       StressOptionRadioButton(8).value = False
       StressOptionRadioButton(9).value = True
       MeasScaleRaioButton(8).value = False
       MeasScaleRaioButton(9).value = True
       DUT1_ParmBox(160).Text = 0
       DUT1_ParmBox(163).Text = 0
       DUT1_ParmBox(185).Text = 0
       DUT1_ParmBox(194).Text = "false"
       StressOptionRadioButton(8).Enabled = False
      End If

     Next
     DUT1_Label(HideOne).Visible = False
     DUT1_ParmBox(HideOne).Visible = False
     HelpMeParmDUT1(HideOne).Visible = False
     
End Select
End Sub

Private Sub MeasOptionRadioOption_Click(index As Integer)
Dim IndexNumber As Integer
Dim NumParmBoxPerTab As Integer
NumParmBoxPerTab = 40
IndexNumber = index

'Measurement options for the selected tab
'Only "Which_Tab_am_I_On = 0" works during loading the form
'and fully works when user clicks a tab
Hidefrom = 21
NumHide = 3
If (Which_Tab_am_I_On = 0) Then 'Tab #1 (DUT #1)
 iStart = Hidefrom + (0 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
 HideOne = 12 + (0 * NumParmBoxPerTab)
ElseIf (Which_Tab_am_I_On = 1) Then  'Tab #2 (DUT #2)
 iStart = Hidefrom + (1 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
 HideOne = 12 + (1 * NumParmBoxPerTab)
ElseIf (Which_Tab_am_I_On = 2) Then 'Tab #3 (DUT #3)
 iStart = Hidefrom + (2 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
 HideOne = 12 + (2 * NumParmBoxPerTab)
ElseIf (Which_Tab_am_I_On = 3) Then 'Tab #4 (DUT #4)
 iStart = Hidefrom + (3 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
 HideOne = 12 + (3 * NumParmBoxPerTab)
ElseIf (Which_Tab_am_I_On = 4) Then 'Tab #5 (DUT #5)
 iStart = Hidefrom + (4 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
 HideOne = 12 + (4 * NumParmBoxPerTab)
End If

Select Case IndexNumber
    'Spot measuremnt option selected
    Case 0, 2, 4, 6, 8
     If (IndexNumber = 0) Then 'Tab #1 (DUT #1)
        iStart = Hidefrom + (0 * NumParmBoxPerTab)
        iEnd = iStart + (NumHide - 1)
        HideOne = 12 + (0 * NumParmBoxPerTab)
     ElseIf (IndexNumber = 2) Then  'Tab #2 (DUT #2)
        iStart = Hidefrom + (1 * NumParmBoxPerTab)
        iEnd = iStart + (NumHide - 1)
        HideOne = 12 + (1 * NumParmBoxPerTab)
     ElseIf (IndexNumber = 4) Then 'Tab #3 (DUT #3)
        iStart = Hidefrom + (2 * NumParmBoxPerTab)
        iEnd = iStart + (NumHide - 1)
        HideOne = 12 + (2 * NumParmBoxPerTab)
     ElseIf (IndexNumber = 6) Then 'Tab #4 (DUT #4)
        iStart = Hidefrom + (3 * NumParmBoxPerTab)
        iEnd = iStart + (NumHide - 1)
        HideOne = 12 + (3 * NumParmBoxPerTab)
     ElseIf (IndexNumber = 8) Then 'Tab #5 (DUT #5)
        iStart = Hidefrom + (4 * NumParmBoxPerTab)
        iEnd = iStart + (NumHide - 1)
        HideOne = 12 + (4 * NumParmBoxPerTab)
     End If
     For i = iStart To iEnd
      DUT1_Label(i).Visible = False
      DUT1_ParmBox(i).Visible = False
      HelpMeParmDUT1(i).Visible = False
      
      If IndexNumber = 0 Then
       DUT1_ParmBox(38).Text = "false"  'Turn off IVSweep
       StressOptionRadioButton(0).Enabled = True
      ElseIf IndexNumber = 2 Then
       DUT1_ParmBox(78).Text = "false"  'Turn off IVSweep
       StressOptionRadioButton(2).Enabled = True
      ElseIf IndexNumber = 4 Then
       DUT1_ParmBox(118).Text = "false"  'Turn off IVSweep
       StressOptionRadioButton(4).Enabled = True
      ElseIf IndexNumber = 6 Then
       DUT1_ParmBox(158).Text = "false"  'Turn off IVSweep
       StressOptionRadioButton(6).Enabled = True
      ElseIf IndexNumber = 8 Then
       DUT1_ParmBox(198).Text = "false"  'Turn off IVSweep
       StressOptionRadioButton(8).Enabled = True
      End If
     
     Next
     DUT1_Label(HideOne).Visible = True
     DUT1_ParmBox(HideOne).Visible = True
     HelpMeParmDUT1(HideOne).Visible = True
     
    'IV measuremnt option selected
    Case 1, 3, 5, 7, 9
     For i = iStart To iEnd
      DUT1_Label(i).Visible = True
      DUT1_ParmBox(i).Visible = True
      HelpMeParmDUT1(i).Visible = True
      
      If IndexNumber = 1 Then
       DUT1_ParmBox(38).Text = "true"  'Turn on IVSweep
       StressOptionRadioButton(0).Enabled = True
      ElseIf IndexNumber = 3 Then
       DUT1_ParmBox(78).Text = "true"  'Turn on IVSweep
       StressOptionRadioButton(2).Enabled = True
      ElseIf IndexNumber = 5 Then
       DUT1_ParmBox(118).Text = "true"  'Turn on IVSweep
       StressOptionRadioButton(4).Enabled = True
      ElseIf IndexNumber = 7 Then
       DUT1_ParmBox(158).Text = "true"  'Turn on IVSweep
       StressOptionRadioButton(6).Enabled = True
      ElseIf IndexNumber = 9 Then
       DUT1_ParmBox(198).Text = "true"  'Turn on IVSweep
       StressOptionRadioButton(8).Enabled = True
      End If
 
     Next
     DUT1_Label(HideOne).Visible = False
     DUT1_ParmBox(HideOne).Visible = False
     HelpMeParmDUT1(HideOne).Visible = False
End Select
End Sub
Private Sub MeasScaleRaioButton_Click(index As Integer)
Dim IndexNumber As Integer
Dim NumParmBoxPerTab As Integer
NumParmBoxPerTab = 40
IndexNumber = index
'Measurement options for the selected tab
'Only "Which_Tab_am_I_On = 0" works during loading the form
'and fully works when user clicks a tab
Hidefrom = 25
NumHide = 1
If (Which_Tab_am_I_On = 0) Then 'Tab #1 (DUT #1)
 iStart = Hidefrom + (0 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
ElseIf (Which_Tab_am_I_On = 1) Then 'Tab #2 (DUT #2)
 iStart = Hidefrom + (1 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
ElseIf (Which_Tab_am_I_On = 2) Then 'Tab #3 (DUT #3)
 iStart = Hidefrom + (2 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
ElseIf (Which_Tab_am_I_On = 3) Then 'Tab #4 (DUT #4)
 iStart = Hidefrom + (3 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
ElseIf (Which_Tab_am_I_On = 4) Then 'Tab #5 (DUT #5)
 iStart = Hidefrom + (4 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
End If

Select Case IndexNumber
    'Log time scale option selected
    Case 0, 2, 4, 6, 8
     If (IndexNumber = 0) Then 'Tab #1 (DUT #1)
      iStart = Hidefrom + (0 * NumParmBoxPerTab)
      iEnd = iStart + (NumHide - 1)
     ElseIf (IndexNumber = 2) Then 'Tab #2 (DUT #2)
      iStart = Hidefrom + (1 * NumParmBoxPerTab)
      iEnd = iStart + (NumHide - 1)
     ElseIf (IndexNumber = 4) Then 'Tab #3 (DUT #3)
      iStart = Hidefrom + (2 * NumParmBoxPerTab)
      iEnd = iStart + (NumHide - 1)
     ElseIf (IndexNumber = 6) Then 'Tab #4 (DUT #4)
      iStart = Hidefrom + (3 * NumParmBoxPerTab)
      iEnd = iStart + (NumHide - 1)
     ElseIf (IndexNumber = 8) Then 'Tab #5 (DUT #5)
      iStart = Hidefrom + (4 * NumParmBoxPerTab)
      iEnd = iStart + (NumHide - 1)
     End If
     For i = iStart To iEnd
      DUT1_Label(i - 1).Visible = True
      DUT1_ParmBox(i - 1).Visible = True
      HelpMeParmDUT1(i - 1).Visible = True
      DUT1_Label(i).Visible = False
      DUT1_ParmBox(i).Visible = False
      HelpMeParmDUT1(i).Visible = False
      DUT1_ParmBox(34).Text = "true"  'Log flag is true
     Next
     
    'Linear time scale option selected
    Case 1, 3, 5, 7, 9
     For i = iStart To iEnd
      DUT1_Label(i - 1).Visible = False
      DUT1_ParmBox(i - 1).Visible = False
      HelpMeParmDUT1(i - 1).Visible = False
      DUT1_Label(i).Visible = True
      DUT1_ParmBox(i).Visible = True
      HelpMeParmDUT1(i).Visible = True
      DUT1_ParmBox(34).Text = "false"  'Log flag is false
     Next
     
End Select
End Sub

Private Sub NumDevWGFMU_Change()
'---------------------------------------------------
'If user makes this blank or place 0 then force the value to be 1
'so that Tab for DUT #1 is always displayed
MaxNumDevices = 5
If (Val(NumDevWGFMU.Text) = 0) Then
    NumDevWGFMU.Text = 1
End If
If (NumDevWGFMU.Text = "") Then
    NumDevWGFMU.Text = 1
End If
'---------------------------------------------------
'Display or hide tabs depending on user's choice on number of devices to stress
If (Val(NumDevWGFMU.Text) <= MaxNumDevices) Then
    For i = 0 To Val(NumDevWGFMU.Text) - 1
        SSTab1.TabVisible(i) = True
    Next
    For i = Val(NumDevWGFMU.Text) To MaxNumDevices - 1
        SSTab1.TabVisible(i) = False
    Next
Else
    MsgBox ("Maximum allowed individually-gated devices to stress is " & Str$(MaxNumDevices) & ". Define the value less than or equal to 5.")
    NumDevWGFMU.Text = 1
End If
'---------------------------------------------------
'Show message if chosen more than 3 devices
If (Val(NumDevWGFMU.Text) > 3) Then
  MsgBoxValue = MsgBox("WARNING: Only up to 10 WGFMU channels are supported for multi-device parallel stressing with which " & _
  "only up to 3 DUTs can be parallelly stressed with WGFMU channels connected to all device terminals. " & _
  "If more than 3 DUTS were to be stressed, some terminals of DUTs cannot use WGFMU channels. " & _
  "User must carefully define WGFMU channel connections to device terminals.")
End If
'---------------------------------------------------
End Sub
Private Sub RetrieveInputButton_Click()
'Show a dialog to specify to read input values from a file
Dim InputFileName As String
CommonDialogOpenFile.Filter = "(*.inp)|*.inp|All files (*.*)|*.*"
CommonDialogOpenFile.DefaultExt = "txt"
CommonDialogOpenFile.DialogTitle = "Select File"
CommonDialogOpenFile.ShowOpen
InputFileName = CommonDialogOpenFile.fileName
If InputFileName = "" Then
 Exit Sub
Else
'Clear all input text box values
iStart = 0
iEnd = 40 * 5 - 1
For i = iStart To iEnd
 DUT1_ParmBox(i) = ""
Next
DUT1_ParmBox(iStart + 35) = DataPathLabel.Caption
'Open file and read input values
Open InputFileName For Input As #1
 Input #1, StressType
 If StressType = 0 Then     'Individual DUT stresses
 Input #1, DummyNum
  NumDevWGFMU.Text = DummyNum
  For i = 0 To Val(NumDevWGFMU.Text) - 1
   SSTab1.TabVisible(i) = True
  Next
  For i = 0 To (40 * DummyNum - 1)
    Input #1, LineNum, InputVal
    DUT1_ParmBox(i).Text = InputVal
  Next
 ElseIf StressType = 1 Then 'Common-gate DUT stresses
 Input #1, DummyNum
  CommnGateNumDevWGFMU.Text = DummyNum
  For i = 0 To 39
    Input #1, LineNum, InputVal
    DUT1_ParmBox(i).Text = InputVal
  Next
 End If
 'Reset number format (only those that need sicientific format)
 Dim ValNeedingScieiticficFormat(8) As Integer
 ValNeedingScieiticficFormat(0) = 6
 ValNeedingScieiticficFormat(1) = 7
 ValNeedingScieiticficFormat(2) = 10
 ValNeedingScieiticficFormat(3) = 14
 ValNeedingScieiticficFormat(4) = 15
 ValNeedingScieiticficFormat(5) = 20
 ValNeedingScieiticficFormat(6) = 28
 ValNeedingScieiticficFormat(7) = 29
 ValNeedingScieiticficFormat(8) = 33
 For i = 0 To 8
 iStart = ValNeedingScieiticficFormat(i)
  For J = iStart To (40 * 5 - 1) Step 40
    DUT1_ParmBox(J).Text = Format(DUT1_ParmBox(J).Text, "0.00E+0")
  Next
 Next
End If
Close #1
'---------------------------------------------------------
'Set AC or DC
'DUT #1
If DUT1_ParmBox(36).Text = "true" Then  'AC flag is true
  StressOptionRadioButton(0).value = True
  StressOptionRadioButton(1).value = False
ElseIf DUT1_ParmBox(36).Text = "false" Then  'DC
  StressOptionRadioButton(0).value = False
  StressOptionRadioButton(1).value = True
End If
'DUT #2
If DUT1_ParmBox(76).Text = "true" Then  'AC flag is true
  StressOptionRadioButton(2).value = True
  StressOptionRadioButton(3).value = False
ElseIf DUT1_ParmBox(76).Text = "false" Then  'DC
  StressOptionRadioButton(2).value = False
  StressOptionRadioButton(3).value = True
End If
'DUT #3
If DUT1_ParmBox(116).Text = "true" Then  'AC flag is true
  StressOptionRadioButton(4).value = True
  StressOptionRadioButton(5).value = False
ElseIf DUT1_ParmBox(116).Text = "false" Then  'DC
  StressOptionRadioButton(4).value = False
  StressOptionRadioButton(5).value = True
End If
'DUT #4
If DUT1_ParmBox(156).Text = "true" Then  'AC flag is true
  StressOptionRadioButton(6).value = True
  StressOptionRadioButton(7).value = False
ElseIf DUT1_ParmBox(156).Text = "false" Then  'DC
  StressOptionRadioButton(6).value = False
  StressOptionRadioButton(7).value = True
End If
'DUT #5
If DUT1_ParmBox(196).Text = "true" Then  'AC flag is true
  StressOptionRadioButton(8).value = True
  StressOptionRadioButton(9).value = False
ElseIf DUT1_ParmBox(196).Text = "false" Then  'DC
  StressOptionRadioButton(8).value = False
  StressOptionRadioButton(9).value = True
End If
'Set Spot or IVSweep
'DUT #1
If DUT1_ParmBox(38).Text = "true" Then  'IV flag is true
  MeasOptionRadioOption(0).value = False
  MeasOptionRadioOption(1).value = True
ElseIf DUT1_ParmBox(38).Text = "false" Then  'Spot
  MeasOptionRadioOption(0).value = True
  MeasOptionRadioOption(1).value = False
End If
'DUT #2
If DUT1_ParmBox(78).Text = "true" Then  'Spot flag is true
  MeasOptionRadioOption(2).value = False
  MeasOptionRadioOption(3).value = True
ElseIf DUT1_ParmBox(78).Text = "false" Then  'Spot
  MeasOptionRadioOption(2).value = True
  MeasOptionRadioOption(3).value = False
End If
'DUT #3
If DUT1_ParmBox(118).Text = "true" Then  'Spot flag is true
  MeasOptionRadioOption(4).value = False
  MeasOptionRadioOption(5).value = True
ElseIf DUT1_ParmBox(118).Text = "false" Then  'Spot
  MeasOptionRadioOption(4).value = True
  MeasOptionRadioOption(5).value = False
End If
'DUT #4
If DUT1_ParmBox(158).Text = "true" Then  'Spot flag is true
  MeasOptionRadioOption(6).value = False
  MeasOptionRadioOption(7).value = True
ElseIf DUT1_ParmBox(158).Text = "false" Then  'Spot
  MeasOptionRadioOption(6).value = True
  MeasOptionRadioOption(7).value = False
End If
'DUT #5
If DUT1_ParmBox(198).Text = "true" Then  'Spot flag is true
  MeasOptionRadioOption(8).value = False
  MeasOptionRadioOption(9).value = True
ElseIf DUT1_ParmBox(198).Text = "false" Then  'Spot
  MeasOptionRadioOption(8).value = True
  MeasOptionRadioOption(9).value = False
End If
'Set Meas. time interval (Linear or Log)
'DUT #1
If DUT1_ParmBox(34).Text = "true" Then  'Log flag is true
  MeasScaleRaioButton(0).value = True
  MeasScaleRaioButton(1).value = False
ElseIf DUT1_ParmBox(34).Text = "false" Then  'Linear
  MeasScaleRaioButton(0).value = False
  MeasScaleRaioButton(1).value = True
End If
'DUT #2
If DUT1_ParmBox(74).Text = "true" Then  'Log flag is true
  MeasScaleRaioButton(2).value = True
  MeasScaleRaioButton(3).value = False
ElseIf DUT1_ParmBox(74).Text = "false" Then  'Linear
  MeasScaleRaioButton(2).value = False
  MeasScaleRaioButton(3).value = True
End If
'DUT #3
If DUT1_ParmBox(114).Text = "true" Then  'Log flag is true
  MeasScaleRaioButton(4).value = True
  MeasScaleRaioButton(5).value = False
ElseIf DUT1_ParmBox(114).Text = "false" Then  'Linear
  MeasScaleRaioButton(4).value = False
  MeasScaleRaioButton(5).value = True
End If
'DUT #4
If DUT1_ParmBox(154).Text = "true" Then  'Log flag is true
  MeasScaleRaioButton(6).value = True
  MeasScaleRaioButton(7).value = False
ElseIf DUT1_ParmBox(154).Text = "false" Then  'Linear
  MeasScaleRaioButton(6).value = False
  MeasScaleRaioButton(7).value = True
End If
'DUT #5
If DUT1_ParmBox(194).Text = "true" Then  'Log flag is true
  MeasScaleRaioButton(8).value = True
  MeasScaleRaioButton(9).value = False
ElseIf DUT1_ParmBox(194).Text = "false" Then  'Linear
  MeasScaleRaioButton(8).value = False
  MeasScaleRaioButton(9).value = True
End If
'---------------------------------------------------------
End Sub
Private Sub RunButton_Click()
Dim OutputFilenameEnd(4) As String
'---------------------------------------------------
'Revision #7
'---------------------------------------------------
'Make sure DC or AC
For i = 1 To 9 Step 2
 If (i = 1) Then
  IndexNum = 36
 ElseIf (i = 3) Then
  IndexNum = 76
 ElseIf (i = 5) Then
  IndexNum = 116
 ElseIf (i = 7) Then
  IndexNum = 156
 ElseIf (i = 9) Then
  IndexNum = 196
 End If
'
 If StressOptionRadioButton(i).value = True Then    'DC
  DUT1_ParmBox(IndexNum).Text = "false"  'Turn off AC flag = DC
 ElseIf StressOptionRadioButton(i).value = False Then
  DUT1_ParmBox(IndexNum).Text = "true"  'Turn on AC flag
 End If
Next
'---------------------------------------------------
'Make sure Spot or IVSweep
InitIVSweepIdent = -1
For i = 1 To 9 Step 2

InitIVSweepIdent = InitIVSweepIdent + 1
 If (i = 1) Then
  IndexNum = 38
 ElseIf (i = 3) Then
  IndexNum = 78
 ElseIf (i = 5) Then
  IndexNum = 118
 ElseIf (i = 7) Then
  IndexNum = 158
 ElseIf (i = 9) Then
  IndexNum = 198
 End If
'
 If InitIVSweepOption(InitIVSweepIdent).value = True Then
   DUT1_ParmBox(IndexNum).Text = "true"  'Turn on IVSweep
 Else
  If MeasOptionRadioOption(i).value = True Then  'IVSweep
   DUT1_ParmBox(IndexNum).Text = "true"  'Turn on IVSweep
  ElseIf MeasOptionRadioOption(i).value = False Then
   DUT1_ParmBox(IndexNum).Text = "false"  'Turn off IVSweep = Spot Measure
  End If
 End If

Next
'---------------------------------------------------
'Make sure meas time interval linear or log
For i = 1 To 9 Step 2
 If (i = 1) Then
  IndexNum = 34
 ElseIf (i = 3) Then
  IndexNum = 74
 ElseIf (i = 5) Then
  IndexNum = 114
 ElseIf (i = 7) Then
  IndexNum = 154
 ElseIf (i = 9) Then
  IndexNum = 194
 End If
'
 If MeasScaleRaioButton(i).value = True Then    'Linear
  'DUT1_ParmBox(IndexNum).Text = "true"  'Turn on Linear
  DUT1_ParmBox(IndexNum).Text = "false"
 ElseIf MeasScaleRaioButton(i).value = False Then
  'DUT1_ParmBox(IndexNum).Text = "false"  'Turn off Linear = Log
  DUT1_ParmBox(IndexNum).Text = "true"
 End If
Next
'---------------------------------------------------
'Number of DUTs to stress
Dim TotNUmDevicestoStress As String
TotNUmDevicestoStress = NumDevWGFMU.Text
'---------------------------------------------------
Dim Index_of_SavePath As Integer
Index_of_SavePath = 35
'---------------------------------------------------
'Prepare strings to pass to external program (e.g., WGFMU_NTI.exe)
'Tab #1
iEnd = -1
iStart = iEnd + 1
iEnd = iStart + 39
SkipIndex = 0 * 40 + Index_of_SavePath
Channel1 = Channel_1_DUT1
Channel2 = Channel_2_DUT1
Channel3 = Channel_3_DUT1
Channel4 = Channel_4_DUT1
HiddenInputStringTMP$ = ""
If InitIVSweepOption(0).value = False Then
 OutputFilenameEnd(0) = "_DUT1_data.dat"
ElseIf InitIVSweepOption(0).value = True Then
 OutputFilenameEnd(0) = "_DUT1_InitIVSweep_data.dat"
End If
HiddenInputStringTMP$ = " -save1:" & DataPathLabel.Caption & "\Fast_BTI_WGFMU_" & LotIDTextBox.Text & "_" & WaferIDTexBox.Text & "_Chip" & ChipIDLabel.Text & "_X" & XCoord.Text & "_Y" & YCoord.Text & OutputFilenameEnd(0) & " ~TotNumDUTs=" & TotNUmDevicestoStress & "~DUTNum1=1" & " ~DUT1_channel1=" & Channel1 & "~DUT1_channel2=" & Channel2 & "~DUT1_channel3=" & Channel3 & "~DUT1_channel4=" & Channel4
For i = iStart To iEnd
    If i <> SkipIndex Then
     EqualPos = InStr(DUT1_Label(i), "=")
     TmpString = Left(DUT1_Label(i), EqualPos - 1) & "1="
     'TmpString = Left(DUT1_Label(I), EqualPos - 1) & "="
      HiddenInputStringTMP$ = HiddenInputStringTMP$ & "~" & TmpString & DUT1_ParmBox(i).Text
    End If
Next
HiddenInputString(0).Text = HiddenInputStringTMP$
'Tab #2
iStart = iEnd + 1
iEnd = iStart + 39
SkipIndex = 1 * 40 + Index_of_SavePath
Channel1 = Channel_1_DUT2
Channel2 = Channel_2_DUT2
Channel3 = Channel_3_DUT2
Channel4 = Channel_4_DUT2
HiddenInputStringTMP$ = ""
If InitIVSweepOption(1).value = False Then
 OutputFilenameEnd(1) = "_DUT2_data.dat"
ElseIf InitIVSweepOption(1).value = True Then
 OutputFilenameEnd(1) = "_DUT2_InitIVSweep_data.dat"
End If
HiddenInputStringTMP$ = "~-save2:" & DataPathLabel.Caption & "\Fast_BTI_WGFMU_" & LotIDTextBox.Text & "_" & WaferIDTexBox.Text & "_Chip" & ChipIDLabel.Text & "_X" & XCoord.Text & "_Y" & YCoord.Text & OutputFilenameEnd(1) & "~DUTNum2=2" & " ~DUT2_channel1=" & Channel1 & "~DUT2_channel2=" & Channel2 & "~DUT2_channel3=" & Channel3 & "~DUT2_channel4=" & Channel4
For i = iStart To iEnd
    If i <> SkipIndex Then
     EqualPos = InStr(DUT1_Label(i), "=")
     TmpString = Left(DUT1_Label(i), EqualPos - 1) & "2="
     HiddenInputStringTMP$ = HiddenInputStringTMP$ & "~" & TmpString & DUT1_ParmBox(i).Text
    End If
Next
'MsgBox (HiddenInputStringTMP$)
HiddenInputString(1).Text = HiddenInputStringTMP$
'Tab #3
iStart = iEnd + 1
iEnd = iStart + 39
SkipIndex = 2 * 40 + Index_of_SavePath
Channel1 = Channel_1_DUT3
Channel2 = Channel_2_DUT3
Channel3 = Channel_3_DUT3
Channel4 = Channel_4_DUT3
HiddenInputStringTMP$ = ""
If InitIVSweepOption(2).value = False Then
 OutputFilenameEnd(2) = "_DUT3_data.dat"
ElseIf InitIVSweepOption(2).value = True Then
 OutputFilenameEnd(2) = "_DUT3_InitIVSweep_data.dat"
End If
HiddenInputStringTMP$ = "~-save3:" & DataPathLabel.Caption & "\Fast_BTI_WGFMU_" & LotIDTextBox.Text & "_" & WaferIDTexBox.Text & "_Chip" & ChipIDLabel.Text & "_X" & XCoord.Text & "_Y" & YCoord.Text & OutputFilenameEnd(2) & "~DUTNum3=3" & " ~DUT3_channel1=" & Channel1 & "~DUT3_channel2=" & Channel2 & "~DUT3_channel3=" & Channel3 & "~DUT3_channel4=" & Channel4
For i = iStart To iEnd
    If i <> SkipIndex Then
     EqualPos = InStr(DUT1_Label(i), "=")
     TmpString = Left(DUT1_Label(i), EqualPos - 1) & "3="
     HiddenInputStringTMP$ = HiddenInputStringTMP$ & "~" & TmpString & DUT1_ParmBox(i).Text
    End If
Next
'MsgBox (HiddenInputStringTMP$)
HiddenInputString(2).Text = HiddenInputStringTMP$
'Tab #4
iStart = iEnd + 1
iEnd = iStart + 39
SkipIndex = 3 * 40 + Index_of_SavePath
Channel1 = Channel_1_DUT4
Channel2 = Channel_2_DUT4
Channel3 = Channel_3_DUT4
Channel4 = Channel_4_DUT4
HiddenInputStringTMP$ = ""
If InitIVSweepOption(3).value = False Then
 OutputFilenameEnd(3) = "_DUT4_data.dat"
ElseIf InitIVSweepOption(3).value = True Then
 OutputFilenameEnd(3) = "_DUT4_InitIVSweep_data.dat"
End If
HiddenInputStringTMP$ = "~-save4:" & DataPathLabel.Caption & "\Fast_BTI_WGFMU_" & LotIDTextBox.Text & "_" & WaferIDTexBox.Text & "_Chip" & ChipIDLabel.Text & "_X" & XCoord.Text & "_Y" & YCoord.Text & OutputFilenameEnd(3) & "~DUTNum4=4" & " ~DUT4_channel1=" & Channel1 & "~DUT4_channel2=" & Channel2 & "~DUT4_channel3=" & Channel3 & "~DUT4_channel4=" & Channel4
For i = iStart To iEnd
    If i <> SkipIndex Then
     EqualPos = InStr(DUT1_Label(i), "=")
     TmpString = Left(DUT1_Label(i), EqualPos - 1) & "4="
     HiddenInputStringTMP$ = HiddenInputStringTMP$ & "~" & TmpString & DUT1_ParmBox(i).Text
    End If
Next
'MsgBox (HiddenInputStringTMP$)
HiddenInputString(3).Text = HiddenInputStringTMP$
'Tab #5
iStart = iEnd + 1
iEnd = iStart + 39
SkipIndex = 4 * 40 + Index_of_SavePath
Channel1 = Channel_1_DUT5
Channel2 = Channel_2_DUT5
Channel3 = Channel_3_DUT5
Channel4 = Channel_4_DUT5
HiddenInputStringTMP$ = ""
If InitIVSweepOption(4).value = False Then
 OutputFilenameEnd(4) = "_DUT5_data.dat"
ElseIf InitIVSweepOption(4).value = True Then
 OutputFilenameEnd(4) = "_DUT5_InitIVSweep_data.dat"
End If
HiddenInputStringTMP$ = "~-save5:" & DataPathLabel.Caption & "\Fast_BTI_WGFMU_" & LotIDTextBox.Text & "_" & WaferIDTexBox.Text & "_Chip" & ChipIDLabel.Text & "_X" & XCoord.Text & "_Y" & YCoord.Text & OutputFilenameEnd(4) & "~DUTNum5=5" & " ~DUT5_channel1=" & Channel1 & "~DUT5_channel2=" & Channel2 & "~DUT5_channel3=" & Channel3 & "~DUT5_channel4=" & Channel4
For i = iStart To iEnd
    If i <> SkipIndex Then
     EqualPos = InStr(DUT1_Label(i), "=")
     TmpString = Left(DUT1_Label(i), EqualPos - 1) & "5="
     HiddenInputStringTMP$ = HiddenInputStringTMP$ & "~" & TmpString & DUT1_ParmBox(i).Text
    End If
Next
'MsgBox (HiddenInputStringTMP$)
HiddenInputString(4).Text = HiddenInputStringTMP$
'---------------------------------------------------------
'Store strings to pass to C# program for fast BTI stress with WGFMU
Dim strShellCommand As String
strShellCommand = WhereaboutEXE.Caption
For i = 0 To (Val(TotNUmDevicestoStress) - 1)
 strShellCommand = strShellCommand & HiddenInputString(i).Text
Next
'---------------------------------------------------------
'Output C# argument to a file for Debugging purposes
TestFileName = "C:\Check_String.txt"
Open TestFileName For Output As #1
 Print #1, strShellCommand
Close #1
'Exit Sub
'------------------------------------------------------
'Excecution of external progmam (e.g. C# WGFMU_BTI.exe)
 Shell strShellCommand, vbNormalFocus
'------------------------------------------------------
End Sub
Private Sub SaveInputAsButton_Click()
'Max=39 parameters are defined in this program
'Only index=0 to 25, 34, 36 and 38 are fed to C# code
'Parameters of which index are from 26 to 33 are forCP (charge Pumping)
'that C# doesn't have proper implementation at the moment
'Tab #1
iEnd = -1
iStart = iEnd + 1
iEnd = iStart + 39
HiddenInputStringTMP$ = ""
For i = iStart To iEnd
    'If (0 <= I <= 25 Or I = 34 Or I = 36 Or I = 38) Then
     HiddenInputStringTMP$ = HiddenInputStringTMP$ & "~" & DUT1_Label(i) & DUT1_ParmBox(i).Text
    'End If
Next
HiddenInputString(0).Text = HiddenInputStringTMP$
'Tab #2
iStart = iEnd + 1
iEnd = iStart + 39
HiddenInputStringTMP$ = ""
For i = iStart To iEnd
    'If (0 <= I <= 25 Or I = 34 Or I = 36 Or I = 38) Then
     HiddenInputStringTMP$ = HiddenInputStringTMP$ & "~" & DUT1_Label(i) & DUT1_ParmBox(i).Text
    'End If
Next
HiddenInputString(1).Text = HiddenInputStringTMP$
'Tab #3
iStart = iEnd + 1
iEnd = iStart + 39
HiddenInputStringTMP$ = ""
For i = iStart To iEnd
    'If (0 <= I <= 25 Or I = 34 Or I = 36 Or I = 38) Then
     HiddenInputStringTMP$ = HiddenInputStringTMP$ & "~" & DUT1_Label(i) & DUT1_ParmBox(i).Text
    'End If
Next
HiddenInputString(2).Text = HiddenInputStringTMP$
'Tab #4
iStart = iEnd + 1
iEnd = iStart + 39
HiddenInputStringTMP$ = ""
For i = iStart To iEnd
    'If (0 <= I <= 25 Or I = 34 Or I = 36 Or I = 38) Then
     HiddenInputStringTMP$ = HiddenInputStringTMP$ & "~" & DUT1_Label(i) & DUT1_ParmBox(i).Text
    'End If
Next
HiddenInputString(3).Text = HiddenInputStringTMP$
'Tab #5
iStart = iEnd + 1
iEnd = iStart + 39
HiddenInputStringTMP$ = ""
For i = iStart To iEnd
    'If (0 <= I <= 25 Or I = 34 Or I = 36 Or I = 38) Then
     HiddenInputStringTMP$ = HiddenInputStringTMP$ & "~" & DUT1_Label(i) & DUT1_ParmBox(i).Text
    'End If
Next
HiddenInputString(4).Text = HiddenInputStringTMP$
'-----------------------------------------------------------
' save all input values to a file
'-----------------------------------------------------------
' Set CancelError is true
' If user presses the cancel button,
' Common Dialog Control will generate
' a runtime error that can be caught

 CommonDialogSaveInput.CancelError = True
 On Error GoTo ErrHandler

 'Specify default filter to *.inp
 CommonDialogSaveInput.Filter = "Input File (*.inp)|*.inp|"
 ' Display the SaveAs dialog box, and save the selected file in the variable strFileName
 CommonDialogSaveInput.ShowSave
 strFileName = CommonDialogSaveInput.fileName
' Check if there is an existing file with the same name
If Dir(strFileName, vbNormal) <> "" Then
   QuestionString$ = "The file " & strFileName & " already exists. Do you want to replace it?"
   return_ = MsgBox(QuestionString$, vbQuestion + vbYesNo, "File overwrite protection")
   'return_=6 if YES selected and return_=7 if No selected
 '---
 If return_ = 6 Then  'Yes is selected. Proceed to overwrite file
   'Overwrite the file and exit this subroutine
   Open strFileName For Output As #1
   If DUTStressOptionIndividual.value = True Then
    Print #1, 0       '0 is an identifier for individual dut stress
    Print #1, Val(NumDevWGFMU)
    For i = 0 To (40 * Val(NumDevWGFMU)) - 1
     Print #1, i, ",", DUT1_ParmBox(i)
    Next
   ElseIf DUTStressOptionCommonGate.value = True Then
    Print #1, 1       '1 is an identifier for common-gate stress
    Print #1, Val(CommnGateNumDevWGFMU)
    For i = 0 To 39
     Print #1, i, ",", DUT1_ParmBox(i)
    Next
   End If
   Close #1
   Exit Sub
 ElseIf return_ = 7 Then      'No is selected
    'Do nothing and go back to the main program
 End If
 '---
Else
  Open strFileName For Output As #1
  If DUTStressOptionIndividual.value = True Then
   Print #1, 0       '0 is an identifier for individual dut stress
   Print #1, Val(NumDevWGFMU)
   For i = 0 To (40 * Val(NumDevWGFMU)) - 1
    Print #1, i, ",", DUT1_ParmBox(i)
   Next
  ElseIf DUTStressOptionCommonGate.value = True Then
   Print #1, 1       '1 is an identifier for common-gate stress
   Print #1, Val(CommnGateNumDevWGFMU)
   For i = 0 To 39
    Print #1, i, ",", DUT1_ParmBox(i)
   Next
  End If
  Close #1
  Exit Sub
End If

ErrHandler:
 'User pressed the Cancel button
End Sub

Private Sub SpecifyEXE_Click()
'Show a dialog to specify to read input values from a file
Dim exeFileName As String
CommonDialogOpenFile.Filter = "(*.exe)|*.exe"
CommonDialogOpenFile.DefaultExt = "exe"
CommonDialogOpenFile.DialogTitle = "Select external WGFMU stress program"
CommonDialogOpenFile.ShowOpen
exeFileName = CommonDialogOpenFile.fileName
If exeFileName = "" Then
  WhereaboutEXE.Caption = "Please, choose a program to run stress"
  Exit Sub
Else
  WhereaboutEXE.Caption = exeFileName
End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
Dim TabIndexNumber As Integer
TabIndexNumber = SSTab1.Tab
Which_Tab_am_I_On = TabIndexNumber
'MsgBox ("Tab# selected = " & TabIndexNumber) '--> only for debugging purposes
Dim Num1 As Integer
Dim Num2 As Integer

Select Case TabIndexNumber
Case 0
'Make controls visible only on DUT#1 Tab
 Tab1MainFrame.Visible = True
 Tab2MainFrame.Visible = False
 Tab3MainFrame.Visible = False
 Tab4MainFrame.Visible = False
 Tab5MainFrame.Visible = False
 '
 Num1 = 0
 Num2 = 1
 If StressOptionRadioButton(Num1).value = True And StressOptionRadioButton(Num2).value = False Then
  StressOptionRadioButton(Num1).value = False
  StressOptionRadioButton(Num2).value = True
  StressOptionRadioButton(Num1).value = True
  StressOptionRadioButton(Num2).value = False
 End If
 If StressOptionRadioButton(Num1).value = False And StressOptionRadioButton(Num2).value = True Then
  StressOptionRadioButton(Num1).value = True
  StressOptionRadioButton(Num2).value = False
  StressOptionRadioButton(Num1).value = False
  StressOptionRadioButton(Num2).value = True
 End If
 '
 If MeasOptionRadioOption(Num1).value = True And MeasOptionRadioOption(Num2).value = False Then
  MeasOptionRadioOption(Num1).value = False
  MeasOptionRadioOption(Num2).value = True
  MeasOptionRadioOption(Num1).value = True
  MeasOptionRadioOption(Num2).value = False
 End If
 If MeasOptionRadioOption(Num1).value = False And MeasOptionRadioOption(Num2).value = True Then
  MeasOptionRadioOption(Num1).value = True
  MeasOptionRadioOption(Num2).value = False
  MeasOptionRadioOption(Num1).value = False
  MeasOptionRadioOption(Num2).value = True
 End If
 '
 If MeasScaleRaioButton(Num1).value = True And MeasScaleRaioButton(Num2).value = False Then
  MeasScaleRaioButton(Num1).value = False
  MeasScaleRaioButton(Num2).value = True
  MeasScaleRaioButton(Num1).value = True
  MeasScaleRaioButton(Num2).value = False
 End If
 If MeasScaleRaioButton(Num1).value = False And MeasScaleRaioButton(Num2).value = True Then
  MeasScaleRaioButton(Num1).value = True
  MeasScaleRaioButton(Num2).value = False
  MeasScaleRaioButton(Num1).value = False
  MeasScaleRaioButton(Num2).value = True
 End If
 
Case 1
'Make controls visible only on DUT#2 Tab
 Tab1MainFrame.Visible = False
 Tab2MainFrame.Visible = True
 Tab3MainFrame.Visible = False
 Tab4MainFrame.Visible = False
 Tab5MainFrame.Visible = False
 Num1 = 2
 Num2 = 3
 If StressOptionRadioButton(Num1).value = True And StressOptionRadioButton(Num2).value = False Then
  StressOptionRadioButton(Num1).value = False
  StressOptionRadioButton(Num2).value = True
  StressOptionRadioButton(Num1).value = True
  StressOptionRadioButton(Num2).value = False
 End If
 If StressOptionRadioButton(Num1).value = False And StressOptionRadioButton(Num2).value = True Then
  StressOptionRadioButton(Num1).value = True
  StressOptionRadioButton(Num2).value = False
  StressOptionRadioButton(Num1).value = False
  StressOptionRadioButton(Num2).value = True
 End If
 '
 If MeasOptionRadioOption(Num1).value = True And MeasOptionRadioOption(Num2).value = False Then
  MeasOptionRadioOption(Num1).value = False
  MeasOptionRadioOption(Num2).value = True
  MeasOptionRadioOption(Num1).value = True
  MeasOptionRadioOption(Num2).value = False
 End If
 If MeasOptionRadioOption(Num1).value = False And MeasOptionRadioOption(Num2).value = True Then
  MeasOptionRadioOption(Num1).value = True
  MeasOptionRadioOption(Num2).value = False
  MeasOptionRadioOption(Num1).value = False
  MeasOptionRadioOption(Num2).value = True
 End If
 '
 If MeasScaleRaioButton(Num1).value = True And MeasScaleRaioButton(Num2).value = False Then
  MeasScaleRaioButton(Num1).value = False
  MeasScaleRaioButton(Num2).value = True
  MeasScaleRaioButton(Num1).value = True
  MeasScaleRaioButton(Num2).value = False
 End If
 If MeasScaleRaioButton(Num1).value = False And MeasScaleRaioButton(Num2).value = True Then
  MeasScaleRaioButton(Num1).value = True
  MeasScaleRaioButton(Num2).value = False
  MeasScaleRaioButton(Num1).value = False
  MeasScaleRaioButton(Num2).value = True
 End If
 
Case 2
'Make controls visible only on DUT#3 Tab
 Tab1MainFrame.Visible = False
 Tab2MainFrame.Visible = False
 Tab3MainFrame.Visible = True
 Tab4MainFrame.Visible = False
 Tab5MainFrame.Visible = False
  Num1 = 4
  Num2 = 5
 If StressOptionRadioButton(Num1).value = True And StressOptionRadioButton(Num2).value = False Then
  StressOptionRadioButton(Num1).value = False
  StressOptionRadioButton(Num2).value = True
  StressOptionRadioButton(Num1).value = True
  StressOptionRadioButton(Num2).value = False
 End If
 If StressOptionRadioButton(Num1).value = False And StressOptionRadioButton(Num2).value = True Then
  StressOptionRadioButton(Num1).value = True
  StressOptionRadioButton(Num2).value = False
  StressOptionRadioButton(Num1).value = False
  StressOptionRadioButton(Num2).value = True
 End If
 '
 If MeasOptionRadioOption(Num1).value = True And MeasOptionRadioOption(Num2).value = False Then
  MeasOptionRadioOption(Num1).value = False
  MeasOptionRadioOption(Num2).value = True
  MeasOptionRadioOption(Num1).value = True
  MeasOptionRadioOption(Num2).value = False
 End If
 If MeasOptionRadioOption(Num1).value = False And MeasOptionRadioOption(Num2).value = True Then
  MeasOptionRadioOption(Num1).value = True
  MeasOptionRadioOption(Num2).value = False
  MeasOptionRadioOption(Num1).value = False
  MeasOptionRadioOption(Num2).value = True
 End If
 '
 If MeasScaleRaioButton(Num1).value = True And MeasScaleRaioButton(Num2).value = False Then
  MeasScaleRaioButton(Num1).value = False
  MeasScaleRaioButton(Num2).value = True
  MeasScaleRaioButton(Num1).value = True
  MeasScaleRaioButton(Num2).value = False
 End If
 If MeasScaleRaioButton(Num1).value = False And MeasScaleRaioButton(Num2).value = True Then
  MeasScaleRaioButton(Num1).value = True
  MeasScaleRaioButton(Num2).value = False
  MeasScaleRaioButton(Num1).value = False
  MeasScaleRaioButton(Num2).value = True
 End If
 
Case 3
'Make controls visible only on DUT#4 Tab
 Tab1MainFrame.Visible = False
 Tab2MainFrame.Visible = False
 Tab3MainFrame.Visible = False
 Tab4MainFrame.Visible = True
 Tab5MainFrame.Visible = False
  Num1 = 6
  Num2 = 7
 If StressOptionRadioButton(Num1).value = True And StressOptionRadioButton(Num2).value = False Then
  StressOptionRadioButton(Num1).value = False
  StressOptionRadioButton(Num2).value = True
  StressOptionRadioButton(Num1).value = True
  StressOptionRadioButton(Num2).value = False
 End If
 If StressOptionRadioButton(Num1).value = False And StressOptionRadioButton(Num2).value = True Then
  StressOptionRadioButton(Num1).value = True
  StressOptionRadioButton(Num2).value = False
  StressOptionRadioButton(Num1).value = False
  StressOptionRadioButton(Num2).value = True
 End If
 '
 If MeasOptionRadioOption(Num1).value = True And MeasOptionRadioOption(Num2).value = False Then
  MeasOptionRadioOption(Num1).value = False
  MeasOptionRadioOption(Num2).value = True
  MeasOptionRadioOption(Num1).value = True
  MeasOptionRadioOption(Num2).value = False
 End If
 If MeasOptionRadioOption(Num1).value = False And MeasOptionRadioOption(Num2).value = True Then
  MeasOptionRadioOption(Num1).value = True
  MeasOptionRadioOption(Num2).value = False
  MeasOptionRadioOption(Num1).value = False
  MeasOptionRadioOption(Num2).value = True
 End If
 '
 If MeasScaleRaioButton(Num1).value = True And MeasScaleRaioButton(Num2).value = False Then
  MeasScaleRaioButton(Num1).value = False
  MeasScaleRaioButton(Num2).value = True
  MeasScaleRaioButton(Num1).value = True
  MeasScaleRaioButton(Num2).value = False
 End If
 If MeasScaleRaioButton(Num1).value = False And MeasScaleRaioButton(Num2).value = True Then
  MeasScaleRaioButton(Num1).value = True
  MeasScaleRaioButton(Num2).value = False
  MeasScaleRaioButton(Num1).value = False
  MeasScaleRaioButton(Num2).value = True
 End If
 
Case 4
'Make controls visible only on DUT#5 Tab
 Tab1MainFrame.Visible = False
 Tab2MainFrame.Visible = False
 Tab3MainFrame.Visible = False
 Tab4MainFrame.Visible = False
 Tab5MainFrame.Visible = True
  Num1 = 8
  Num2 = 9
 If StressOptionRadioButton(Num1).value = True And StressOptionRadioButton(Num2).value = False Then
  StressOptionRadioButton(Num1).value = False
  StressOptionRadioButton(Num2).value = True
  StressOptionRadioButton(Num1).value = True
  StressOptionRadioButton(Num2).value = False
 End If
 If StressOptionRadioButton(Num1).value = False And StressOptionRadioButton(Num2).value = True Then
  StressOptionRadioButton(Num1).value = True
  StressOptionRadioButton(Num2).value = False
  StressOptionRadioButton(Num1).value = False
  StressOptionRadioButton(Num2).value = True
 End If
 '
 If MeasOptionRadioOption(Num1).value = True And MeasOptionRadioOption(Num2).value = False Then
  MeasOptionRadioOption(Num1).value = False
  MeasOptionRadioOption(Num2).value = True
  MeasOptionRadioOption(Num1).value = True
  MeasOptionRadioOption(Num2).value = False
 End If
 If MeasOptionRadioOption(Num1).value = False And MeasOptionRadioOption(Num2).value = True Then
  MeasOptionRadioOption(Num1).value = True
  MeasOptionRadioOption(Num2).value = False
  MeasOptionRadioOption(Num1).value = False
  MeasOptionRadioOption(Num2).value = True
 End If
 '
 If MeasScaleRaioButton(Num1).value = True And MeasScaleRaioButton(Num2).value = False Then
  MeasScaleRaioButton(Num1).value = False
  MeasScaleRaioButton(Num2).value = True
  MeasScaleRaioButton(Num1).value = True
  MeasScaleRaioButton(Num2).value = False
 End If
 If MeasScaleRaioButton(Num1).value = False And MeasScaleRaioButton(Num2).value = True Then
  MeasScaleRaioButton(Num1).value = True
  MeasScaleRaioButton(Num2).value = False
  MeasScaleRaioButton(Num1).value = False
  MeasScaleRaioButton(Num2).value = True
 End If
 
End Select

End Sub

Private Sub StressOptionRadioButton_Click(index As Integer)
Dim IndexNumber As Integer
Dim NumParmBoxPerTab As Integer
NumParmBoxPerTab = 40
IndexNumber = index
Hidefrom = 8
NumHide = 4
'AC parameters for the selected tab
'Only "Which_Tab_am_I_On = 0" works during loading the form
'and fully works when user clicks a tab
If (Which_Tab_am_I_On = 0) Then 'Tab #1 (DUT #1)
 iStart = Hidefrom + (0 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
ElseIf (Which_Tab_am_I_On = 1) Then 'Tab #2 (DUT #2)
 iStart = Hidefrom + (1 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
ElseIf (Which_Tab_am_I_On = 2) Then 'Tab #3 (DUT #3)
 iStart = Hidefrom + (2 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
ElseIf (Which_Tab_am_I_On = 3) Then 'Tab #4 (DUT #4)
 iStart = Hidefrom + (3 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
ElseIf (Which_Tab_am_I_On = 4) Then 'Tab #5 (DUT #5)
 iStart = Hidefrom + (4 * NumParmBoxPerTab)
 iEnd = iStart + (NumHide - 1)
End If

Select Case IndexNumber
    'AC selected
    Case 0, 2, 4, 6, 8
     For i = iStart To iEnd
      DUT1_Label(i).Visible = True
      DUT1_ParmBox(i).Visible = True
      HelpMeParmDUT1(i).Visible = True
      If IndexNumber = 0 Then
       DUT1_ParmBox(36).Text = "true"  'AC flag is true
      ElseIf IndexNumber = 2 Then
       DUT1_ParmBox(76).Text = "true"  'AC flag is true
      ElseIf IndexNumber = 4 Then
       DUT1_ParmBox(116).Text = "true"  'AC flag is true
      ElseIf IndexNumber = 6 Then
       DUT1_ParmBox(156).Text = "true"  'AC flag is true
      ElseIf IndexNumber = 8 Then
       DUT1_ParmBox(196).Text = "true"  'AC flag is true
      End If
     Next
     
    'DC selected
    Case 1, 3, 5, 7, 9
     For i = iStart To iEnd
      DUT1_Label(i).Visible = False
      DUT1_ParmBox(i).Visible = False
      HelpMeParmDUT1(i).Visible = False
      If IndexNumber = 1 Then
       DUT1_ParmBox(36).Text = "false"  'AC flag is false
      ElseIf IndexNumber = 3 Then
       DUT1_ParmBox(76).Text = "false"  'AC flag is false
      ElseIf IndexNumber = 5 Then
       DUT1_ParmBox(116).Text = "false"  'AC flag is false
      ElseIf IndexNumber = 7 Then
       DUT1_ParmBox(156).Text = "false"  'AC flag is false
      ElseIf IndexNumber = 9 Then
       DUT1_ParmBox(196).Text = "false"  'AC flag is false
      End If
     Next
     
End Select

End Sub

