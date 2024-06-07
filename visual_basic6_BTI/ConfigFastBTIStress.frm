VERSION 5.00
Begin VB.Form ConfigFastBTIStress 
   Caption         =   "Configure Channel connections for fast BTI stresses"
   ClientHeight    =   7656
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7656
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame SBOptions 
      Height          =   492
      Left            =   1104
      TabIndex        =   167
      Top             =   384
      Width           =   8700
      Begin VB.OptionButton UserDefinedOption 
         Caption         =   "User Defined"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   6000
         TabIndex        =   170
         Top             =   192
         Width           =   1500
      End
      Begin VB.OptionButton SBSysGND 
         Caption         =   "S/B @ Sys GND (5 DUTS)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   3216
         TabIndex        =   169
         Top             =   192
         Width           =   2805
      End
      Begin VB.OptionButton ALLWGFMU 
         Caption         =   "ALl WGFMU channels (3 DUTS)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   96
         TabIndex        =   168
         Top             =   192
         Width           =   3270
      End
   End
   Begin VB.CommandButton DoneButton 
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   9024
      TabIndex        =   166
      Top             =   7104
      Width           =   1548
   End
   Begin VB.Frame WGFMUChannelStatus 
      Caption         =   "WGFMU Slots"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6288
      Left            =   8928
      TabIndex        =   138
      Top             =   660
      Width           =   1932
      Begin VB.Frame Slot9 
         Caption         =   "Slot #9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   48
         TabIndex        =   147
         Top             =   288
         Width           =   1752
         Begin VB.CheckBox C902 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   1056
            TabIndex        =   149
            Top             =   288
            Width           =   396
         End
         Begin VB.CheckBox C901 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   192
            TabIndex        =   148
            Top             =   288
            Width           =   396
         End
      End
      Begin VB.Frame Slot8 
         Caption         =   "Slot #8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   96
         TabIndex        =   146
         Top             =   1008
         Width           =   1752
         Begin VB.CheckBox C801 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   240
            TabIndex        =   151
            Top             =   240
            Width           =   396
         End
         Begin VB.CheckBox C802 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   1104
            TabIndex        =   150
            Top             =   240
            Width           =   396
         End
      End
      Begin VB.Frame Slot7 
         Caption         =   "Slot #7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   96
         TabIndex        =   145
         Top             =   1632
         Width           =   1752
         Begin VB.CheckBox C701 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   240
            TabIndex        =   153
            Top             =   288
            Width           =   396
         End
         Begin VB.CheckBox C702 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   1104
            TabIndex        =   152
            Top             =   288
            Width           =   396
         End
      End
      Begin VB.Frame Slot1 
         Caption         =   "Slot #1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   96
         TabIndex        =   144
         Top             =   5664
         Width           =   1752
         Begin VB.CheckBox C101 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   240
            TabIndex        =   165
            Top             =   240
            Width           =   396
         End
         Begin VB.CheckBox C102 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   1104
            TabIndex        =   164
            Top             =   240
            Width           =   396
         End
      End
      Begin VB.Frame Slot2 
         Caption         =   "Slot #2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   96
         TabIndex        =   143
         Top             =   5040
         Width           =   1752
         Begin VB.CheckBox C201 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   288
            TabIndex        =   163
            Top             =   288
            Width           =   396
         End
         Begin VB.CheckBox C202 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   1152
            TabIndex        =   162
            Top             =   288
            Width           =   396
         End
      End
      Begin VB.Frame Slot3 
         Caption         =   "Slot #3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   96
         TabIndex        =   142
         Top             =   4368
         Width           =   1752
         Begin VB.CheckBox C301 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   384
            TabIndex        =   161
            Top             =   288
            Width           =   396
         End
         Begin VB.CheckBox C302 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   1248
            TabIndex        =   160
            Top             =   288
            Width           =   396
         End
      End
      Begin VB.Frame Slot4 
         Caption         =   "Slot #4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   96
         TabIndex        =   141
         Top             =   3696
         Width           =   1752
         Begin VB.CheckBox C401 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   384
            TabIndex        =   159
            Top             =   288
            Width           =   396
         End
         Begin VB.CheckBox C402 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   1248
            TabIndex        =   158
            Top             =   288
            Width           =   396
         End
      End
      Begin VB.Frame Slot5 
         Caption         =   "Slot #5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   144
         TabIndex        =   140
         Top             =   2976
         Width           =   1752
         Begin VB.CheckBox C501 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   192
            TabIndex        =   157
            Top             =   240
            Width           =   396
         End
         Begin VB.CheckBox C502 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   1056
            TabIndex        =   156
            Top             =   240
            Width           =   396
         End
      End
      Begin VB.Frame Slot6 
         Caption         =   "Slot #6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   552
         Left            =   96
         TabIndex        =   139
         Top             =   2256
         Width           =   1752
         Begin VB.CheckBox C601 
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   144
            TabIndex        =   155
            Top             =   240
            Width           =   396
         End
         Begin VB.CheckBox C602 
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   204
            Left            =   1008
            TabIndex        =   154
            Top             =   240
            Width           =   396
         End
      End
   End
   Begin VB.CommandButton getWGFMUChannels 
      Caption         =   "Get # of WGFMU Channels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   4512
      TabIndex        =   121
      Top             =   60
      Width           =   4140
   End
   Begin VB.Frame FrameforDUTInfo 
      Caption         =   "DUT9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   8
      Left            =   6108
      TabIndex        =   107
      Top             =   5340
      Width           =   2700
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   35
         Left            =   540
         TabIndex        =   115
         Text            =   "Slot #"
         Top             =   1620
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   35
         Left            =   1620
         TabIndex        =   114
         Text            =   "Chnl. #"
         Top             =   1680
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   34
         Left            =   540
         TabIndex        =   113
         Text            =   "Slot #"
         Top             =   1260
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   34
         Left            =   1620
         TabIndex        =   112
         Text            =   "Chnl. #"
         Top             =   1260
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   33
         Left            =   600
         TabIndex        =   111
         Text            =   "Slot #"
         Top             =   840
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   33
         Left            =   1620
         TabIndex        =   110
         Text            =   "Chnl. #"
         Top             =   900
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   32
         Left            =   540
         TabIndex        =   109
         Text            =   "Slot #"
         Top             =   420
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   32
         Left            =   1620
         TabIndex        =   108
         Text            =   "Chnl. #"
         Top             =   480
         Width           =   912
      End
      Begin VB.Label Label55 
         Alignment       =   2  'Center
         Caption         =   "Slot #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   300
         TabIndex        =   137
         Top             =   180
         Width           =   852
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         Caption         =   "Channel #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1320
         TabIndex        =   136
         Top             =   180
         Width           =   1152
      End
      Begin VB.Label Label38 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   119
         Top             =   0
         Width           =   372
      End
      Begin VB.Label Label37 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   118
         Top             =   360
         Width           =   372
      End
      Begin VB.Label Label36 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   117
         Top             =   720
         Width           =   372
      End
      Begin VB.Label Label35 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   116
         Top             =   1080
         Width           =   372
      End
   End
   Begin VB.Frame FrameforDUTInfo 
      Caption         =   "DUT8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   7
      Left            =   6108
      TabIndex        =   94
      Top             =   3000
      Width           =   2700
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   31
         Left            =   1620
         TabIndex        =   102
         Text            =   "Chnl. #"
         Top             =   1680
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   31
         Left            =   420
         TabIndex        =   101
         Text            =   "Slot #"
         Top             =   1680
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   30
         Left            =   1620
         TabIndex        =   100
         Text            =   "Chnl. #"
         Top             =   1260
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   30
         Left            =   420
         TabIndex        =   99
         Text            =   "Slot #"
         Top             =   1260
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   29
         Left            =   1560
         TabIndex        =   98
         Text            =   "Chnl. #"
         Top             =   900
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   29
         Left            =   540
         TabIndex        =   97
         Text            =   "Slot #"
         Top             =   900
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   28
         Left            =   1560
         TabIndex        =   96
         Text            =   "Chnl. #"
         Top             =   480
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   28
         Left            =   360
         TabIndex        =   95
         Text            =   "Slot #"
         Top             =   480
         Width           =   912
      End
      Begin VB.Label Label53 
         Alignment       =   2  'Center
         Caption         =   "Slot #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   135
         Top             =   180
         Width           =   852
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         Caption         =   "Channel #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1260
         TabIndex        =   134
         Top             =   180
         Width           =   1152
      End
      Begin VB.Label Label34 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   106
         Top             =   1080
         Width           =   372
      End
      Begin VB.Label Label33 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   105
         Top             =   720
         Width           =   372
      End
      Begin VB.Label Label32 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   104
         Top             =   360
         Width           =   372
      End
      Begin VB.Label Label31 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   103
         Top             =   0
         Width           =   372
      End
   End
   Begin VB.Frame FrameforDUTInfo 
      Caption         =   "DUT7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   6
      Left            =   6108
      TabIndex        =   41
      Top             =   660
      Width           =   2700
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   27
         Left            =   480
         TabIndex        =   93
         Text            =   "Slot #"
         Top             =   1800
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   27
         Left            =   1620
         TabIndex        =   92
         Text            =   "Chnl. #"
         Top             =   1740
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   26
         Left            =   540
         TabIndex        =   91
         Text            =   "Slot #"
         Top             =   1320
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   26
         Left            =   1560
         TabIndex        =   90
         Text            =   "Chnl. #"
         Top             =   1320
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   25
         Left            =   600
         TabIndex        =   89
         Text            =   "Slot #"
         Top             =   900
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   25
         Left            =   1620
         TabIndex        =   88
         Text            =   "Chnl. #"
         Top             =   840
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   24
         Left            =   600
         TabIndex        =   87
         Text            =   "Slot #"
         Top             =   600
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   24
         Left            =   1680
         TabIndex        =   86
         Text            =   "Chnl. #"
         Top             =   480
         Width           =   912
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         Caption         =   "Slot #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   300
         TabIndex        =   133
         Top             =   180
         Width           =   852
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Caption         =   "Channel #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1320
         TabIndex        =   132
         Top             =   180
         Width           =   1152
      End
      Begin VB.Label Label30 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Width           =   372
      End
      Begin VB.Label Label29 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   44
         Top             =   360
         Width           =   372
      End
      Begin VB.Label Label28 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   43
         Top             =   720
         Width           =   372
      End
      Begin VB.Label Label27 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   42
         Top             =   1080
         Width           =   372
      End
   End
   Begin VB.Frame FrameforDUTInfo 
      Caption         =   "DUT6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   5
      Left            =   3108
      TabIndex        =   36
      Top             =   5340
      Width           =   2700
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   23
         Left            =   420
         TabIndex        =   85
         Text            =   "Slot #"
         Top             =   1620
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   23
         Left            =   1500
         TabIndex        =   84
         Text            =   "Chnl. #"
         Top             =   1680
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   22
         Left            =   420
         TabIndex        =   83
         Text            =   "Slot #"
         Top             =   1200
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   22
         Left            =   1500
         TabIndex        =   82
         Text            =   "Chnl. #"
         Top             =   1260
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   21
         Left            =   420
         TabIndex        =   81
         Text            =   "Slot #"
         Top             =   840
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   21
         Left            =   1440
         TabIndex        =   80
         Text            =   "Chnl. #"
         Top             =   900
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   20
         Left            =   360
         TabIndex        =   79
         Text            =   "Slot #"
         Top             =   420
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   20
         Left            =   1440
         TabIndex        =   78
         Text            =   "Chnl. #"
         Top             =   480
         Width           =   912
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "Slot #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   240
         TabIndex        =   131
         Top             =   120
         Width           =   852
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         Caption         =   "Channel #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1260
         TabIndex        =   130
         Top             =   120
         Width           =   1152
      End
      Begin VB.Label Label26 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   372
      End
      Begin VB.Label Label25 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   39
         Top             =   720
         Width           =   372
      End
      Begin VB.Label Label24 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   38
         Top             =   360
         Width           =   372
      End
      Begin VB.Label Label23 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   372
      End
   End
   Begin VB.Frame FrameforDUTInfo 
      Caption         =   "DUT1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   0
      Left            =   108
      TabIndex        =   1
      Top             =   660
      Width           =   2700
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   1
         Left            =   360
         TabIndex        =   17
         Text            =   "Slot #"
         Top             =   960
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   2
         Left            =   360
         TabIndex        =   16
         Text            =   "Slot #"
         Top             =   1320
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   3
         Left            =   360
         TabIndex        =   15
         Text            =   "Slot #"
         Top             =   1680
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   0
         Left            =   1380
         TabIndex        =   14
         Text            =   "Chnl. #"
         Top             =   600
         Width           =   1100
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   1
         Left            =   1380
         TabIndex        =   13
         Text            =   "Chnl. #"
         Top             =   960
         Width           =   1100
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   2
         Left            =   1380
         TabIndex        =   12
         Text            =   "Chnl. #"
         Top             =   1320
         Width           =   1100
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   3
         Left            =   1380
         TabIndex        =   11
         Text            =   "Chnl. #"
         Top             =   1680
         Width           =   1100
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   0
         ItemData        =   "ConfigFastBTIStress.frx":0000
         Left            =   360
         List            =   "ConfigFastBTIStress.frx":0002
         TabIndex        =   10
         Text            =   "Slot #"
         Top             =   600
         Width           =   912
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Channel #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1320
         TabIndex        =   19
         Top             =   300
         Width           =   1152
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Slot #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   360
         TabIndex        =   18
         Top             =   300
         Width           =   852
      End
      Begin VB.Label Label4 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   9
         Top             =   1740
         Width           =   372
      End
      Begin VB.Label Label3 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   8
         Top             =   1380
         Width           =   372
      End
      Begin VB.Label Label2 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   7
         Top             =   1020
         Width           =   372
      End
      Begin VB.Label Label1 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   660
         Width           =   372
      End
   End
   Begin VB.Frame FrameforDUTInfo 
      Caption         =   "DUT5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   4
      Left            =   3108
      TabIndex        =   5
      Top             =   3000
      Width           =   2700
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   19
         Left            =   420
         TabIndex        =   77
         Text            =   "Slot #"
         Top             =   1740
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   19
         Left            =   1500
         TabIndex        =   76
         Text            =   "Chnl. #"
         Top             =   1800
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   18
         Left            =   420
         TabIndex        =   75
         Text            =   "Slot #"
         Top             =   1320
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   18
         Left            =   1500
         TabIndex        =   74
         Text            =   "Chnl. #"
         Top             =   1440
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   17
         Left            =   420
         TabIndex        =   73
         Text            =   "Slot #"
         Top             =   960
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   17
         Left            =   1500
         TabIndex        =   72
         Text            =   "Chnl. #"
         Top             =   960
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   16
         Left            =   420
         TabIndex        =   71
         Text            =   "Slot #"
         Top             =   600
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   16
         Left            =   1500
         TabIndex        =   70
         Text            =   "Chnl. #"
         Top             =   660
         Width           =   912
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         Caption         =   "Slot #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   360
         TabIndex        =   129
         Top             =   180
         Width           =   852
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "Channel #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1380
         TabIndex        =   128
         Top             =   180
         Width           =   1152
      End
      Begin VB.Label Label22 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   372
      End
      Begin VB.Label Label21 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   34
         Top             =   360
         Width           =   372
      End
      Begin VB.Label Label20 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   33
         Top             =   720
         Width           =   372
      End
      Begin VB.Label Label19 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   372
      End
   End
   Begin VB.Frame FrameforDUTInfo 
      Caption         =   "DUT4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   3
      Left            =   3108
      TabIndex        =   4
      Top             =   660
      Width           =   2700
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   15
         Left            =   660
         TabIndex        =   69
         Text            =   "Slot #"
         Top             =   1680
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   15
         Left            =   1740
         TabIndex        =   68
         Text            =   "Chnl. #"
         Top             =   1680
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   14
         Left            =   660
         TabIndex        =   67
         Text            =   "Slot #"
         Top             =   1260
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   14
         Left            =   1740
         TabIndex        =   66
         Text            =   "Chnl. #"
         Top             =   1260
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   13
         Left            =   600
         TabIndex        =   65
         Text            =   "Slot #"
         Top             =   900
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   13
         Left            =   1680
         TabIndex        =   64
         Text            =   "Chnl. #"
         Top             =   900
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   12
         Left            =   600
         TabIndex        =   63
         Text            =   "Slot #"
         Top             =   660
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   12
         Left            =   1680
         TabIndex        =   62
         Text            =   "Chnl. #"
         Top             =   600
         Width           =   912
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Caption         =   "Slot #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   420
         TabIndex        =   127
         Top             =   120
         Width           =   852
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         Caption         =   "Channel #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1440
         TabIndex        =   126
         Top             =   120
         Width           =   1152
      End
      Begin VB.Label Label18 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   372
      End
      Begin VB.Label Label17 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   30
         Top             =   360
         Width           =   372
      End
      Begin VB.Label Label16 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   0
         TabIndex        =   29
         Top             =   720
         Width           =   372
      End
      Begin VB.Label Label15 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   372
      End
   End
   Begin VB.Frame FrameforDUTInfo 
      Caption         =   "DUT3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   2
      Left            =   108
      TabIndex        =   3
      Top             =   5340
      Width           =   2700
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   11
         Left            =   1620
         TabIndex        =   61
         Text            =   "Chnl. #"
         Top             =   1680
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   10
         Left            =   1500
         TabIndex        =   60
         Text            =   "Chnl. #"
         Top             =   1260
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   11
         Left            =   420
         TabIndex        =   59
         Text            =   "Slot #"
         Top             =   1680
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   10
         Left            =   420
         TabIndex        =   58
         Text            =   "Slot #"
         Top             =   1260
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   8
         Left            =   420
         TabIndex        =   57
         Text            =   "Slot #"
         Top             =   480
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   9
         Left            =   420
         TabIndex        =   56
         Text            =   "Slot #"
         Top             =   840
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   8
         Left            =   1500
         TabIndex        =   55
         Text            =   "Chnl. #"
         Top             =   480
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   9
         Left            =   1560
         TabIndex        =   54
         Text            =   "Chnl. #"
         Top             =   840
         Width           =   912
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         Caption         =   "Slot #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   360
         TabIndex        =   125
         Top             =   180
         Width           =   852
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         Caption         =   "Channel #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1380
         TabIndex        =   124
         Top             =   180
         Width           =   1152
      End
      Begin VB.Label Label14 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   372
      End
      Begin VB.Label Label13 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   60
         TabIndex        =   26
         Top             =   840
         Width           =   372
      End
      Begin VB.Label Label12 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   372
      End
      Begin VB.Label Label11 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   60
         TabIndex        =   24
         Top             =   1680
         Width           =   372
      End
   End
   Begin VB.Frame FrameforDUTInfo 
      Caption         =   "DUT2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Index           =   1
      Left            =   108
      TabIndex        =   2
      Top             =   3000
      Width           =   2700
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   4
         Left            =   600
         TabIndex        =   53
         Text            =   "Slot #"
         Top             =   420
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   5
         Left            =   540
         TabIndex        =   52
         Text            =   "Slot #"
         Top             =   780
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   6
         Left            =   420
         TabIndex        =   51
         Text            =   "Slot #"
         Top             =   1260
         Width           =   912
      End
      Begin VB.ComboBox SlotDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   7
         Left            =   540
         TabIndex        =   50
         Text            =   "Slot #"
         Top             =   1620
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   4
         Left            =   1500
         TabIndex        =   49
         Text            =   "Chnl. #"
         Top             =   480
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   5
         Left            =   1440
         TabIndex        =   48
         Text            =   "Chnl. #"
         Top             =   840
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   6
         Left            =   1440
         TabIndex        =   47
         Text            =   "Chnl. #"
         Top             =   1320
         Width           =   912
      End
      Begin VB.ComboBox ChnDUT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   336
         Index           =   7
         Left            =   1440
         TabIndex        =   46
         Text            =   "Chnl. #"
         Top             =   1680
         Width           =   912
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         Caption         =   "Slot #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   360
         TabIndex        =   123
         Top             =   240
         Width           =   852
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         Caption         =   "Channel #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1380
         TabIndex        =   122
         Top             =   240
         Width           =   1152
      End
      Begin VB.Label Label10 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   23
         Top             =   1620
         Width           =   372
      End
      Begin VB.Label Label9 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   372
      End
      Begin VB.Label Label8 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   60
         TabIndex        =   21
         Top             =   840
         Width           =   372
      End
      Begin VB.Label Label7 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   372
      End
   End
   Begin VB.TextBox AvailableWFGMUSlots 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   3840
      TabIndex        =   0
      Top             =   60
      Width           =   552
   End
   Begin VB.Label Label39 
      Caption         =   "Number of active WGFMU channels ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Left            =   180
      TabIndex        =   120
      Top             =   60
      Width           =   3672
   End
End
Attribute VB_Name = "ConfigFastBTIStress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ALLWGFMU_Click()
'------------------------------------------------
'Slot number              = 9/8/7/6/5/4/3/2/1/-1 corresponds to
'Index number of combobox = 0/1/2/3/4/5/6/7/8/9
 'DUT #1
 SlotDUT(0).ListIndex = 0 'G slot=9
 SlotDUT(1).ListIndex = 0 'D slot=9
 SlotDUT(2).ListIndex = 1 'S slot=8
 SlotDUT(3).ListIndex = 4 'B slot=5
 ChnDUT(0).ListIndex = 0  'G channel = 1
 ChnDUT(1).ListIndex = 1  'D channel = 2
 ChnDUT(2).ListIndex = 0  'S channel = 1
 ChnDUT(3).ListIndex = 1  'B channel = 2
 'DUT #2
 SlotDUT(4).ListIndex = 1 'G slot=8
 SlotDUT(5).ListIndex = 2 'D slot=7
 SlotDUT(6).ListIndex = 2 'S slot=7
 SlotDUT(7).ListIndex = 4 'B slot=5
 ChnDUT(4).ListIndex = 1  'G channel = 2
 ChnDUT(5).ListIndex = 0  'D channel = 1
 ChnDUT(6).ListIndex = 1  'S channel = 2
 ChnDUT(7).ListIndex = 1  'B channel = 2
 'DUT #3
 SlotDUT(8).ListIndex = 3 'G slot=6
 SlotDUT(9).ListIndex = 3 'D slot=6
 SlotDUT(10).ListIndex = 4 'S slot=5
 SlotDUT(11).ListIndex = 4 'B slot=5
 ChnDUT(8).ListIndex = 0  'G channel = 1
 ChnDUT(9).ListIndex = 1  'D channel = 2
 ChnDUT(10).ListIndex = 0  'S channel = 1
 ChnDUT(11).ListIndex = 1  'B channel = 2
'------------------------------------------------
'Show frames only for # of DUTs user specified
Dim MaxPossibleDUTstoStress As Integer
'Get # of DUTs to be stressed from user input
MaxPossibleDUTstoStress = Val(WGFMUVBInput.NumDevWGFMU.Text)
For i = 0 To (MaxPossibleDUTstoStress - 1)
 FrameforDUTInfo(i).Visible = True
Next
For i = MaxPossibleDUTstoStress To 8
 FrameforDUTInfo(i).Visible = False
Next
'If more than 3 DUTs were specified, then only shows 3 DUTs
'for the all WGFMU channel option
If MaxPossibleDUTstoStress > 3 Then
 For i = 3 To 8
  FrameforDUTInfo(i).Visible = False
 Next
End If

End Sub

Private Sub AvailableWFGMUSlots_Change()
'-----------------------------------
'Do not display 0 but just empty box for total number of active WGFMU channels
If Val(AvailableWFGMUSlots.Text) = 0 Then AvailableWFGMUSlots.Text = ""
'-----------------------------------
'Error pretection when user specifes values 0, 1 or > 10
'for the number of active WGFMU channels
Dim MaxNumChannelsAvailable As Integer
MaxNumChannelsAvailable = 10
TotNumActiveChannels = Val(AvailableWFGMUSlots.Text)
'If box becomes empty
'If (AvailableWFGMUSlots.Text = "") Then
'    MsgBox ("Value must be between 2 and 10")
'    AvailableWFGMUSlots.Text = 2
'End If
'If box has 0 value
'If (TotNumActiveChannels <= 0) Then
'    MsgBox ("Value must be between 2 and 10")
'    AvailableWFGMUSlots.Text = 2
'End If
'If (TotNumActiveChannels > MaxNumChannelsAvailable) Then
'    MsgBox ("Maximum allowed WGFMU challens for fast BTI stress = " & Str$(MaxNumChannelsAvailable))
'    AvailableWFGMUSlots.Text = MaxNumChannelsAvailable
'End If
'-----------------------------------
Dim MaxPossibleDUTstoStress As Integer

'Always round down integer to calculate the maximum possible number of DUTs
'for a given number of active WGFMU channels
'Note: backslash \ operator always returns an integer
'MaxPossibleDUTstoStress is in fact the total number channels, not the total number of WGFMu slots
'MaxPossibleDUTstoStress = Int(Val(AvailableWFGMUSlots.Text) + 0.99) \ 2
'For I = 0 To (MaxPossibleDUTstoStress - 1)
' FrameforDUTInfo(I).Visible = True
'Next
'For I = MaxPossibleDUTstoStress To 8
' FrameforDUTInfo(I).Visible = False
'Next
'--------------------------------------------------------
'Set the slot and channel number of Substrate (or Body) terminal
'with the lowest slot number and channel number = 2
'Dim LowestSlotNumber As Integer
'LowestSlotNumber = Int(Val(AvailableWFGMUSlots.Text) + 0.99) \ 2
'MaxPossibleTotNumActiveChannels = MaxNumChannelsAvailable
'IncVal = 0
'IndexInc = -1
'For I = 0 To ((MaxPossibleTotNumActiveChannels - 1) - 1)
 'Set WGFMU slot numbers
' IncVal = 4 * I
 'ListIndex=1/2/3/4/5/6/7/8/9 ==> Value of combobox=9/8/7/6/5/4/3/2/1/-1
' If (IndexInc <= 8) Then
'  SlotDUT(3 + IncVal).ListIndex = (LowestSlotNumber - 1) 'show slot number for Substrate
 'Set the channel numbers
'  ChnDUT(3 + IncVal).ListIndex = 1   'show channel number =2 for Substrate
' ElseIf (IndexInc > 8) Then 'ran out of slot numbers and set all equal to -1
'  SlotDUT(3 + IncVal).ListIndex = (LowestSlotNumber - 1)  'show slot number = -1for Substrate
 'Set the channel numbers
'  ChnDUT(3 + IncVal).ListIndex = 2   'show channel number =2 for Substrate
' End If
'Next

If TotNumActiveChannels = 2 Then
  C901ValueStored = 1
  C902ValueStored = 1
  C801ValueStored = 0
  C802ValueStored = 0
  C701ValueStored = 0
  C702ValueStored = 0
  C601ValueStored = 0
  C602ValueStored = 0
  C501ValueStored = 0
  C502ValueStored = 0
ElseIf TotNumActiveChannels = 4 Then
  C901ValueStored = 1
  C902ValueStored = 1
  C801ValueStored = 1
  C802ValueStored = 1
  C701ValueStored = 0
  C702ValueStored = 0
  C601ValueStored = 0
  C602ValueStored = 0
  C501ValueStored = 0
  C502ValueStored = 0
ElseIf TotNumActiveChannels = 6 Then
  C901ValueStored = 1
  C902ValueStored = 1
  C801ValueStored = 1
  C802ValueStored = 1
  C701ValueStored = 1
  C702ValueStored = 1
  C601ValueStored = 0
  C602ValueStored = 0
  C501ValueStored = 0
  C502ValueStored = 0
ElseIf TotNumActiveChannels = 8 Then
  C901ValueStored = 1
  C902ValueStored = 1
  C801ValueStored = 1
  C802ValueStored = 1
  C701ValueStored = 1
  C702ValueStored = 1
  C601ValueStored = 1
  C602ValueStored = 1
  C501ValueStored = 0
  C502ValueStored = 0
ElseIf TotNumActiveChannels = 10 Then
  C901ValueStored = 1
  C902ValueStored = 1
  C801ValueStored = 1
  C802ValueStored = 1
  C701ValueStored = 1
  C702ValueStored = 1
  C601ValueStored = 1
  C602ValueStored = 1
  C501ValueStored = 1
  C502ValueStored = 1
ElseIf TotNumActiveChannels < 0 Then
  C901ValueStored = 0
  C902ValueStored = 0
  C801ValueStored = 0
  C802ValueStored = 0
  C701ValueStored = 0
  C702ValueStored = 0
  C601ValueStored = 0
  C602ValueStored = 0
  C501ValueStored = 0
  C502ValueStored = 0
End If
'Reset CheckerBox value
C901.value = C901ValueStored
C902.value = C902ValueStored
C801.value = C801ValueStored
C802.value = C802ValueStored
C701.value = C701ValueStored
C702.value = C702ValueStored
C601.value = C601ValueStored
C602.value = C602ValueStored
C501.value = C501ValueStored
C502.value = C502ValueStored
End Sub
Private Sub DoneButton_Click()
Dim ICancel As Integer
Call Form_Unload(ICancel)
Unload Me
End Sub
Private Sub Form_Load()
'----------------------------------------
'Do not allow any changes on the main form when Config form is active
WGFMUVBInput.Enabled = False
'----------------------------------------
'Do not allow user to change value (total number
'of active WGFMU channels) remotely obtained from B1500
AvailableWFGMUSlots.Locked = True
'----------------------------------------
WGFMUVBInput.RunButton.Enabled = False
'----------------------------------------
Dim TotNumActiveChannels As Integer
TotNumActiveChannels = Val(AvailableWFGMUSlots.Text)
'-----------------------------------
'Count how many times Config_form is loaded
NumPopup_ConfigFastBTIStress = NumPopup_ConfigFastBTIStress + 1
'MsgBox (NumPopup_ConfigFastBTIStress)
'-----------------------------------
If NumPopup_ConfigFastBTIStress > 1 Then
 TotNumActiveChannels = StoreTotNumActiveChannels
 AvailableWFGMUSlots.Text = TotNumActiveChannels
End If
'-----------------------------------
ConfigFastBTIStress.Width = 8350
ConfigFastBTIStress.Height = 9000
'-----------------------------------
Slot9.Left = 96
Slot8.Left = 96
Slot7.Left = 96
Slot6.Left = 96
Slot5.Left = 96
Slot4.Left = 96
Slot3.Left = 96
Slot2.Left = 96
Slot1.Left = 96
DelY = Slot9.Height + 150
Slot9.Top = 336
Slot8.Top = Slot9.Top + DelY
Slot7.Top = Slot9.Top + 2 * DelY
Slot6.Top = Slot9.Top + 3 * DelY
Slot5.Top = Slot9.Top + 4 * DelY
Slot4.Top = Slot9.Top + 5 * DelY
Slot3.Top = Slot9.Top + 6 * DelY
Slot2.Top = Slot9.Top + 7 * DelY
Slot1.Top = Slot9.Top + 8 * DelY
'Channel radio button
C901.Left = 200
C801.Left = 200
C701.Left = 200
C601.Left = 200
C501.Left = 200
C401.Left = 200
C301.Left = 200
C201.Left = 200
C101.Left = 200
'
C902.Left = 1200
C802.Left = 1200
C702.Left = 1200
C602.Left = 1200
C502.Left = 1200
C402.Left = 1200
C302.Left = 1200
C202.Left = 1200
C102.Left = 1200
'
C901.Top = 250
C902.Top = C901.Top
C801.Top = C901.Top
C802.Top = C901.Top
C701.Top = C901.Top
C702.Top = C901.Top
C601.Top = C901.Top
C602.Top = C901.Top
C501.Top = C901.Top
C502.Top = C901.Top
C401.Top = C901.Top
C402.Top = C901.Top
C301.Top = C901.Top
C302.Top = C901.Top
C201.Top = C901.Top
C202.Top = C901.Top
C101.Top = C901.Top
C102.Top = C901.Top
'
C901.Enabled = False
C902.Enabled = False
C801.Enabled = False
C802.Enabled = False
C701.Enabled = False
C702.Enabled = False
C601.Enabled = False
C602.Enabled = False
C501.Enabled = False
C502.Enabled = False
C401.Enabled = False
C402.Enabled = False
C301.Enabled = False
C302.Enabled = False
C201.Enabled = False
C202.Enabled = False
C101.Enabled = False
C102.Enabled = False
'--------------------------------------------------
'Set the widt hof button to get the totl number of active WGFMU channels from B1500
getWGFMUChannels.Width = 3500
'--------------------------------------------------
DeltaX = 3000
DeltaY = 2340
'Frame for DUT #1
FrameforDUTInfo(0).Left = 120
FrameforDUTInfo(0).Top = 1000
SBOptions.Left = FrameforDUTInfo(0).Left
SBOptions.Top = FrameforDUTInfo(0).Top - 600
SBOptions.Width = FrameforDUTInfo(0).Width + DeltaX + 300 + WGFMUChannelStatus.Width

'Frame for DUT #2
FrameforDUTInfo(1).Left = FrameforDUTInfo(0).Left
FrameforDUTInfo(1).Top = FrameforDUTInfo(0).Top + (1 * DeltaY)
'Frame for DUT #3
FrameforDUTInfo(2).Left = FrameforDUTInfo(0).Left
FrameforDUTInfo(2).Top = FrameforDUTInfo(0).Top + (2 * DeltaY)
'
'Frame for DUT #4
FrameforDUTInfo(3).Left = FrameforDUTInfo(0).Left + (1 * DeltaX)
FrameforDUTInfo(3).Top = FrameforDUTInfo(0).Top
'Frame for DUT #5
FrameforDUTInfo(4).Left = FrameforDUTInfo(3).Left
FrameforDUTInfo(4).Top = FrameforDUTInfo(0).Top + (1 * DeltaY)
'Frame for DUT #6
FrameforDUTInfo(5).Left = FrameforDUTInfo(3).Left
FrameforDUTInfo(5).Top = FrameforDUTInfo(0).Top + (2 * DeltaY)
'
'Frame for DUT #7
FrameforDUTInfo(6).Left = FrameforDUTInfo(0).Left + (2 * DeltaX)
FrameforDUTInfo(6).Top = FrameforDUTInfo(0).Top
'Frame for DUT #8
FrameforDUTInfo(7).Left = FrameforDUTInfo(6).Left
FrameforDUTInfo(7).Top = FrameforDUTInfo(0).Top + (1 * DeltaY)
'Frame for DUT #9
FrameforDUTInfo(8).Left = FrameforDUTInfo(6).Left
FrameforDUTInfo(8).Top = FrameforDUTInfo(0).Top + (2 * DeltaY)
'-----------------------------------
'Read total number of active WGFMU channels
TotNumActiveChannels = Val(AvailableWFGMUSlots.Text)   'this value must come from B1500
'-----------------------------------
'Set locations and alignment of comboboxes
'Labels of G/D/S/B are not made using an array - simple mistake
'Set label locations and alignment using their names
'------------------------------------------------------
'Labels for Slot # and Channel #
'For DUT #1
Me.Label5.Left = 360
Me.Label5.Top = 300
Me.Label6.Left = 1320
Me.Label6.Top = 300
X1Loc = Me.Label5.Left
Y1Loc = Me.Label5.Top
X2Loc = Me.Label6.Left
Y2Loc = Me.Label6.Top
'For DUT #2
Me.Label41.Left = X1Loc
Me.Label41.Top = Y1Loc
Me.Label40.Left = X2Loc
Me.Label40.Top = Y2Loc
'For DUT #3
Me.Label43.Left = X1Loc
Me.Label43.Top = Y1Loc
Me.Label42.Left = X2Loc
Me.Label42.Top = Y2Loc
'For DUT #4
Me.Label45.Left = X1Loc
Me.Label45.Top = Y1Loc
Me.Label44.Left = X2Loc
Me.Label44.Top = Y2Loc
'For DUT #5
Me.Label47.Left = X1Loc
Me.Label47.Top = Y1Loc
Me.Label46.Left = X2Loc
Me.Label46.Top = Y2Loc
'For DUT #6
Me.Label49.Left = X1Loc
Me.Label49.Top = Y1Loc
Me.Label48.Left = X2Loc
Me.Label48.Top = Y2Loc
'For DUT #7
Me.Label51.Left = X1Loc
Me.Label51.Top = Y1Loc
Me.Label50.Left = X2Loc
Me.Label50.Top = Y2Loc
'For DUT #8
Me.Label53.Left = X1Loc
Me.Label53.Top = Y1Loc
Me.Label52.Left = X2Loc
Me.Label52.Top = Y2Loc
'For DUT #9
Me.Label55.Left = X1Loc
Me.Label55.Top = Y1Loc
Me.Label54.Left = X2Loc
Me.Label54.Top = Y2Loc
'------------------------------------------------------
'Labels for G/D/S/B
DeltaY = 350
'Frame for DUT #1
Me.Label1.Left = 120
Me.Label2.Left = Me.Label1.Left
Me.Label3.Left = Me.Label1.Left
Me.Label4.Left = Me.Label1.Left
Me.Label1.Top = 660
Me.Label2.Top = Me.Label1.Top + (1 * DeltaY)
Me.Label3.Top = Me.Label1.Top + (2 * DeltaY)
Me.Label4.Top = Me.Label1.Top + (3 * DeltaY)
X1Loc = Me.Label1.Left
X2Loc = Me.Label1.Left
X3Loc = Me.Label1.Left
X4Loc = Me.Label1.Left
Y1Loc = Me.Label1.Top
Y2Loc = Me.Label2.Top
Y3Loc = Me.Label3.Top
Y4Loc = Me.Label4.Top
'Frame for DUT #2
Me.Label7.Left = X1Loc
Me.Label8.Left = X2Loc
Me.Label9.Left = X3Loc
Me.Label10.Left = X4Loc
Me.Label7.Top = Y1Loc
Me.Label8.Top = Y2Loc
Me.Label9.Top = Y3Loc
Me.Label10.Top = Y4Loc
'Frame for DUT #3
Me.Label11.Left = X1Loc
Me.Label12.Left = X2Loc
Me.Label13.Left = X3Loc
Me.Label14.Left = X4Loc
Me.Label11.Top = Me.Label1.Top
Me.Label12.Top = Me.Label2.Top
Me.Label13.Top = Me.Label3.Top
Me.Label14.Top = Me.Label4.Top
'Frame for DUT #4
Me.Label15.Left = X1Loc
Me.Label16.Left = X2Loc
Me.Label17.Left = X3Loc
Me.Label18.Left = X4Loc
Me.Label15.Top = Me.Label1.Top
Me.Label16.Top = Me.Label2.Top
Me.Label17.Top = Me.Label3.Top
Me.Label18.Top = Me.Label4.Top
'Frame for DUT #5
Me.Label19.Left = X1Loc
Me.Label20.Left = X2Loc
Me.Label21.Left = X3Loc
Me.Label22.Left = X4Loc
Me.Label19.Top = Me.Label1.Top
Me.Label20.Top = Me.Label2.Top
Me.Label21.Top = Me.Label3.Top
Me.Label22.Top = Me.Label4.Top
'Frame for DUT #6
Me.Label23.Left = X1Loc
Me.Label24.Left = X2Loc
Me.Label25.Left = X3Loc
Me.Label26.Left = X4Loc
Me.Label23.Top = Me.Label1.Top
Me.Label24.Top = Me.Label2.Top
Me.Label25.Top = Me.Label3.Top
Me.Label26.Top = Me.Label4.Top
'Frame for DUT #7
Me.Label27.Left = X1Loc
Me.Label28.Left = X2Loc
Me.Label29.Left = X3Loc
Me.Label30.Left = X4Loc
Me.Label27.Top = Me.Label1.Top
Me.Label28.Top = Me.Label2.Top
Me.Label29.Top = Me.Label3.Top
Me.Label30.Top = Me.Label4.Top
'Frame for DUT #8
Me.Label31.Left = X1Loc
Me.Label32.Left = X2Loc
Me.Label33.Left = X3Loc
Me.Label34.Left = X4Loc
Me.Label31.Top = Me.Label1.Top
Me.Label32.Top = Me.Label2.Top
Me.Label33.Top = Me.Label3.Top
Me.Label34.Top = Me.Label4.Top
'Frame for DUT #9
Me.Label35.Left = X1Loc
Me.Label36.Left = X2Loc
Me.Label37.Left = X3Loc
Me.Label38.Left = X4Loc
Me.Label35.Top = Me.Label1.Top
Me.Label36.Top = Me.Label2.Top
Me.Label37.Top = Me.Label3.Top
Me.Label38.Top = Me.Label4.Top
'------------------------------------------------------
'Set width of comboboxes
ComboBoxWidth = 912
For i = 0 To 35
 SlotDUT(i).Width = ComboBoxWidth
 ChnDUT(i).Width = ComboBoxWidth
Next
'------------------------------------------------------
'Combobox for Slot #
'For DUT #1
DeltaY = 360
SlotDUT(0).Left = 360
SlotDUT(0).Top = 600
SlotDUT(1).Left = SlotDUT(0).Left
SlotDUT(1).Top = SlotDUT(0).Top + (1 * DeltaY)
SlotDUT(2).Left = SlotDUT(0).Left
SlotDUT(2).Top = SlotDUT(0).Top + (2 * DeltaY)
SlotDUT(3).Left = SlotDUT(0).Left
SlotDUT(3).Top = SlotDUT(0).Top + (3 * DeltaY)
X1Loc = SlotDUT(0).Left
Y1Loc = SlotDUT(0).Top
Y2Loc = SlotDUT(1).Top
Y3Loc = SlotDUT(2).Top
Y4Loc = SlotDUT(3).Top
'For DUT #2 to #9
JStart = 4
For i = 2 To 9
 JCount = 0
 For J = JStart To (JStart + 3)
  JCount = JCount + 1
  SlotDUT(J).Left = X1Loc
  If (JCount = 1) Then
   SlotDUT(J).Top = Y1Loc
  ElseIf (JCount = 2) Then
   SlotDUT(J).Top = Y2Loc
  ElseIf (JCount = 3) Then
   SlotDUT(J).Top = Y3Loc
  ElseIf (JCount = 4) Then
   SlotDUT(J).Top = Y4Loc
  End If
 Next
 JStart = JStart + 4
Next
'------------------------------------------------------
'Combobox for Chnl. #
'For DUT #1
DeltaY = 360
ChnDUT(0).Left = 1380
ChnDUT(0).Top = 600
ChnDUT(1).Left = ChnDUT(0).Left
ChnDUT(1).Top = ChnDUT(0).Top + (1 * DeltaY)
ChnDUT(2).Left = ChnDUT(0).Left
ChnDUT(2).Top = ChnDUT(0).Top + (2 * DeltaY)
ChnDUT(3).Left = ChnDUT(0).Left
ChnDUT(3).Top = ChnDUT(0).Top + (3 * DeltaY)
X1Loc = ChnDUT(0).Left
Y1Loc = ChnDUT(0).Top
Y2Loc = ChnDUT(1).Top
Y3Loc = ChnDUT(2).Top
Y4Loc = ChnDUT(3).Top
'For DUT #2 to #9
JStart = 4
For i = 2 To 9
 JCount = 0
 For J = JStart To (JStart + 3)
  JCount = JCount + 1
  ChnDUT(J).Left = X1Loc
  If (JCount = 1) Then
   ChnDUT(J).Top = Y1Loc
  ElseIf (JCount = 2) Then
   ChnDUT(J).Top = Y2Loc
  ElseIf (JCount = 3) Then
   ChnDUT(J).Top = Y3Loc
  ElseIf (JCount = 4) Then
   ChnDUT(J).Top = Y4Loc
  End If
 Next
 JStart = JStart + 4
Next
'-----------------------------------
'Frame to display WGFMU channel status
WGFMUChannelStatus.Top = FrameforDUTInfo(0).Top
WGFMUChannelStatus.Left = FrameforDUTInfo(6).Left
WGFMUChannelStatus.Height = 9 * (Slot9.Height + 180)
DoneButton.Left = WGFMUChannelStatus.Left + 200
DoneButton.Top = WGFMUChannelStatus.Top + WGFMUChannelStatus.Height + 50

'-----------------------------------
'Opening up this form for the first time, show the following values
'Default setting and values only when this form shows up for the 1st time
'These settings have nothing to do with TotNumActiveChannels obtained from B1500
Dim MaxPossibleTotNumActiveChannels As Integer
MaxPossibleTotNumActiveChannels = 10
For i = 0 To 8
 FrameforDUTInfo(i).Visible = True
Next
'Make all combo boxes visible at start up
For i = 0 To (4 * (MaxPossibleTotNumActiveChannels - 1) - 1)
 SlotDUT(i).Visible = True
 ChnDUT(i).Visible = True
Next
'Default text for combo boxes
For i = 0 To (4 * (MaxPossibleTotNumActiveChannels - 1) - 1)
 'SlotDUT(I).Text = "slot #"
 'ChnDUT(I).Text = "channel #"
 SlotDUT(i).Text = ""
 ChnDUT(i).Text = ""
Next
'----------------------------------------------------------
'Here's where slot and channel numbers are set
'Add available slot numbers to each combo box
For i = 0 To (4 * (MaxPossibleTotNumActiveChannels - 1) - 1)
 StartslotNum = 9
 SlotNum = StartslotNum
  For J = 0 To 9
  If SlotNum = 0 Then
   SlotDUT(i).AddItem -1
  Else
   SlotDUT(i).AddItem SlotNum
  End If
  SlotNum = SlotNum - 1
 Next
Next
'Add available channel numbers to each combo box
For i = 0 To (4 * (MaxPossibleTotNumActiveChannels - 1) - 1)
 ChnDUT(i).AddItem 1
 ChnDUT(i).AddItem 2
 ChnDUT(i).AddItem -1
Next
'----------------------------------------------------------
If NumPopup_ConfigFastBTIStress = 1 Then
 For i = 0 To (4 * (MaxPossibleTotNumActiveChannels - 1) - 1)
  Store_Slot_IndexNumber(i) = SlotDUT(i).ListIndex
  Store_Channel_IndexNumber(i) = ChnDUT(i).ListIndex
 Next
 If TotNumActiveChannels = 2 Then
  C901ValueStored = 1
  C902ValueStored = 1
  C801ValueStored = 0
  C802ValueStored = 0
  C701ValueStored = 0
  C702ValueStored = 0
  C601ValueStored = 0
  C602ValueStored = 0
  C501ValueStored = 0
  C502ValueStored = 0
 ElseIf TotNumActiveChannels = 4 Then
  C901ValueStored = 1
  C902ValueStored = 1
  C801ValueStored = 1
  C802ValueStored = 1
  C701ValueStored = 0
  C702ValueStored = 0
  C601ValueStored = 0
  C602ValueStored = 0
  C501ValueStored = 0
  C502ValueStored = 0
 ElseIf TotNumActiveChannels = 6 Then
  C901ValueStored = 1
  C902ValueStored = 1
  C801ValueStored = 1
  C802ValueStored = 1
  C701ValueStored = 1
  C702ValueStored = 1
  C601ValueStored = 0
  C602ValueStored = 0
  C501ValueStored = 0
  C502ValueStored = 0
 ElseIf TotNumActiveChannels = 8 Then
  C901ValueStored = 1
  C902ValueStored = 1
  C801ValueStored = 1
  C802ValueStored = 1
  C701ValueStored = 1
  C702ValueStored = 1
  C601ValueStored = 1
  C602ValueStored = 1
  C501ValueStored = 0
  C502ValueStored = 0
 ElseIf TotNumActiveChannels = 10 Then
  C901ValueStored = 1
  C902ValueStored = 1
  C801ValueStored = 1
  C802ValueStored = 1
  C701ValueStored = 1
  C702ValueStored = 1
  C601ValueStored = 1
  C602ValueStored = 1
  C501ValueStored = 1
  C502ValueStored = 1
 ElseIf TotNumActiveChannels < 0 Then
  C901ValueStored = 0
  C902ValueStored = 0
  C801ValueStored = 0
  C802ValueStored = 0
  C701ValueStored = 0
  C702ValueStored = 0
  C601ValueStored = 0
  C602ValueStored = 0
  C501ValueStored = 0
  C502ValueStored = 0
 End If
End If
'-----------------------------------
'Show frames only for # of DUTs user specified
Dim MaxPossibleDUTstoStress As Integer
'Get # of DUTs to be stressed from user input
MaxPossibleDUTstoStress = Val(WGFMUVBInput.NumDevWGFMU.Text)
For i = 0 To (MaxPossibleDUTstoStress - 1)
 FrameforDUTInfo(i).Visible = True
Next
For i = MaxPossibleDUTstoStress To 8
 FrameforDUTInfo(i).Visible = False
Next
'If more than 3 DUTs were specified, then only shows 3 DUTs
'for the all WGFMU channel option
If MaxPossibleDUTstoStress > 3 Then
 For i = 3 To 8
  FrameforDUTInfo(i).Visible = False
 Next
End If
'-----------------------------------
'Set Default channel numbers (based on the assumption that
'full 10 WGFMU channels are available) --> max up to 3 DUTs can be stressed with 10 WGFMU channels
'Slot number              = 9/8/7/6/5/4/3/2/1/-1 corresponds to
'Index number of combobox = 0/1/2/3/4/5/6/7/8/9
'DUT #1
SlotDUT(0).ListIndex = 0 'G slot=9
SlotDUT(1).ListIndex = 0 'D slot=9
SlotDUT(2).ListIndex = 1 'S slot=8
SlotDUT(3).ListIndex = 4 'B slot=5
ChnDUT(0).ListIndex = 0  'G channel = 1
ChnDUT(1).ListIndex = 1  'D channel = 2
ChnDUT(2).ListIndex = 0  'S channel = 1
ChnDUT(3).ListIndex = 1  'B channel = 2
'DUT #2
SlotDUT(4).ListIndex = 1 'G slot=8
SlotDUT(5).ListIndex = 2 'D slot=7
SlotDUT(6).ListIndex = 2 'S slot=7
SlotDUT(7).ListIndex = 4 'B slot=5
ChnDUT(4).ListIndex = 1  'G channel = 2
ChnDUT(5).ListIndex = 0  'D channel = 1
ChnDUT(6).ListIndex = 1  'S channel = 2
ChnDUT(7).ListIndex = 1  'B channel = 2
'DUT #3
SlotDUT(8).ListIndex = 3 'G slot=6
SlotDUT(9).ListIndex = 3 'D slot=6
SlotDUT(10).ListIndex = 4 'S slot=5
SlotDUT(11).ListIndex = 4 'B slot=5
ChnDUT(8).ListIndex = 0  'G channel = 1
ChnDUT(9).ListIndex = 1  'D channel = 2
ChnDUT(10).ListIndex = 0  'S channel = 1
ChnDUT(11).ListIndex = 1  'B channel = 2
'DUT #4
'SlotDUT(12).ListIndex = 6 'G slot=3
'SlotDUT(13).ListIndex = 6 'D slot=3
'SlotDUT(14).ListIndex = 7 'S slot=2
'SlotDUT(15).ListIndex = 7 'B slot=2
'ChnDUT(12).ListIndex = 0  'G channel = 1
'ChnDUT(13).ListIndex = 1  'D channel = 2
'ChnDUT(14).ListIndex = 0  'S channel = 1
'ChnDUT(15).ListIndex = 1  'B channel = 2
'DUT #5
'SlotDUT(16).ListIndex = 8 'G slot=1
'SlotDUT(17).ListIndex = 8 'D slot=1
'SlotDUT(18).ListIndex = 9 'S slot=-1
'SlotDUT(19).ListIndex = 9 'B slot=-1
'ChnDUT(16).ListIndex = 0  'G channel = 1
'ChnDUT(17).ListIndex = 1  'D channel = 2
'ChnDUT(18).ListIndex = 0  'S channel = 1
'ChnDUT(19).ListIndex = 1  'B channel = 2
'-----------------------------------
C901.value = C901ValueStored
C902.value = C902ValueStored
C801.value = C801ValueStored
C802.value = C802ValueStored
C701.value = C701ValueStored
C702.value = C702ValueStored
C601.value = C601ValueStored
C602.value = C602ValueStored
C501.value = C501ValueStored
C502.value = C502ValueStored
'-----------------------------------
'Reset channel numbers based on the number of channels available
'MsgBox (MaxPossibleTotNumActiveChannels)
If MaxPossibleTotNumActiveChannels = 10 Then

ElseIf MaxPossibleTotNumActiveChannels = 8 Then
ElseIf MaxPossibleTotNumActiveChannels = 6 Then
ElseIf MaxPossibleTotNumActiveChannels = 4 Then
ElseIf MaxPossibleTotNumActiveChannels = 2 Then

End If

'-----------------------------------
'From 2nd time opening up this form, display previously saved values
If NumPopup_ConfigFastBTIStress > 1 Then
'MsgBox (NumPopup_ConfigFastBTIStress)
'MsgBox (TotNumActiveChannels)
 For i = 0 To (4 * (TotNumActiveChannels - 1) - 1)
  SlotDUT(i).ListIndex = Store_Slot_IndexNumber(i)
  ChnDUT(i).ListIndex = Store_Channel_IndexNumber(i)
 Next
End If
'-----------------------------------

ALLWGFMU.value = True
SBSysGND.value = False
UserDefinedOption.value = False

End Sub
Private Sub Form_Unload(Cancel As Integer)
'----------------------------------------
'Save number of active WGFMU channels
StoreTotNumActiveChannels = Val(AvailableWFGMUSlots.Text)
'----------------------------------------
'Store CheckerBox values
C901ValueStored = C901.value
C902ValueStored = C902.value
C801ValueStored = C801.value
C802ValueStored = C802.value
C701ValueStored = C701.value
C702ValueStored = C702.value
C601ValueStored = C601.value
C602ValueStored = C602.value
C501ValueStored = C501.value
C502ValueStored = C502.value
'----------------------------------------
'Save items of ComboBoxes before unloading this form
TotNumActiveChannels = Val(AvailableWFGMUSlots.Text)
 For i = 0 To (4 * (TotNumActiveChannels - 1) - 1)
  Store_Slot_IndexNumber(i) = SlotDUT(i).ListIndex
  Store_Channel_IndexNumber(i) = ChnDUT(i).ListIndex
 Next
'----------------------------------------
'Save channel numbers for each DUT
'---------------------
'DUT #1
IndexInc = 0
'> Channel #1
Channel_1_DUT1 = 100 * Val(SlotDUT(0 + IndexInc).Text) + Val(ChnDUT(0 + IndexInc).Text)
'> Channel #2
Channel_2_DUT1 = 100 * Val(SlotDUT(1 + IndexInc).Text) + Val(ChnDUT(1 + IndexInc).Text)
'> Channel #3
Channel_3_DUT1 = 100 * Val(SlotDUT(2 + IndexInc).Text) + Val(ChnDUT(2 + IndexInc).Text)
'If SlotDUT(2 + IndexInc).Text = -1 Then
' Channel_3_DUT1 = -1
'ElseIf SlotDUT(2 + IndexInc).Text <> -1 And ChnDUT(2 + IndexInc).Text <> -1 Then
' Channel_3_DUT1 = 100 * Val(SlotDUT(2 + IndexInc).Text) + Val(ChnDUT(2 + IndexInc).Text)
'ElseIf SlotDUT(2 + IndexInc).Text <> -1 And ChnDUT(2 + IndexInc).Text = -1 Then
' MsgBox ("For DUT #1," & " WGFMU slot = " & SlotDUT(2 + IndexInc).Text & " is specified for " _
' & "Source terminal. Then, channel number other than SG must be selected. " _
' & "Please correct the issue.")
' Cancel = True  'prevent this form from unloading
'End If
'> Channel #4
Channel_4_DUT1 = 100 * Val(SlotDUT(3 + IndexInc).Text) + Val(ChnDUT(3 + IndexInc).Text)
'If SlotDUT(3 + IndexInc).Text = -1 Then
' Channel_4_DUT1 = -1
'ElseIf SlotDUT(3 + IndexInc).Text <> -1 And ChnDUT(3 + IndexInc).Text <> -1 Then
' Channel_4_DUT1 = 100 * Val(SlotDUT(3 + IndexInc).Text) + Val(ChnDUT(3 + IndexInc).Text)
'ElseIf SlotDUT(3 + IndexInc).Text <> -1 And ChnDUT(3 + IndexInc).Text = -1 Then
' MsgBox ("For DUT #1," & " WGFMU slot = " & SlotDUT(3 + IndexInc).Text & " is specified for " _
' & "Body terminal. Then, channel number other than SG must be selected. " _
' & "Please correct the issue.")
' Cancel = True  'prevent this form from unloading
'End If
'---------------------
'DUT #2
IndexInc = 4
'> Channel #1
Channel_1_DUT2 = 100 * Val(SlotDUT(0 + IndexInc).Text) + Val(ChnDUT(0 + IndexInc).Text)
'> Channel #2
Channel_2_DUT2 = 100 * Val(SlotDUT(1 + IndexInc).Text) + Val(ChnDUT(1 + IndexInc).Text)
'> Channel #3
Channel_3_DUT2 = 100 * Val(SlotDUT(2 + IndexInc).Text) + Val(ChnDUT(2 + IndexInc).Text)
'If SlotDUT(2 + IndexInc).Text = -1 Then
' Channel_3_DUT2 = -1
'ElseIf SlotDUT(2 + IndexInc).Text <> -1 And ChnDUT(2 + IndexInc).Text <> -1 Then
' Channel_3_DUT2 = 100 * Val(SlotDUT(2 + IndexInc).Text) + Val(ChnDUT(2 + IndexInc).Text)
'ElseIf SlotDUT(2 + IndexInc).Text <> -1 And ChnDUT(2 + IndexInc).Text = -1 Then
' MsgBox ("For DUT #2," & " WGFMU slot = " & SlotDUT(2 + IndexInc).Text & " is specified for " _
' & "Source terminal. Then, channel number other than SG must be selected. " _
' & "Please correct the issue.")
' Cancel = True  'prevent this form from unloading
'End If
'> Channel #4
Channel_4_DUT2 = 100 * Val(SlotDUT(3 + IndexInc).Text) + Val(ChnDUT(3 + IndexInc).Text)
'If SlotDUT(3 + IndexInc).Text = -1 Then
' Channel_4_DUT2 = -1
'ElseIf SlotDUT(3 + IndexInc).Text <> -1 And ChnDUT(3 + IndexInc).Text <> -1 Then
' Channel_4_DUT2 = 100 * Val(SlotDUT(3 + IndexInc).Text) + Val(ChnDUT(3 + IndexInc).Text)
'ElseIf SlotDUT(3 + IndexInc).Text <> -1 And ChnDUT(3 + IndexInc).Text = -1 Then
' MsgBox ("For DUT #2," & " WGFMU slot = " & SlotDUT(3 + IndexInc).Text & " is specified for " _
' & "Body terminal. Then, channel number other than SG must be selected. " _
' & "Please correct the issue.")
' Cancel = True  'prevent this form from unloading
'End If
'---------------------
'DUT #3
IndexInc = 8
'> Channel #1
Channel_1_DUT3 = 100 * Val(SlotDUT(0 + IndexInc).Text) + Val(ChnDUT(0 + IndexInc).Text)
'> Channel #2
Channel_2_DUT3 = 100 * Val(SlotDUT(1 + IndexInc).Text) + Val(ChnDUT(1 + IndexInc).Text)
'> Channel #3
Channel_3_DUT3 = 100 * Val(SlotDUT(2 + IndexInc).Text) + Val(ChnDUT(2 + IndexInc).Text)
'If SlotDUT(2 + IndexInc).Text = -1 Then
' Channel_3_DUT3 = -1
'ElseIf SlotDUT(2 + IndexInc).Text <> -1 And ChnDUT(2 + IndexInc).Text <> -1 Then
' Channel_3_DUT3 = 100 * Val(SlotDUT(2 + IndexInc).Text) + Val(ChnDUT(2 + IndexInc).Text)
'ElseIf SlotDUT(2 + IndexInc).Text <> -1 And ChnDUT(2 + IndexInc).Text = -1 Then
' MsgBox ("For DUT #3," & " WGFMU slot = " & SlotDUT(2 + IndexInc).Text & " is specified for " _
' & "Source terminal. Then, channel number other than SG must be selected. " _
' & "Please correct the issue.")
' Cancel = True  'prevent this form from unloading
'End If
'> Channel #4
Channel_4_DUT3 = 100 * Val(SlotDUT(3 + IndexInc).Text) + Val(ChnDUT(3 + IndexInc).Text)
'If SlotDUT(3 + IndexInc).Text = -1 Then
' Channel_4_DUT3 = -1
'ElseIf SlotDUT(3 + IndexInc).Text <> -1 And ChnDUT(3 + IndexInc).Text <> -1 Then
' Channel_4_DUT3 = 100 * Val(SlotDUT(3 + IndexInc).Text) + Val(ChnDUT(3 + IndexInc).Text)
'ElseIf SlotDUT(3 + IndexInc).Text <> -1 And ChnDUT(3 + IndexInc).Text = -1 Then
' MsgBox ("For DUT #3," & " WGFMU slot = " & SlotDUT(3 + IndexInc).Text & " is specified for " _
' & "Body terminal. Then, channel number other than SG must be selected. " _
' & "Please correct the issue.")
' Cancel = True  'prevent this form from unloading
'End If
'---------------------
'DUT #4
IndexInc = 12
'> Channel #1
Channel_1_DUT4 = 100 * Val(SlotDUT(0 + IndexInc).Text) + Val(ChnDUT(0 + IndexInc).Text)
'> Channel #2
Channel_2_DUT4 = 100 * Val(SlotDUT(1 + IndexInc).Text) + Val(ChnDUT(1 + IndexInc).Text)
'> Channel #3
Channel_3_DUT4 = 100 * Val(SlotDUT(2 + IndexInc).Text) + Val(ChnDUT(2 + IndexInc).Text)
'If SlotDUT(2 + IndexInc).Text = -1 Then
' Channel_3_DUT4 = -1
'ElseIf SlotDUT(2 + IndexInc).Text <> -1 And ChnDUT(2 + IndexInc).Text <> -1 Then
' Channel_3_DUT4 = 100 * Val(SlotDUT(2 + IndexInc).Text) + Val(ChnDUT(2 + IndexInc).Text)
'ElseIf SlotDUT(2 + IndexInc).Text <> -1 And ChnDUT(2 + IndexInc).Text = -1 Then
' MsgBox ("For DUT #4," & " WGFMU slot = " & SlotDUT(2 + IndexInc).Text & " is specified for " _
' & "Source terminal. Then, channel number other than SG must be selected. " _
' & "Please correct the issue.")
' Cancel = True  'prevent this form from unloading
'End If
'> Channel #4
Channel_4_DUT4 = 100 * Val(SlotDUT(3 + IndexInc).Text) + Val(ChnDUT(3 + IndexInc).Text)
'If SlotDUT(3 + IndexInc).Text = -1 Then
' Channel_4_DUT4 = -1
'ElseIf SlotDUT(3 + IndexInc).Text <> -1 And ChnDUT(3 + IndexInc).Text <> -1 Then
' Channel_4_DUT4 = 100 * Val(SlotDUT(3 + IndexInc).Text) + Val(ChnDUT(3 + IndexInc).Text)
'ElseIf SlotDUT(3 + IndexInc).Text <> -1 And ChnDUT(3 + IndexInc).Text = -1 Then
' MsgBox ("For DUT #4," & " WGFMU slot = " & SlotDUT(3 + IndexInc).Text & " is specified for " _
' & "Body terminal. Then, channel number other than SG must be selected. " _
' & "Please correct the issue.")
' Cancel = True  'prevent this form from unloading
'End If
'---------------------
'DUT #5
IndexInc = 16
'> Channel #1
Channel_1_DUT5 = 100 * Val(SlotDUT(0 + IndexInc).Text) + Val(ChnDUT(0 + IndexInc).Text)
'> Channel #2
Channel_2_DUT5 = 100 * Val(SlotDUT(1 + IndexInc).Text) + Val(ChnDUT(1 + IndexInc).Text)
'> Channel #3
Channel_3_DUT5 = 100 * Val(SlotDUT(2 + IndexInc).Text) + Val(ChnDUT(2 + IndexInc).Text)
'If SlotDUT(2 + IndexInc).Text = -1 Then
' Channel_3_DUT5 = -1
'ElseIf SlotDUT(2 + IndexInc).Text <> -1 And ChnDUT(2 + IndexInc).Text <> -1 Then
' Channel_3_DUT5 = 100 * Val(SlotDUT(2 + IndexInc).Text) + Val(ChnDUT(2 + IndexInc).Text)
'ElseIf SlotDUT(2 + IndexInc).Text <> -1 And ChnDUT(2 + IndexInc).Text = -1 Then
' MsgBox ("For DUT #5," & " WGFMU slot = " & SlotDUT(2 + IndexInc).Text & " is specified for " _
' & "Source terminal. Then, channel number other than SG must be selected. " _
' & "Please correct the issue.")
' Cancel = True  'prevent this form from unloading'
'End If
'> Channel #4
Channel_4_DUT5 = 100 * Val(SlotDUT(3 + IndexInc).Text) + Val(ChnDUT(3 + IndexInc).Text)
'If SlotDUT(3 + IndexInc).Text = -1 Then
' Channel_4_DUT5 = -1
'ElseIf SlotDUT(3 + IndexInc).Text <> -1 And ChnDUT(3 + IndexInc).Text <> -1 Then
' Channel_4_DUT5 = 100 * Val(SlotDUT(3 + IndexInc).Text) + Val(ChnDUT(3 + IndexInc).Text)
'ElseIf SlotDUT(3 + IndexInc).Text <> -1 And ChnDUT(3 + IndexInc).Text = -1 Then
' MsgBox ("For DUT #5," & " WGFMU slot = " & SlotDUT(3 + IndexInc).Text & " is specified for " _
' & "Body terminal. Then, channel number other than SG must be selected. " _
' & "Please correct the issue.")
' Cancel = True  'prevent this form from unloading
'End If
'----------------------------------------
'Enable Run button on the main form
WGFMUVBInput.RunButton.Enabled = True
'----------------------------------------
'Allow changes on the main form As Config form closes
WGFMUVBInput.Enabled = True
'----------------------------------------
End Sub

Private Sub getWGFMUChannels_Click()
On Error GoTo HandleErrors
'Reset WGFMU and get the total number of available channels
Dim GPIBString As String
Dim openSessioStatus As Long
Dim InitializationStatus As Long
Dim CloseSessionStatus As Long
'Initialize GPIB
GPIBAdressNumber = 18
GPIBString = "GPIB0::" & CStr(GPIBAdressNumber) & "::INSTR"
openSessioStatus = WGFMU_openSession(GPIBString)
InitializationStatus = WGFMU_initialize()
'Get the total number of available WGFMU channels
Dim NumChannelsAvailable As Long
Dim NumChannelsAvailableStatus As Long
NumChannelsAvailableStatus = WGFMU_getChannelIdSize(NumChannelsAvailable)
CloseSessionStatus = WGFMU_closeSession()
'Display number of active channels
AvailableWFGMUSlots.Text = NumChannelsAvailable
StoreTotNumActiveChannels = Val(AvailableWFGMUSlots.Text)
If (AvailableWFGMUSlots.Text < 0) Then
MsgBox ("WARNING: Error occured during trying to count active number of WGFMU channels. Please, check " _
& "the remote connection to B1500 via GPIB interface and try it again. Without being " _
& "remotely connected to B1500, fast BTI stress cannot be performed!")
End If

HandleErrors:
If Val(AvailableWFGMUSlots.Text) > 0 Then
 Exit Sub
Else
 MsgBox ("GPIB communication failed. Please, check and make sure GPIB is active and try again!")
End If
End Sub

Private Sub SBSysGND_Click()
'------------------------------------------------
'Slot number              = 9/8/7/6/5/4/3/2/1/-1 corresponds to
'Index number of combobox = 0/1/2/3/4/5/6/7/8/9
 'DUT #1
 SlotDUT(0).ListIndex = 0 'G slot=9
 SlotDUT(1).ListIndex = 0 'D slot=9
 SlotDUT(2).ListIndex = 9 'S slot=-1
 SlotDUT(3).ListIndex = 9 'B slot=-1
 ChnDUT(0).ListIndex = 0  'G channel = 1
 ChnDUT(1).ListIndex = 1  'D channel = 2
 ChnDUT(2).ListIndex = 2  'S channel = -1
 ChnDUT(3).ListIndex = 2  'B channel = -1
 'DUT #2
 SlotDUT(4).ListIndex = 1 'G slot=8
 SlotDUT(5).ListIndex = 1 'D slot=8
 SlotDUT(6).ListIndex = 9 'S slot=-1
 SlotDUT(7).ListIndex = 9 'B slot=-1
 ChnDUT(4).ListIndex = 0  'G channel = 1
 ChnDUT(5).ListIndex = 1  'D channel = 2
 ChnDUT(6).ListIndex = 2  'S channel = -1
 ChnDUT(7).ListIndex = 2  'B channel = -1
 'DUT #3
 SlotDUT(8).ListIndex = 2 'G slot=7
 SlotDUT(9).ListIndex = 2 'D slot=7
 SlotDUT(10).ListIndex = 9 'S slot=-1
 SlotDUT(11).ListIndex = 9 'B slot=-1
 ChnDUT(8).ListIndex = 0  'G channel = 1
 ChnDUT(9).ListIndex = 1  'D channel = 2
 ChnDUT(10).ListIndex = 2  'S channel = -1
 ChnDUT(11).ListIndex = 2  'B channel = -1
 'DUT #4
 SlotDUT(12).ListIndex = 3 'G slot=6
 SlotDUT(13).ListIndex = 3 'D slot=6
 SlotDUT(14).ListIndex = 9 'S slot=-1
 SlotDUT(15).ListIndex = 9 'B slot=-1
 ChnDUT(12).ListIndex = 0  'G channel = 1
 ChnDUT(13).ListIndex = 1  'D channel = 2
 ChnDUT(14).ListIndex = 2  'S channel = -1
 ChnDUT(15).ListIndex = 2  'B channel = -1
 'DUT #5
 SlotDUT(16).ListIndex = 4 'G slot=5
 SlotDUT(17).ListIndex = 4 'D slot=5
 SlotDUT(18).ListIndex = 9 'S slot=-1
 SlotDUT(19).ListIndex = 9 'B slot=-1
 ChnDUT(16).ListIndex = 0  'G channel = 1
 ChnDUT(17).ListIndex = 1  'D channel = 2
 ChnDUT(18).ListIndex = 2  'S channel = -1
 ChnDUT(19).ListIndex = 2  'B channel = -1
'------------------------------------------------
'Make all 5 DUT frames visible
For i = 0 To 4
 FrameforDUTInfo(i).Visible = True
Next
'------------------------------------------------
End Sub

Private Sub UserDefinedOption_Click()
'--------------------------------
'Make all 5 DUT frames visible
For i = 0 To 4
 FrameforDUTInfo(i).Visible = True
Next
'--------------------------------
'reset all listbox items empty
For i = 0 To 19
 SlotDUT(i).ListIndex = -1
 ChnDUT(i).ListIndex = -1
 Next
End Sub
