VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDM_Interface 
   Caption         =   "Dungeon Master"
   ClientHeight    =   11310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   ScaleHeight     =   11310
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMinimap 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2145
      ScaleWidth      =   3345
      TabIndex        =   13
      Top             =   7920
      Width           =   3375
      Begin VB.Image imgMinimap 
         Height          =   855
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1215
      End
      Begin VB.Shape shpFocus 
         Height          =   615
         Index           =   0
         Left            =   2160
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   120
      ScaleHeight     =   7065
      ScaleWidth      =   10905
      TabIndex        =   12
      Top             =   480
      Width           =   10935
      Begin VB.Line linGridy 
         Index           =   0
         X1              =   2400
         X2              =   2400
         Y1              =   2040
         Y2              =   3240
      End
      Begin VB.Line linGridx 
         Index           =   0
         X1              =   2400
         X2              =   3720
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Image imgMap 
         Height          =   1095
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.TextBox txtLog 
      Height          =   2175
      Left            =   8280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "DM_Interface.frx":0000
      Top             =   7920
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      Height          =   2295
      Left            =   11400
      TabIndex        =   9
      Top             =   7800
      Width           =   3375
   End
   Begin VB.ComboBox cboSort 
      Height          =   315
      Left            =   11880
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   0
      Width           =   2535
   End
   Begin VB.VScrollBar vscPlayers 
      Height          =   7455
      Left            =   14520
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.Frame fraPlayers 
      Height          =   7575
      Left            =   11400
      TabIndex        =   2
      Top             =   240
      Width           =   3015
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   570
         Index           =   0
         Left            =   50
         TabIndex        =   6
         Top             =   130
         Width           =   2910
         Begin VB.Label lblRaceclass 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Race Class"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   705
            TabIndex        =   10
            Top             =   405
            Width           =   1935
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            Height          =   465
            Index           =   0
            Left            =   2710
            Top             =   50
            Width           =   135
         End
         Begin VB.Shape Shape2 
            Height          =   495
            Index           =   0
            Left            =   2700
            Top             =   30
            Width           =   160
         End
         Begin VB.Label lblPlayername 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Player Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   705
            TabIndex        =   8
            Top             =   270
            Width           =   1695
         End
         Begin VB.Label lblCharname 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Caption         =   "Character Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   0
            Left            =   705
            TabIndex        =   7
            Top             =   60
            Width           =   2025
         End
         Begin VB.Image Image2 
            Height          =   570
            Index           =   0
            Left            =   0
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   615
         Index           =   0
         Left            =   30
         Top             =   120
         Width           =   2950
      End
   End
   Begin VB.TextBox txtDialogue 
      Height          =   2175
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "DM_Interface.frx":0006
      Top             =   7920
      Width           =   4695
   End
   Begin MSComctlLib.TabStrip tabMap 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   13785
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "No Map"
            Object.Tag             =   "0"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "+"
            Object.Tag             =   "+"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   10200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSort 
      AutoSize        =   -1  'True
      Caption         =   "Sort:"
      Height          =   195
      Left            =   11400
      TabIndex        =   4
      Top             =   50
      Width           =   330
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu popMap 
      Caption         =   "Invisible"
      Visible         =   0   'False
      Begin VB.Menu popLoadmap 
         Caption         =   "Load Map"
      End
      Begin VB.Menu popClosemap 
         Caption         =   "Close Map"
         Enabled         =   0   'False
      End
      Begin VB.Menu popBar1 
         Caption         =   "-"
      End
      Begin VB.Menu popDrawgrid 
         Caption         =   "Draw Grid"
         Enabled         =   0   'False
      End
      Begin VB.Menu popGridproperties 
         Caption         =   "Grid Properties"
      End
   End
End
Attribute VB_Name = "frmDM_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
With Me
    .Width = 14985
    .Height = 10890
End With
With tabMap
    .Left = 0
    .Top = 0
    .Width = 11295
    .Height = 7815
End With
With picMinimap
    .Left = 0
    .Top = 7920
    .Height = 2175
    .Width = 3375
End With
With imgMinimap(0)
    .Left = 0
    .Top = 0
    .Height = 2175
    .Width = 3375
    .Visible = False
End With
With txtDialogue
    .Left = 3480
    .Top = 7920
    .Height = 2175
    .Width = 4695
End With
With txtLog
    .Left = 8280
    .Top = 7920
    .Width = 3015
    .Height = 2175
End With
With lblSort
    .Left = 11400
    .Top = 50
    .Width = 330
    .Height = 195
End With
With cboSort
    .Left = 11880
    .Top = 0
    .Width = 2535
End With
With fraPlayers
    .Left = 11400
    .Top = 240
    .Width = 3015
    .Height = 7575
End With
With Frame3
    .Left = 11400
    .Top = 7800
    .Width = 3375
    .Height = 2295
End With
With vscPlayers
    .Left = 14520
    .Top = 360
    .Width = 255
    .Height = 7455
End With
With Shape1(0)
    .Left = 30
    .Top = 120
    .Width = 2950
    .Height = 615
    .Visible = False
End With
With Image2(0)
    .Left = 0
    .Top = 0
    .Width = 615
    .Height = 570
    .Visible = False
End With
With lblCharname(0)
    .Left = 705
    .Top = 60
    .Width = 2025
    .Height = 300
    .Visible = False
End With
With lblPlayername(0)
    .Left = 705
    .Top = 270
    .Width = 1695
    .Height = 255
    .Visible = False
End With
With lblRaceclass(0)
    .Left = 705
    .Top = 405
    .Width = 1935
    .Height = 255
    .Visible = False
End With
With Shape2(0)
    .Left = 2700
    .Top = 30
    .Width = 160
    .Height = 495
    .Visible = False
End With
With Shape3(0)
    .Left = 2710
    .Top = 50
    .Width = 135
    .Height = 460
    .Visible = False
End With
With picMap
    .Left = 120
    .Top = 360
    .Width = 11055
    .Height = 7335
    .Visible = False
End With
With Frame2(0)
    .Left = 50
    .Top = 130
    .Height = 570
    .Width = 2910
    .Visible = False
End With
With shpFocus(0)
    .Left = picMinimap.Left
    .Top = picMinimap.Top
    .Width = 100
    .Height = 100
    .Visible = False
End With
With imgMap(0)
    .Left = 0
    .Top = 0
    .Visible = False
    .Stretch = False
End With
'Locate reference grid lines and make them invisible
With linGridx(0)
    .X1 = 0
    .X2 = 500
    .Y1 = 0
    .Y2 = 0
    .Visible = False
End With
With linGridy(0)
    .X1 = 0
    .X2 = 0
    .Y1 = 0
    .Y2 = 500
    .Visible = False
End With
'
''Set grid properties
'    SetGridProperties 375, &H80000010, vbBSDot, 1
'
''Set default grid size
'    grid_size = 375



End Sub

Private Sub imgMinimap_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picMinimap_MouseDown Button, Shift, X + imgMinimap(Index).Left, Y + imgMinimap(Index).Top
End Sub

Private Sub imgMinimap_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
picMinimap_MouseMove Button, Shift, X + imgMinimap(Index).Left, Y + imgMinimap(Index).Top
End Sub

Private Sub tabMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then      'This user right clicked
    If tabMap.SelectedItem.Tag <> 0 And tabMap.SelectedItem.Tag <> "+" Then   'This means that selected tab has an associated map
        popClosemap.Enabled = True              'Enable Close map popup menu
    ElseIf tabMap.SelectedItem.Tag = 0 Then     'This means that selected tab has no associated map
        popClosemap.Enabled = False             'Disable Close map popup menu
    End If
    PopupMenu popMap          'Open pop up menu
End If
End Sub

Private Sub tabMap_Click()
Dim i As Integer
Dim iMaxtab As Integer
Dim iSelected As Integer

iMaxtab = 6    'Maximum 6 tabs are allowed
iSelected = tabMap.SelectedItem.Index

'If a map tab is clicked, open the corresponding map. Update minimap
'accordingly.
If tabMap.Tabs(iSelected).Tag <> "+" Then   'This means that a map tab is clicked
    For i = 1 To tabMap.Tabs.Count      'Make all maps,minimaps, focus rectangles invisible so that nothing is shown
        If tabMap.Tabs(i).Tag <> "+" Then       'This means that tab is a map tab and has an associated map
            imgMap(tabMap.SelectedItem.Tag).Visible = False      'Make map invisible
            imgMinimap(tabMap.SelectedItem.Tag).Visible = False  'Make minimap invisible
            shpFocus(tabMap.SelectedItem.Tag).Visible = False    'Make focus rectangle invisible
        End If
    Next i
    If tabMap.Tabs(iSelected).Tag <> 0 Then     'This means that a map is loaded in clicked tab
        imgMap(tabMap.SelectedItem.Tag).Visible = True  'Make map of selected tab visible
        imgMinimap(tabMap.SelectedItem.Tag).Visible = True  'Make minimap of selected tab visible
        shpFocus(tabMap.SelectedItem.Tag).Visible = True    'Make focus rectangle of selected tab visible
    End If
End If

'If + tab is clicked and max number of tabs is not reached then add a new tab
If tabMap.SelectedItem.Tag = "+" And tabMap.Tabs.Count <= iMaxtab Then    'This means that + tab is clicked and maximum number of tabs is not reached
    tabMap.Tabs.Add                         'Create new tab. Tab count increased by one
    tabMap.Tabs(tabMap.Tabs.Count).Caption = "+"    'Change caption of last tab to +
    tabMap.Tabs(tabMap.Tabs.Count).Tag = "+"        'Change tag of last tab to +
    tabMap.Tabs(tabMap.Tabs.Count - 1).Caption = "No Map"       'Change clicked + tabs caption to no map
    tabMap.Tabs(tabMap.Tabs.Count - 1).Tag = 0                  'Change clicked + tabs tag to indicate no map
    tabMap.SelectedItem = tabMap.Tabs(tabMap.Tabs.Count - 1)    'Select newly created map tab
    tabMap_Click                                                'Focus on newly created tab so that map changes
End If

'When max number of tabs is reached remove + tab at the end so that no more map tabs can be added
If tabMap.Tabs.Count = iMaxtab + 1 Then     'This means there are max number of map tabs and still + tab at the end
    tabMap.Tabs.Remove (iMaxtab + 1)        'Close the last tab which is + tab
    tabMap.SelectedItem = tabMap.Tabs(iMaxtab)  'Select last tab
    tabMap_Click                                'Focus on last tab so that map changes
End If

If tabMap.Tabs(iSelected).Tag <> 0 And tabMap.Tabs(iSelected).Tag <> "+" Then    'This means that a map is associated with clicked tab
    popDrawgrid.Enabled = True
    popClosemap.Enabled = True
ElseIf tabMap.Tabs(iSelected).Tag = 0 Then      'This means no map is associated with selected map
    popDrawgrid.Enabled = False
    popClosemap.Enabled = False
End If

End Sub

Private Sub popClosemap_Click()
'Closes a map tab
Dim iSelected As Integer

iSelected = tabMap.SelectedItem.Index

'If only one map is open dont close the tab. Just
'remove map only
If tabMap.Tabs.Count <= 2 Then  'This means only one map is open (The other tab is + tab)
    If tabMap.Tabs(iSelected).Tag <> 0 Then   'This means a map is loaded in the tab
        Unload imgMap(tabMap.Tabs(iSelected).Tag)          'Unload map related to closed tab
        Unload imgMinimap(tabMap.Tabs(iSelected).Tag)      'Unload minimap related to closed tab
        Unload shpFocus(tabMap.Tabs(iSelected).Tag)        'Unload focus rectangle related to closed tab
        tabMap.Tabs(iSelected).Tag = 0              'Clear tag of tab since map is unloaded
        tabMap.Tabs(iSelected).Caption = "No Map"   'Change title of tab since map is unloaded
    End If
Else                            'This means more than one map is open. Both unload map, if a map is loaded, and close the tab
    If tabMap.Tabs(iSelected).Tag <> 0 Then   'This means a map is loaded in the tab
        Unload imgMap(tabMap.Tabs(iSelected).Tag)          'Unload map related to closed tab
        Unload imgMinimap(tabMap.Tabs(iSelected).Tag)      'Unload minimap related to closed tab
        Unload shpFocus(tabMap.Tabs(iSelected).Tag)        'Unload focus rectangle related to closed tab
    End If
    tabMap.Tabs.Remove iSelected                'Close the selected tab
    If iSelected = 1 Then       'This means first tab is closed
        tabMap.SelectedItem = tabMap.Tabs(1)    'Select first tab
    ElseIf iSelected > 1 Then
        tabMap.SelectedItem = tabMap.Tabs(iSelected - 1)    'Select previous tab from the closed tab
    End If
End If
tabMap_Click                                        'Focus on previous tab so that map changes

'If maximum number of tabs were reached, then there
'was no + tab when tab was closed. Since + tab is always
'the last tab, look if last tab is + tab. If not add it
If tabMap.Tabs(tabMap.Tabs.Count).Tag <> "+" Then   'This means last tab is not + tab
    tabMap.Tabs.Add                                 'Add a new tab. Tab count increased by one
    tabMap.Tabs(tabMap.Tabs.Count).Caption = "+"    'Change title of tab to +
    tabMap.Tabs(tabMap.Tabs.Count).Tag = "+"        'Change tag of tab to +
Else                                                'This means that last tab is + tab
    'Do nothing
End If

End Sub

Private Sub popDrawgrid_Click()
DrawGrid 500
End Sub

Private Sub popLoadmap_Click()
'Loads a map to selected tab from map file
Dim fs As Object
Dim f As Object
Dim loadfile As Object
Dim i As Integer
Dim dScale As Double

CD.FileName = ""
CD.Filter = "Map(*.map)|*.map|All Files(*.*)|*.*"
CD.ShowOpen
If CD.FileName <> "" Then 'a file is selected
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(CD.FileName)
    Set loadfile = f.OpenAsTextStream(1, -2)
    
    'If a map is already open, remove it first
    If tabMap.Tabs(tabMap.SelectedItem.Index).Tag <> 0 Then
        Unload imgMap(tabMap.SelectedItem.Tag)
        Unload imgMinimap(tabMap.SelectedItem.Tag)
        Unload shpFocus(tabMap.SelectedItem.Tag)
    End If
    
    'Search for an empty index. Empty index can either be
    'found inside lbound and ubound due to a removed map, or
    'after the last index, which is ubound+1
    Do While i <= imgMap.UBound + 1
        If imgMap(i) Is Nothing Then
            Exit Do
        End If
        i = i + 1
    Loop
    'Associate map to tab using tag of tab. In order to determine the index of
    'a map in a tab, just look at the tag of the tab
    tabMap.Tabs(tabMap.SelectedItem.Index).Tag = i
    Load imgMap(tabMap.SelectedItem.Tag)
    Load imgMinimap(tabMap.SelectedItem.Tag)
    Load shpFocus(tabMap.SelectedItem.Tag)
    With picMap
        .ZOrder 0
        .Visible = True
    End With
    With imgMap(tabMap.SelectedItem.Tag)
        Set .Container = picMap
        .Picture = LoadPicture(CD.FileName)
        .Visible = True
    End With

    tabMap.Tabs(tabMap.SelectedItem.Index).Caption = CD.FileTitle   'Change title of tab
    popDrawgrid.Enabled = True     'Enable apply grid pop up menu since a map is loaded
       
        
    'Load picture to minimap also and scale it
    imgMinimap(tabMap.SelectedItem.Tag).Picture = LoadPicture(CD.FileName)
    If imgMap(tabMap.SelectedItem.Tag).Height / picMinimap.Height >= imgMap(tabMap.SelectedItem.Tag).Width / picMinimap.Width Then
        dScale = imgMap(tabMap.SelectedItem.Tag).Height / picMinimap.Height
    Else
        dScale = imgMap(tabMap.SelectedItem.Tag).Width / picMinimap.Width
    End If
    With imgMinimap(tabMap.SelectedItem.Tag)
        .Width = imgMap(tabMap.SelectedItem.Tag).Width / dScale
        .Height = imgMap(tabMap.SelectedItem.Tag).Height / dScale
        .Left = picMinimap.Width / 2 - imgMinimap(tabMap.SelectedItem.Tag).Width / 2
        .Top = picMinimap.Height / 2 - imgMinimap(tabMap.SelectedItem.Tag).Height / 2
        .Visible = True
    End With
    
    'Place focus rectangle
    With shpFocus(tabMap.SelectedItem.Tag)
        .Width = imgMinimap(tabMap.SelectedItem.Tag).Width * picMap.Width / imgMap(tabMap.SelectedItem.Tag).Width
        .Height = imgMinimap(tabMap.SelectedItem.Tag).Height * picMap.Height / imgMap(tabMap.SelectedItem.Tag).Height
        .Left = imgMinimap(tabMap.SelectedItem.Tag).Left
        .Top = imgMinimap(tabMap.SelectedItem.Tag).Top
        .Visible = True
        .ZOrder 0
    End With

    If shpFocus(tabMap.SelectedItem.Tag).Width >= imgMinimap(tabMap.SelectedItem.Tag).Width And shpFocus(tabMap.SelectedItem.Tag).Height >= imgMinimap(tabMap.SelectedItem.Tag).Height Then
        shpFocus(tabMap.SelectedItem.Tag).Visible = False
    End If
Else
    'Do nothing
End If

End Sub

Private Sub picMinimap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nx As Integer
Dim ny As Integer
Dim n As Integer

n = tabMap.SelectedItem.Tag
If Button = 1 And imgMinimap(n).Picture <> 0 And imgMinimap(n).Visible = True And shpFocus(n).Visible = True Then
    nx = X - shpFocus(n).Width / 2
    ny = Y - shpFocus(n).Height / 2
    If nx < 0 Then
        nx = 0
    End If
    If nx + shpFocus(n).Width > picMinimap.Width Then
        nx = picMinimap.Width - shpFocus(n).Width
    End If
    If ny < 0 Then
        ny = 0
    End If
    If ny + shpFocus(n).Height > picMinimap.Height Then
        ny = picMinimap.Height - shpFocus(n).Height
    End If
    Call shpFocus(n).Move(nx, ny)
    Call imgMap(n).Move(-(nx - imgMinimap(n).Left) * picMap.Width / shpFocus(n).Width, -(ny - imgMinimap(n).Top) * picMap.Height / shpFocus(n).Height)
End If
End Sub


Private Sub picMinimap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'When focus window is dragged across minimap, move map accordingly
Dim ny As Integer
Dim nx As Integer
Dim n As Integer

n = tabMap.SelectedItem.Tag
'If left mouse is clicked and a map is loaded
If Button = 1 And imgMinimap(n).Picture <> 0 And imgMinimap(n).Visible = True And shpFocus(n).Visible = True Then
    nx = X - shpFocus(n).Width / 2
    ny = Y - shpFocus(n).Height / 2
    If nx < 0 Then
        nx = 0
    End If
    If nx + shpFocus(n).Width > picMinimap.Width Then
        nx = picMinimap.Width - shpFocus(n).Width
    End If
    If ny < 0 Then
        ny = 0
    End If
    If ny + shpFocus(n).Height > picMinimap.Height Then
        ny = picMinimap.Height - shpFocus(n).Height
    End If
    Call shpFocus(n).Move(nx, ny)
    Call imgMap(n).Move(-(nx - imgMinimap(n).Left) * picMap.Width / shpFocus(n).Width, -(ny - imgMinimap(n).Top) * picMap.Height / shpFocus(n).Height)

''    'Shift grid
''    For i = 0 To linGridX.UBound
'        With linGridX(i)
'            .X1 = .X1 + xshift
'            .X2 = .X2 + xshift
'            .Y1 = .Y1 + yshift
'            .Y2 = .Y2 + yshift
'        End With
'    Next i
'
'    For i = 0 To linGridY.UBound
'        With linGridY(i)
'            .X1 = .X1 + xshift
'            .X2 = .X2 + xshift
'            .Y1 = .Y1 + yshift
'            .Y2 = .Y2 + yshift
'        End With
'    Next i
'    'Shift miniatures
'    For i = 0 To imgMini.UBound
'        imgMini(i).Left = imgMini(i).Left + xshift
'        imgMini(i).Top = imgMini(i).Top + yshift
'    Next i
End If
End Sub

Public Function DrawGrid(gridsize As Integer, mapnum As Integer, _
                         Optional offsetX As Integer = 0, _
                         Optional offsetY As Integer = 0, _
                         Optional gridcolor As ColorConstants = vbBlack, _
                         Optional gridstyle As BorderStyleConstants = vbBSSolid, _
                         Optional gridwidth As Integer = 1)
                         
'Draws square grid patterns on image. Each map has its own grid. For example:
' -Array number 0 is for reference only.
' -Map1 has an applied grid with 100 elements: linGridx(1 to 100), linGridy(1 to 100)
' -Grid is applied to Map2 with 250 elements: linGridx(101 to 351), linGridy(101 to 351)
' Each gridline's tag tells to which map they are attached to

Dim i As Integer
Dim j As Integer
Dim iMapstart As Integer
Dim iMapend As Integer
Dim iLengthX As Integer
Dim iLengthY As Integer
Dim iNewlengthX As Integer
Dim iNewlengthY As Integer
Dim row As Integer
Dim col As Integer
    
    'Offset values cannot be greater than gridsize. It is meaningless
    offsetX = offsetX Mod gridsize
'    offsetY = offsetY Mod gridsize
    
    'First find the space covered by previous grid, if a previous grid was there
    If linGridx.UBound > 0 Then    'This means at least one grid is drawn.
        i = 1
        Do While i <= linGridx.UBound   'Find the first appearance of current map in grid collection
            If linGridx(i).Tag = mapnum Then    'This means grid is a part of current map
                j = 1
                Exit Do
            Else
                j = 0
            End If
            i = i + 1
        Loop
        iMapstart = i
        Do While linGridx(i).Tag = mapnum And i <= linGridx.UBound    'Find the last appearance of current map in grid collection
            i = i + 1
        Loop
        iMapend = i - 1
    End If
    iLength = iMapend - iMapstart + 1   'Length of the portion corresponding to current map
    
    'Now calculate the new length required
    iNewlengthX = ((imgMap(mapnum).Height - offsetX) \ gridsize) + 1
'    iNewlengthY = ((imgMap(mapnum).Width - offsetY) \ gridsize) + 1
    
    'If new length is greater than before, empty some space and offset the collection accordingly
    If iNewlengthX > iLengthX Then      'This means new grid requires more space than before
        If linGridx.UBound > iMapend Then   'This means there are gridlines after the last gridline for this map. So offset them
            j = linGridx.UBound
            For i = 0 To j - iMapend - 1
                Load linGridx(j + iNewlengthX - iLengthX - i)
                With linGridx(j + iNewlengthX - iLengthX - i)
                    .X1 = linGridx(j - i).X1
                    .X2 = linGridx(j - i).X2
                    .Y1 = linGridx(j - i).Y1
                    .Y2 = linGridx(j - i).Y2
                    .BorderColor = linGridx(j - i).BorderColor
                    .BorderStyle = linGridx(j - i).BorderStyle
                    .BorderWidth = linGridx(j - i).BorderWidth
                    .Tag = linGridx(j - i).Tag
                End With
            Next i
    End If
    
    'If new length is greater than before, empty some space and offset the collection accordingly
    
    
    
    'First unload previous grid, if it is available
    If linGridx.UBound > 0 Then
        For i = 1 To linGridx.UBound
            If linGridx(i).Tag = mapnum Then
                Unload linGridx(i)
        
    'First unload previous grid
    If linGridx.UBound > 0 Then
        For i = 1 To linGridx.UBound
            Unload linGridx(i)
        Next i
    End If
    If linGridy.UBound > 0 Then
        For i = 1 To linGridy.UBound
            Unload linGridy(i)
        Next i
    End If

    'Set properties of the grid to reference lines
    With linGridx(0)
        .X2 = .X1 + imgMap.Width
        .BorderColor = gridcolor
        .BorderStyle = gridstyle
        .BorderWidth = gridwidth
    End With
    With linGridy(0)
        .Y2 = .Y1 + imgMap.Height
        .BorderColor = gridcolor
        .BorderStyle = gridstyle
        .BorderWidth = gridwidth
    End With

    'Create other lines from reference and create grid
    Load linGridx(1)
    With linGridx(1)
        .Y1 = 0 + offsetY
        .Y2 = 0 + offsetY
        .Visible = True
        .ZOrder 0
    End With

    Load linGridy(1)
    With linGridy(1)
        .X1 = 0 + offsetX
        .X2 = 0 + offsetX
        .Visible = True
        .ZOrder 0
    End With

    i = 1
    Do While (linGridx(i).Y1 + gridsize <= imgMap.Height)
        Load linGridx(i + 1)
        With linGridx(i + 1)
            .Y1 = linGridx(i).Y1 + gridsize
            .Y2 = linGridx(i).Y1 + gridsize
            .ZOrder 0
            .Visible = True
        End With
        i = i + 1
    Loop

    i = 1
    Do While (linGridy(1).X1 + i * gridsize <= imgMap.Width)
        Load linGridy(i + 1)
        With linGridy(i + 1)
            .X1 = linGridy(i).X1 + gridsize
            .X2 = linGridy(i).X1 + gridsize
            .ZOrder 0
            .Visible = True
        End With
        i = i + 1
    Loop
End Function



'Private Function pixel2mickey(pixel As Long, xy As String) As Long
'If xy = "X" Then
'    pixel2mickey = pixel / (Screen.Width / Screen.TwipsPerPixelX) * 65535
'ElseIf xy = "Y" Then
'    pixel2mickey = pixel / (Screen.Height / Screen.TwipsPerPixelY) * 65535
'End If
'End Function
'
'Private Function twip2mickey(twip As Single, xy As String) As Long
'If xy = "X" Then
'    twip2mickey = pixel2mickey((twip / Screen.TwipsPerPixelX), xy)
'ElseIf xy = "Y" Then
'    twip2mickey = pixel2mickey((twip / Screen.TwipsPerPixelY), xy)
'End If
'End Function
'
'Private Function mickey2pixel(mickey As Long, xy As String) As Long
'If xy = "X" Then
'    mickey2pixel = mickey / 65535 * (Screen.Width / Screen.TwipsPerPixelX)
'ElseIf xy = "Y" Then
'    mickey2pixel = mickey / 65535 * (Screen.Height / Screen.TwipsPerPixelY)
'End If
'End Function
'
'Private Function mickey2twip(mickey As Long, xy As String) As Single
'If xy = "X" Then
'    mickey2twip = mickey2pixel(mickey * Screen.TwipsPerPixelX, xy)
'ElseIf xy = "Y" Then
'    mickey2twip = mickey2pixel(mickey * Screen.TwipsPerPixelY, xy)
'End If
'End Function
'
'Private Function twip2pixel(twip As Single, xy As String) As Long
'If xy = "X" Then
'    twip2pixel = twip / Screen.TwipsPerPixelX
'ElseIf xy = "Y" Then
'    twip2pixel = twip / Screen.TwipsPerPixelY
'End If
'End Function
