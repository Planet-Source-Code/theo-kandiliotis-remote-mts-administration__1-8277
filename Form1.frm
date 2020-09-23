VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote MTS Administration"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbComputer 
      Height          =   315
      Left            =   3210
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   3375
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   5
      Left            =   210
      ScaleHeight     =   2895
      ScaleWidth      =   6315
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6150
      Width           =   6315
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refres&h"
         Height          =   345
         Left            =   4650
         TabIndex        =   13
         Top             =   2460
         Width           =   1605
      End
      Begin MSComctlLib.TreeView trvPackages 
         Height          =   2325
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   4101
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   6
         HotTracking     =   -1  'True
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Component"
         Height          =   195
         Index           =   2
         Left            =   1950
         TabIndex        =   28
         Top             =   2520
         Width           =   810
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   1
         Left            =   1500
         Top             =   2490
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Package"
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   27
         Top             =   2520
         Width           =   645
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Index           =   0
         Left            =   90
         Top             =   2490
         Width           =   405
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   2775
      Index           =   2
      Left            =   6780
      ScaleHeight     =   2775
      ScaleWidth      =   6315
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8460
      Width           =   6315
      Begin VB.TextBox txtPackageName 
         Height          =   315
         Index           =   2
         Left            =   2220
         TabIndex        =   5
         Top             =   1500
         Width           =   3165
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   600
         Picture         =   "Form1.frx":0442
         Top             =   510
         Width           =   480
      End
      Begin VB.Label Label7 
         Caption         =   "This option will delete the package with the specified name and create a new one with the same name and the same components."
         Height          =   675
         Left            =   1200
         TabIndex        =   26
         Top             =   510
         Width           =   4905
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Package name:"
         Height          =   195
         Left            =   810
         TabIndex        =   20
         Top             =   1560
         Width           =   1125
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   2805
      Index           =   3
      Left            =   6720
      ScaleHeight     =   2805
      ScaleWidth      =   6315
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5610
      Width           =   6315
      Begin VB.TextBox txtPackageName 
         Height          =   315
         Index           =   3
         Left            =   2220
         TabIndex        =   6
         Top             =   930
         Width           =   3165
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Package name:"
         Height          =   195
         Index           =   0
         Left            =   810
         TabIndex        =   22
         Top             =   990
         Width           =   1125
      End
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   2805
      Index           =   4
      Left            =   6840
      ScaleHeight     =   2805
      ScaleWidth      =   6315
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2730
      Width           =   6315
      Begin VB.TextBox txtShutdown 
         Height          =   315
         Left            =   4770
         TabIndex        =   8
         Top             =   180
         Width           =   765
      End
      Begin VB.CommandButton cmdRemoveComponent 
         Height          =   615
         Left            =   5160
         Picture         =   "Form1.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdAddComponent 
         Height          =   615
         Left            =   5160
         Picture         =   "Form1.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1140
         Width           =   975
      End
      Begin VB.ListBox lstComponents 
         Height          =   1815
         ItemData        =   "Form1.frx":1108
         Left            =   150
         List            =   "Form1.frx":110A
         TabIndex        =   9
         Top             =   870
         Width           =   4905
      End
      Begin VB.TextBox txtPackageName 
         Height          =   315
         Index           =   4
         Left            =   1320
         TabIndex        =   7
         Top             =   180
         Width           =   2085
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Shutdown after                    mins"
         Height          =   195
         Index           =   2
         Left            =   3600
         TabIndex        =   29
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Package components:"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   25
         Top             =   630
         Width           =   1605
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Package name:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   24
         Top             =   240
         Width           =   1125
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3900
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3300
      Picture         =   "Form1.frx":110C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1785
   End
   Begin VB.CommandButton cmdDoIt 
      Caption         =   "&Do it"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1470
      Picture         =   "Form1.frx":154E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1785
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      Height          =   2805
      Index           =   1
      Left            =   180
      ScaleHeight     =   2805
      ScaleWidth      =   6315
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1890
      Width           =   6315
      Begin VB.TextBox txtPackageName 
         Height          =   315
         Index           =   1
         Left            =   2220
         TabIndex        =   4
         Top             =   930
         Width           =   3165
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Package name:"
         Height          =   195
         Left            =   810
         TabIndex        =   18
         Top             =   990
         Width           =   1125
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3525
      Left            =   120
      TabIndex        =   3
      Top             =   1260
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   6218
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Shut down package"
            Key             =   "ShutDown"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Recreate"
            Key             =   "Recreate"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "De&lete"
            Key             =   "Delete"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Create new"
            Key             =   "Create"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Enumerate packages"
            Key             =   "Enum"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Action to perform:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   780
      TabIndex        =   16
      Top             =   840
      Width           =   1785
   End
   Begin VB.Image Image1 
      Height          =   510
      Index           =   1
      Left            =   120
      Picture         =   "Form1.frx":1990
      Top             =   645
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   405
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":1E87
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Computer to connect to:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   780
      TabIndex        =   15
      Top             =   180
      Width           =   2715
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Catalog As MTSAdmin.Catalog

Private ComputerNames As New Collection

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Function SystemPath() As String
   Dim SystemDir As String
   SystemDir = String(500, " ")
   GetSystemDirectory SystemDir, 499
   SystemPath = RTrim(SystemDir)
   SystemPath = Left(SystemPath, Len(SystemPath) - 1)
   If Right(SystemPath, 1) <> "\" Then SystemPath = SystemPath & "\"
End Function

Private Function TempPath() As String
   Dim TempDir As String
   TempDir = String(500, " ")
   GetTempPath 499, TempDir
   TempPath = RTrim(TempDir)
   TempPath = Left(TempPath, Len(TempPath) - 1)
   If Right(TempPath, 1) <> "\" Then TempPath = TempPath & "\"
End Function



Private Sub cmdAbout_Click()
   Dim msg As String
End Sub

Private Sub cmdAddComponent_Click()
   On Error GoTo ErrHandler
   
   With CommonDialog1
      .DefaultExt = 0
      .DialogTitle = "Select component to add..."
      .Filter = "ActiveX DLL (*.DLL)|*.DLL|ActiveX EXE (*.EXE)|*.EXE|All files|*.*"
      .FilterIndex = 0
      .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNPathMustExist
      .ShowOpen
   
      If .FileName <> "" Then lstComponents.AddItem .FileName
         
   End With
   
   Exit Sub
   
ErrHandler:
      
End Sub

Private Sub ClearArray(MyArray() As String)
   Dim x, y, z As Long
   Dim OK As Boolean
   Dim NewArray() As String
   ReDim NewArray(1 To 1)
   
   z = 0
   For x = 1 To UBound(MyArray)
      OK = True
      For y = x + 1 To UBound(MyArray)
         If MyArray(y) = MyArray(x) Then OK = False
      Next
      If OK Then
         z = z + 1
         ReDim NewArray(1 To z)
         NewArray(z) = MyArray(x)
         
      End If
   Next
   
   MyArray = NewArray
   
End Sub

Private Sub cmdDoIt_Click()
   
'   On Error GoTo ErrHandler
   
   Dim Found As Boolean, i As Long, j As Long
   Dim PackageName As String
   Dim Packages As MTSAdmin.CatalogCollection
   Dim PackUtil As New MTSAdmin.PackageUtil
   Dim Pack As MTSAdmin.CatalogObject
   Dim Comp As MTSAdmin.CatalogObject
   Dim Components As MTSAdmin.CatalogCollection
   Dim CompUtil As MTSAdmin.ComponentUtil
   Dim LocalComputer As MTSAdmin.CatalogCollection

   
   If Trim(cmbComputer.Text) = "" Then
      MsgBox "You must provide a computer name to connect to.", vbInformation
      GoTo CleanExit
   End If
   
   Dim ConnectResult As String
   ConnectResult = Connect(cmbComputer.Text)
   If ConnectResult <> "" Then
      MsgBox ConnectResult, vbExclamation + vbOKOnly, "Error"
      GoTo CleanExit
   End If
   
   Found = False
   Screen.MousePointer = vbHourglass
   
   Select Case TabStrip1.SelectedItem.Index
   
   Case 1 'Shutdown package
      PackageName = txtPackageName(1)
      If Trim(PackageName) = "" Then
         MsgBox "You must provide the name of the package you want to shut down.", vbExclamation + vbOKOnly
         GoTo CleanExit
      End If
      Set Packages = Catalog.GetCollection("Packages")
      Packages.Populate
      For Each Pack In Packages
         If Pack.Name = PackageName Then
            Found = True
            Set PackUtil = Packages.GetUtilInterface
            PackUtil.ShutdownPackage Pack.Value("ID")
            Exit For
         End If
      Next
      If Found = False Then MsgBox "Package '" & PackageName & "' not found in computer '" & cmbComputer & "'", vbExclamation + vbOKOnly
   
   
   Case 2 'Recreate
      PackageName = txtPackageName(2)
      If Trim(PackageName) = "" Then
         MsgBox "You must provide the name of the package you want to recreate.", vbExclamation + vbOKOnly
         GoTo CleanExit
      End If
      
      Dim PackageID As String, PackageIndex As Long, PackageShutDownTime As Integer
      Dim DLLs() As String: ReDim DLLs(1 To 1)
   
      Set Packages = Catalog.GetCollection("Packages"): Packages.Populate
      For i = 0 To Packages.Count - 1
         Set Pack = Packages.Item(i)
         If Pack.Name = PackageName Then
            Found = True
            PackageIndex = i
            PackageID = Pack.Value("ID")
            PackageShutDownTime = Pack.Value("ShutdownAfter")
            Set Components = Packages.GetCollection("ComponentsInPackage", PackageID): Components.Populate
            For j = 0 To Components.Count - 1
               ReDim Preserve DLLs(1 To j + 1)
               DLLs(j + 1) = Components.Item(j).Value("DLL")
            Next
            Exit For
         End If
      Next
      
      If Found = False Then
         MsgBox "Package '" & PackageName & "' not found in computer '" & cmbComputer & "'", vbExclamation + vbOKOnly
      Else
         ClearArray DLLs
         Packages.Remove PackageIndex
         Packages.SaveChanges
         
         Set Packages = Catalog.GetCollection("Packages"): Packages.Populate
         Set Pack = Packages.Add
         Pack.Value("Name") = PackageName
         Pack.Value("ShutdownAfter") = PackageShutDownTime
         Packages.SaveChanges
         
         Set Packages = Catalog.GetCollection("Packages"): Packages.Populate
         Set Components = Packages.GetCollection("ComponentsInPackage", Pack.Value("ID")): Components.Populate
         Set CompUtil = Components.GetUtilInterface
         For i = 1 To UBound(DLLs)
            CompUtil.InstallComponent DLLs(i), "", ""
         Next
         Packages.SaveChanges
         
'         Set PackUtil = Packages.GetUtilInterface
'         Set LocalComputer = Catalog.GetCollection("LocalComputer"): LocalComputer.Populate
'         PackUtil.ExportPackage PackageID, TempPath & "TempPackage.PAK", 0
'         Packages.Remove i
'         Packages.SaveChanges
'         Set Packages = Catalog.GetCollection("Packages"): Packages.Populate
'         Set PackUtil = Packages.GetUtilInterface
'         PackUtil.InstallPackage TempPath & "TempPackage.PAK", LocalComputer.Item(0).Value("PackageInstallPath"), 0
'         Packages.SaveChanges
      End If
 
   
   
   Case 3 'Delete
      PackageName = txtPackageName(3)
      If Trim(PackageName) = "" Then
         MsgBox "You must provide the name of the package you want to delete.", vbExclamation + vbOKOnly
         GoTo CleanExit
      End If
      Set Packages = Catalog.GetCollection("Packages")
      Packages.Populate
      For i = 0 To Packages.Count - 1
         Set Pack = Packages.Item(i)
         If Pack.Name = PackageName Then
            Found = True
            Set Components = Packages.GetCollection("ComponentsInPackage", Pack.Key)
            Components.Populate
            Dim msg As String
            For Each Comp In Components
               msg = msg & vbTab & Comp.Name & vbCrLf
            Next
            If msg <> "" Then
               If MsgBox("The package '" & PackageName & "' contains the following components:" & vbCrLf & vbCrLf & msg & vbCrLf & vbCrLf & "Are you sure you want to delete this package?", vbQuestion + vbYesNo) = vbNo Then GoTo CleanExit
            End If
            Packages.Remove i
            Packages.SaveChanges
            Exit For
         End If
      Next
      If Found = False Then MsgBox "Package '" & PackageName & "' not found in computer '" & cmbComputer & "'", vbExclamation + vbOKOnly
         
   
   Case 4 'Create new
      PackageName = txtPackageName(4)
      If Trim(PackageName) = "" Then
         MsgBox "You must provide the name of the package you want to create.", vbExclamation + vbOKOnly
         GoTo CleanExit
      End If
      
      Set Packages = Catalog.GetCollection("Packages"): Packages.Populate
      Set Pack = Packages.Add
      Pack.Value("Name") = PackageName
      Pack.Value("ShutdownAfter") = CLng(txtShutdown)
      Packages.SaveChanges
      
      Set Packages = Catalog.GetCollection("Packages"): Packages.Populate
      Set Components = Packages.GetCollection("ComponentsInPackage", Pack.Value("ID")): Components.Populate
      Set CompUtil = Components.GetUtilInterface
      
      For i = 0 To lstComponents.ListCount - 1
         CompUtil.InstallComponent lstComponents.List(i), "", ""
      Next
      
      Packages.SaveChanges
      
   
   End Select
   
   GoTo CleanExit

ErrHandler:

   MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "ERROR in " & Err.Source


CleanExit:
   Screen.MousePointer = vbDefault
End Sub



Private Sub cmdExit_Click()
   Unload Me
End Sub


Private Function Connect(ByVal Computer As String)
   
   Dim s, Exists As Boolean
   
   On Error Resume Next
   Screen.MousePointer = vbHourglass
   Set Catalog = New MTSAdmin.Catalog
   Err.Clear
   Catalog.Connect Computer
   Connect = Err.Description
   If Connect = "" Then
      For Each s In ComputerNames
         If s = Computer Then Exists = True
      Next
      If Not Exists Then
         ComputerNames.Add Computer
         cmbComputer.AddItem Computer
      End If
   End If
   
   Screen.MousePointer = vbDefault
End Function

Private Sub cmdRefresh_Click()
   EnumPackages
End Sub

Private Sub cmdRemoveComponent_Click()
   On Error Resume Next
   lstComponents.RemoveItem lstComponents.ListIndex
End Sub

Private Sub Form_Load()
   
   Screen.MousePointer = vbHourglass
   
   PopulateComputersList
   
   Dim i As Integer
   For i = 2 To 5
      picTab(i).Move picTab(1).Left, picTab(1).Top, picTab(1).Width, picTab(1).Height
   Next
   
   TabStrip1.Tabs(5).Selected = True
   
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub PopulateComputersList()

   Dim ComputerName As String, Names, s
   ComputerName = String(100, " ")
   GetComputerName ComputerName, 99
      
   Set ComputerNames = Nothing
   ComputerName = RTrim(ComputerName)
   ComputerName = Left(ComputerName, Len(ComputerName) - 1)
   
   
   ComputerNames.Add ComputerName, ComputerName
   
   Names = GetAllSettings(App.Title, "Computers")

   On Error Resume Next

   Dim i As Integer
   For i = LBound(Names, 1) To UBound(Names, 1)
      If Trim(Names(i, 1)) <> "" Then ComputerNames.Add Names(i, 1), Names(i, 1)
   Next
   
   cmbComputer.Clear
   For Each s In ComputerNames
      cmbComputer.AddItem s
   Next
   
   
   cmbComputer.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim s, i As Integer
   On Error Resume Next
   DeleteSetting App.Title, "Computers"
   For Each s In ComputerNames
      i = i + 1
      SaveSetting App.Title, "Computers", CStr(i), s
   Next
      
End Sub

Private Sub TabStrip1_Click()
   Dim i As Integer
   With TabStrip1
      
   If .SelectedItem.Index = 5 Then EnumPackages
   If .SelectedItem.Index = 4 Then txtShutdown = 3
   
   For i = 1 To .Tabs.Count
      If i = .SelectedItem.Index Then
         picTab(i).Enabled = True
         picTab(i).ZOrder 0
      Else
         picTab(i).Enabled = False
      End If
   Next
   cmdDoIt.Enabled = Not (.SelectedItem.Index = 5)
   End With
End Sub

Private Sub EnumPackages()
   On Error GoTo ErrHandler
   
   picTab(5).Enabled = False
   
   trvPackages.Nodes.Clear
   
   Dim i As Long
   Dim PackageName As String
   Dim Packages As MTSAdmin.CatalogCollection
   Dim Pack As MTSAdmin.CatalogObject
   Dim Comp As MTSAdmin.CatalogObject
   Dim Components As MTSAdmin.CatalogCollection
   Dim ConnectResult As String
   
   Dim MyNode As MSComctlLib.Node, OnotherNode As MSComctlLib.Node, RootNode As MSComctlLib.Node
   
   ConnectResult = Connect(cmbComputer.Text)
   If ConnectResult <> "" Then
      MsgBox ConnectResult, vbExclamation + vbOKOnly, "Error"
      GoTo CleanExit
   End If
   Screen.MousePointer = vbHourglass

   Set RootNode = trvPackages.Nodes.Add(, , , cmbComputer): RootNode.Bold = True
   
   Set Packages = Catalog.GetCollection("Packages")
   Packages.Populate
   For i = 0 To Packages.Count - 1
      Set Pack = Packages.Item(i)
      Set MyNode = trvPackages.Nodes.Add(RootNode.Index, tvwChild, , Pack.Name): MyNode.ForeColor = vbRed: MyNode.Bold = True: MyNode.EnsureVisible
      Set Components = Packages.GetCollection("ComponentsInPackage", Pack.Key)
      Components.Populate
      For Each Comp In Components
         Set OnotherNode = trvPackages.Nodes.Add(MyNode.Index, tvwChild, , Comp.Name): OnotherNode.ForeColor = vbBlue
         
      Next
   Next
 
   GoTo CleanExit

ErrHandler:
   MsgBox Err.Number & vbCrLf & Err.Description, vbExclamation + vbOKOnly, "ERROR in " & Err.Source
CleanExit:
   Screen.MousePointer = vbDefault
   picTab(5).Enabled = True
End Sub

Private Sub txtShutdown_Validate(Cancel As Boolean)
   Dim OK As Boolean
   
   OK = True
   
      
   If Trim(txtShutdown) = "" Then
      OK = False
   ElseIf Not IsNumeric(txtShutdown) Then
      OK = False
   ElseIf Int(txtShutdown) < 0 Or Int(txtShutdown) > 1440 Then
      OK = False
   End If

   If Not OK Then
      MsgBox "You must provide a shutdown delay for the package.", vbExclamation + vbOKOnly
      Cancel = True
      txtShutdown = 3
   End If
End Sub
