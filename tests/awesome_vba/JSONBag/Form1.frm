VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "JsonBag to TreeView"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5910
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4815
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.TreeView TreeView1 
      Height          =   2535
      Left            =   300
      TabIndex        =   0
      Top             =   180
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   4471
      _Version        =   327682
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3660
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0000
            Key             =   "Array"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":02BA
            Key             =   "Object"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":03BC
            Key             =   "Element"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":04BE
            Key             =   "Property"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub TreeTheNode( _
    ByVal Name As Variant, _
    ByVal Item As Variant, _
    Optional ByVal Parent As ComctlLib.Node)

    Dim ImageKey As String
    Dim NewNode As ComctlLib.Node
    Dim I As Long
    Dim ItemAsText As String
    Dim Text As String

    With TreeView1.Nodes
        If VarType(Item) = vbObject Then
            If Item.IsArray Then
                ImageKey = "Array"
            Else
                ImageKey = "Object"
            End If
            If VarType(Name) <> vbString Then
                Name = "(" & CStr(Name) & ")"
            End If
            If Parent Is Nothing Then
                Set NewNode = .Add(, , Name, Name, ImageKey)
            Else
                Set NewNode = .Add(Parent.Key, tvwChild, Parent.Key & "\" & Name, Name, ImageKey)
            End If
            For I = 1 To Item.Count
                TreeTheNode Item.Name(I), Item.Item(I), NewNode
            Next
            NewNode.Expanded = True
        Else 'Value.
            Select Case VarType(Item)
                Case vbNull
                    ItemAsText = "Null"
                Case vbString
                    ItemAsText = """" & Item & """"
                Case Else
                    ItemAsText = CStr(Item)
            End Select
            If VarType(Name) = vbString Then
                ImageKey = "Property"
                Text = Name & ": " & ItemAsText
            Else
                ImageKey = "Element"
                Name = "(" & CStr(Name) & ")"
                Text = Name & " " & ItemAsText
            End If
            If Parent Is Nothing Then
                .Add , , Name, Text, ImageKey
            Else
                .Add Parent.Key, tvwChild, Parent.Key & "\" & Name, Text, ImageKey
            End If
        End If
    End With
End Sub

Private Sub Form_Load()
    Dim JsonBag As JsonBag
    Dim F As Integer

    F = FreeFile(0)
    Open "sample.json" For Input As #F
    Set JsonBag = New JsonBag
    JsonBag.JSON = Input$(LOF(F), #F)
    Close #F
    TreeTheNode "[Root]", JsonBag
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        TreeView1.Move 0, 0, ScaleWidth, ScaleHeight
    End If
End Sub
