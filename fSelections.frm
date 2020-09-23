VERSION 5.00
Begin VB.Form fSelections 
   BackColor       =   &H00C5A774&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selection Tools"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   446
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox canvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawStyle       =   4  'Dash-Dot-Dot
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00800000&
      Height          =   5565
      Left            =   45
      ScaleHeight     =   367
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   464
      TabIndex        =   0
      Top             =   645
      Width           =   7020
      Begin VB.Timer tmrSelection 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   5235
         Top             =   4785
      End
      Begin VB.Timer tmrAnts 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   4515
         Top             =   4890
      End
      Begin VB.Line lnPoly 
         BorderStyle     =   5  'Dash-Dot-Dot
         Index           =   0
         Visible         =   0   'False
         X1              =   49
         X2              =   171
         Y1              =   81
         Y2              =   141
      End
      Begin VB.Shape shpSelection 
         BorderStyle     =   5  'Dash-Dot-Dot
         Height          =   1110
         Left            =   555
         Top             =   2550
         Visible         =   0   'False
         Width           =   1755
      End
   End
   Begin VB.PictureBox picInvis 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5670
      Left            =   60
      ScaleHeight     =   378
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   2
      Top             =   555
      Visible         =   0   'False
      Width           =   7035
   End
   Begin VB.Image btnFLood 
      Height          =   480
      Left            =   4440
      Picture         =   "fSelections.frx":0000
      Top             =   60
      Width           =   540
   End
   Begin VB.Image btnMode 
      Height          =   480
      Index           =   1
      Left            =   3660
      Picture         =   "fSelections.frx":0585
      Top             =   60
      Width           =   540
   End
   Begin VB.Image btnMode 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   0
      Left            =   3075
      Picture         =   "fSelections.frx":0A9A
      Top             =   60
      Width           =   600
   End
   Begin VB.Image imgAnts 
      Height          =   120
      Index           =   1
      Left            =   4455
      Picture         =   "fSelections.frx":0FBC
      Top             =   6270
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image imgAnts 
      Height          =   120
      Index           =   0
      Left            =   4290
      Picture         =   "fSelections.frx":1026
      Top             =   6270
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   6315
      Width           =   4050
   End
   Begin VB.Image btnPicture 
      Height          =   465
      Left            =   5175
      Picture         =   "fSelections.frx":1090
      ToolTipText     =   "Click to load a picture"
      Top             =   60
      Width           =   555
   End
   Begin VB.Image btnSelectionTools 
      Height          =   480
      Index           =   6
      Left            =   2310
      Picture         =   "fSelections.frx":1687
      ToolTipText     =   "Magic Wand"
      Top             =   60
      Width           =   540
   End
   Begin VB.Image btnSelectionTools 
      Height          =   480
      Index           =   5
      Left            =   1740
      Picture         =   "fSelections.frx":1BEE
      ToolTipText     =   "Brush Mask"
      Top             =   60
      Width           =   540
   End
   Begin VB.Image btnSelectionTools 
      Height          =   480
      Index           =   4
      Left            =   1170
      Picture         =   "fSelections.frx":218B
      ToolTipText     =   "Polygon Mask"
      Top             =   60
      Width           =   540
   End
   Begin VB.Image btnSelectionTools 
      Height          =   480
      Index           =   3
      Left            =   1725
      Picture         =   "fSelections.frx":26C2
      ToolTipText     =   "Circle Mask"
      Top             =   60
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image btnSelectionTools 
      Height          =   480
      Index           =   2
      Left            =   585
      Picture         =   "fSelections.frx":2BF1
      ToolTipText     =   "Elliptical Mask"
      Top             =   60
      Width           =   540
   End
   Begin VB.Image btnSelectionTools 
      Height          =   480
      Index           =   1
      Left            =   585
      Picture         =   "fSelections.frx":311A
      ToolTipText     =   "Square Mask"
      Top             =   60
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image btnSelectionTools 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   0
      Left            =   15
      Picture         =   "fSelections.frx":363D
      ToolTipText     =   "Rectangle Mask"
      Top             =   60
      Width           =   600
   End
   Begin VB.Image btnRedo 
      Height          =   465
      Left            =   6510
      Picture         =   "fSelections.frx":3B69
      ToolTipText     =   "Redo"
      Top             =   60
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Image btnUndo 
      Height          =   480
      Left            =   5925
      Picture         =   "fSelections.frx":4156
      ToolTipText     =   "Undo"
      Top             =   60
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "fSelections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===================================================================================
' ===================================================================================
' Selections test code ver 1.0

' by M Ferris - Intact Interactive Software.
' url   : http://www.intactinteractive.com
' email : mferris@zfree.co.nz
'
' ===================================================================================
' ===================================================================================
'
' This sample code has been given to the development community by Intact Interactive
' software. The code is not finished commercial level stuff so you will find bugs and
' quirks in it I'm sure. Neither is the code a complete application, it is merely a
' testbed for the functionalily being tested!
' So you use this code on an as is basis with no warranty given or implied by Intact
' Interacive. If you are cool with this then go for it, if not, then delete this.
'
' Also remember that this is not a one way street, I have to survive as a developer
' as well, so obviously I am not giving away all my knowledge. And I would appreciate
' it if you gave me feedback, vote for me at PSC or wherever you found this posted
' and, of course, download the eval versions of my commercial stuff, if you like what
' you see you might even consider buying some of it ...
'
' A more complete example of this sample source code can be found at
' http://www.intactinteractive.com so I encourage you to visit our site and check it
' out as well as our other source code samples ...
'
' Some notes about this version.
'
' Features :
'
'           - multiple undo/redo levels.
'           - progressive mask building, i.e. you can add/subtract in multiple steps
'           - add/subtract mode of mask building
'           - rectangle,ellipse and polygon selection tools functional
'
' How to use :
'              - Click on a selection tool to use it.
'              - For rectangle and ellipse tool - click on canvas and drag to
'              define your selection.
'              - For polygon tool - click to start a line, move to endpoint and
'              click again to end the line and begin next line, to close the shape
'              click near to the first lines starting point.
'              - Click the plus button for additive mask, and the minus for
'              subtractive mask mode.
'              - When undo is possible, the undo button will appear, click it to
'              undo to the last level.
'              - When redo is possible, the redo button will appear, click it to
'              redo to the last level.
'              - To fill the selection with red paint - click the paint bucket button.
'
' To Do :
'           - implement the paintbrush selection tool
'           - implement the wand selection tool
'           - implement disabled button states for undo/redo
'           - implement button states for pressed and over on all buttons
'           - modify code to work as an activex class (dll) and activex control (ocx)
'           - combine code with gradients sample to show implementation of area
'           - filling ...
'           - add combine mode for xor - i.e. removes the unions
'           - add combine mode for and - i.e. keeps only the union
'           - add mask transparency functionality
'           - add mask feather functionality
'
' ===================================================================================
' this form illustrates the use of regions to create selection areas for a paint
' program, it shows how to use CombineRgn() to progressively add to a mask region,
' how to use GetRegionData() to implement an undo/redo feature, and generally shows
' how the use of regions can make the implementation of complex selections possible
' in VB !
'
' the selection tools included in this sample are :
'
'                                   rectangular selection
'                                   elliptical selection
'                                   magic wand selection (yet to be implemented)
'                                   ploygonal selection
'                                   paint brush selection (yet to be implemented)
'
' also an undo function is available to allow the mask to be undone and redone up
' to all levels of the mask creation ...
' ===================================================================================

Option Explicit

' this is a container for an array of region data
Private Type rgns
    data() As Byte
    length As Integer
End Type
' our undo and redo arrays and counters
Private undoRgn() As rgns
Private numUndos As Integer
Private RedoRGN() As rgns
Private NumRedos As Integer
' the master region which we use for our selection
Private MasterRgn As Long
' mouse tracking vars
Private oldx As Single
Private oldy As Single
Private xOrig As Single
Private yOrig As Single
Private xDiff As Integer
Private yDiff As Integer
' flag to indicate selection has changed
Private bSelectionChanged As Boolean
' index var for the marching ants brushes
Private outlineType As Integer
' the marching ants brushes
Private antsBrush(1) As Integer
' a temporary region - used with paintbrush selection
Private rgnTmp As Long
' the selection tool currently active
Private CurrentSelectionTool As Integer
' the current mode for mask combination - i.e additive or subtractive
Private CombineMode As Long

Private Sub AddToUndo(rgn As Long)
    numUndos = numUndos + 1
    ReDim Preserve undoRgn(1 To numUndos)
    undoRgn(numUndos).length = GetRegionData(rgn, undoRgn(numUndos).length, ByVal 0&)
    ReDim Preserve undoRgn(numUndos).data(undoRgn(numUndos).length)
    GetRegionData rgn, undoRgn(numUndos).length, undoRgn(numUndos).data(0)
    btnUndo.Visible = True
End Sub

Private Sub AddToRedo(rgn As Long)
    NumRedos = NumRedos + 1
    ReDim Preserve RedoRGN(1 To NumRedos)
    RedoRGN(NumRedos).length = GetRegionData(rgn, RedoRGN(NumRedos).length, ByVal 0&)
    ReDim Preserve RedoRGN(NumRedos).data(RedoRGN(NumRedos).length)
    GetRegionData rgn, RedoRGN(NumRedos).length, RedoRGN(NumRedos).data(0)
    btnRedo.Visible = True
End Sub

Private Sub btnFLood_Click()
    Dim hbr As Long
    Dim lbr As LOGBRUSH
    
    ' this only floods the current selection and doesn't work completely yet -
    ' i.e. when you add/subtract - undo/redo then mask the flood disappears at
    ' present ...
    lbr.lbColor = RGB(255, 0, 0)
    lbr.lbStyle = BS_SOLID
    hbr = CreateBrushIndirect(lbr)
    FillRgn canvas.hdc, MasterRgn, hbr
    DeleteObject hbr
End Sub

Private Sub btnMode_Click(Index As Integer)
    ' change the combine mode to either additive or subtractive
    If Index = 0 Then
        btnMode(Index).BorderStyle = 1
        btnMode(1).BorderStyle = 0
        CombineMode = RGN_OR
    Else
        btnMode(Index).BorderStyle = 1
        btnMode(0).BorderStyle = 0
        CombineMode = RGN_DIFF
    End If
End Sub

Private Sub btnPicture_Click()
    MsgBox "Visit http://www.intactinteractive.com for a more complete example of this code!", , "Not Implemented yet!"
End Sub

Private Sub btnSelectionTools_Click(Index As Integer)
    Dim i As Integer

    CurrentSelectionTool = Index
    For i = 0 To 6
        If Index = i Then
            btnSelectionTools(i).BorderStyle = 1
        Else
            btnSelectionTools(i).BorderStyle = 0
        End If
    Next i
'    If Index = 5 Then
'        ' set the brush width ...
'        fBrushSettings.Show 1
'    End If
    
    If Index > 4 Then
        MsgBox "Visit http://www.intactinteractive.com for a more complete example of this code!", , "Not Implemented yet!"
    End If
End Sub

Private Sub btnRedo_Click()
    Dim idx As Integer
    
    idx = UBound(RedoRGN)
    ' add the existing region to the undo array ...
    AddToUndo MasterRgn
    ' recreate the mask from the redo region data ...
    DeleteObject MasterRgn
    MasterRgn = ExtCreateRegion(ByVal 0&, RedoRGN(idx).length, RedoRGN(idx).data(0))
    bSelectionChanged = True
    ' now trim off the region from the redo array ...
    Erase RedoRGN(idx).data
    If idx - 1 <= 0 Then
        ReDim RedoRGN(1 To 1)
        Erase RedoRGN(1).data
        NumRedos = 0
        btnRedo.Visible = False
    Else
        idx = idx - 1
        ReDim Preserve RedoRGN(1 To idx)
        NumRedos = idx
    End If
End Sub

Private Sub btnUndo_Click()
    Dim idx As Integer
    
    idx = UBound(undoRgn)
    
    ' copy the last undo region into the redo array ...
    AddToRedo MasterRgn
    ' recreate the mask from the undo region data ...
    DeleteObject MasterRgn
    MasterRgn = ExtCreateRegion(ByVal 0&, undoRgn(idx).length, undoRgn(idx).data(0))
    bSelectionChanged = True
    ' now trim off the region from the undo array ...
    Erase undoRgn(idx).data
    If idx - 1 <= 0 Then
        ReDim undoRgn(1 To 1)
        Erase undoRgn(1).data
        numUndos = 0
        btnUndo.Visible = False
    Else
        idx = idx - 1
        ReDim Preserve undoRgn(1 To idx)
        numUndos = idx
    End If
End Sub

Private Sub canvas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' set the beginning of mask selection depending on the type of mask tool ...
    oldx = x
    oldy = y
    xOrig = x
    yOrig = y
    Select Case CurrentSelectionTool
        Case 0, 1, 2, 3 ' rect,square,circle,ellipse ...
            shpSelection.left = x
            shpSelection.top = y
            shpSelection.Width = 1
            shpSelection.Height = 1
            shpSelection.Visible = True
            tmrAnts.Enabled = True
            shpSelection.Shape = CurrentSelectionTool
            
        Case 4 ' polygon
                ' add another line to the region and continue building the polygon ...
            If lnPoly(0).Visible Then
                Load lnPoly(lnPoly.Count)
                lnPoly(lnPoly.UBound).x1 = lnPoly(lnPoly.UBound - 1).x2
                lnPoly(lnPoly.UBound).y1 = lnPoly(lnPoly.UBound - 1).y2
                lnPoly(lnPoly.UBound).x2 = x
                lnPoly(lnPoly.UBound).y2 = y
                lnPoly(lnPoly.UBound).Visible = True
            Else
                lnPoly(0).x1 = x
                lnPoly(0).y1 = y
                lnPoly(0).x2 = x
                lnPoly(0).y2 = y
                lnPoly(0).Visible = True
                tmrAnts.Enabled = True
                CurrentSelectionTool = 4
            End If
        Case 5 ' brush mask
        Case 6 ' magic wand
    End Select
End Sub

Private Sub canvas_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' handle the movement of the selection tool ...
    Dim pt As POINTAPI
    Dim rgn As Long
    Dim res As Long
    
    xDiff = x - oldx
    yDiff = y - oldy
    Select Case CurrentSelectionTool
        Case 0, 1, 2, 3 ' rect,square,circle,ellipse ...
            If Button = 0 Then Exit Sub
            If x < xOrig Then
                shpSelection.left = x
                shpSelection.Width = xOrig - x
            Else
                shpSelection.Width = shpSelection.Width + xDiff
            End If
            If y < yOrig Then
                shpSelection.top = y
                shpSelection.Height = yOrig - y
            Else
                shpSelection.Height = shpSelection.Height + yDiff
            End If
        Case 4 ' polygon
            lnPoly(lnPoly.UBound).x2 = x
            lnPoly(lnPoly.UBound).y2 = y
            canvas.Refresh
        Case 5 ' brush mask
        Case 6 ' magic wand
    End Select
    oldx = x
    oldy = y
End Sub

Private Sub canvas_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim rgnRects() As rect
    Dim rgnSize As Integer
    Dim rc As rect
    Dim sz As Integer
    Dim pts() As POINTAPI
    Dim numPoints As Integer
    Dim i As Integer
    Dim res As Long
    Dim rgn As Long
    
    shpSelection.Visible = False
    
    Select Case CurrentSelectionTool
        Case 0, 1 ' rectangle, square
            tmrAnts.Enabled = False
            rgn = CreateRectRgn(shpSelection.left, _
                                shpSelection.top, _
                                shpSelection.left + shpSelection.Width, _
                                shpSelection.top + shpSelection.Height)
            If MasterRgn = 0 Then
                MasterRgn = CreateRectRgn(shpSelection.left, _
                                            shpSelection.top, _
                                            shpSelection.left + shpSelection.Width, _
                                            shpSelection.top + shpSelection.Height)
                bSelectionChanged = True
            Else
                AddToUndo MasterRgn
                CombineRgn MasterRgn, MasterRgn, rgn, CombineMode
                bSelectionChanged = True
            End If
        Case 2, 3 ' circle, ellipse
            tmrAnts.Enabled = False
            rgn = CreateEllipticRgn(shpSelection.left, _
                                    shpSelection.top, _
                                    shpSelection.left + shpSelection.Width, _
                                    shpSelection.top + shpSelection.Height)
                 
           If MasterRgn = 0 Then
                MasterRgn = CreateEllipticRgn(shpSelection.left, _
                                            shpSelection.top, _
                                            shpSelection.left + shpSelection.Width, _
                                            shpSelection.top + shpSelection.Height)
                
                bSelectionChanged = True
            Else
                ' save the old master region into a data array
                AddToUndo MasterRgn
                CombineRgn MasterRgn, MasterRgn, rgn, CombineMode
                
                bSelectionChanged = True
            End If
                                    
       Case 4 ' polygon
            ' first see if the polygon is closing ...
            If lnPoly.UBound <> 0 And x > lnPoly(0).x1 - 5 And x < lnPoly(0).x1 + 5 And y > lnPoly(0).y1 - 5 And y < lnPoly(0).y1 + 5 Then
                ' and if it is do the following ...
                tmrAnts.Enabled = False
                ' build a polygon region based on the lines ...
                numPoints = lnPoly.Count + 1
                ReDim pts(numPoints)
                For i = 0 To lnPoly.UBound
                    pts(i).x = lnPoly(i).x1
                    pts(i).y = lnPoly(i).y1
                Next i
                pts(numPoints - 1).x = lnPoly(0).x1
                pts(numPoints - 1).y = lnPoly(0).y1
                rgn = CreatePolygonRgn(pts(0), numPoints, WINDING)
                ' clean up the existing lines ...
                For i = 1 To lnPoly.UBound
                    Unload lnPoly(i)
                Next i
                lnPoly(0).Visible = False
                ' combine it to the current master region ...
                If MasterRgn = 0 Then
                    MasterRgn = CreatePolygonRgn(pts(0), numPoints, WINDING)
                    bSelectionChanged = True
                Else
                    AddToUndo MasterRgn
                    CombineRgn MasterRgn, MasterRgn, rgn, CombineMode
                    bSelectionChanged = True
                End If
            End If
        Case 5 ' brush mask
        Case 6 ' magic wand
    End Select
    tmrSelection.Enabled = True
End Sub

Private Sub Form_Load()
    CurrentSelectionTool = 0
    btnSelectionTools(CurrentSelectionTool).BorderStyle = 1
    antsBrush(0) = CreatePatternBrush(imgAnts(0).Picture.Handle)
    antsBrush(1) = CreatePatternBrush(imgAnts(1).Picture.Handle)
    canvas.Picture = canvas.Image
    canvas.Refresh
    CombineMode = RGN_OR
    numUndos = 0
    NumRedos = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteObject antsBrush(0)
    DeleteObject antsBrush(1)
End Sub

Private Sub tmrAnts_Timer()
    
    Select Case CurrentSelectionTool
        Case 0, 1, 2, 3
            If shpSelection.BorderStyle = 4 Then
                shpSelection.BorderStyle = 3
            Else
                shpSelection.BorderStyle = 4
            End If
        Case 4
            Dim i As Integer
            
            For i = 0 To lnPoly.UBound
                lnPoly(i).BorderStyle = IIf(lnPoly(i).BorderStyle = 4, 5, 4)
            Next i
    End Select
End Sub

Private Sub tmrSelection_Timer()
    Dim res As Integer
    Static idx As Integer
    
    If bSelectionChanged Then
        bSelectionChanged = False
        canvas.Cls
        canvas.Picture = LoadPicture("")
        canvas.Refresh
        outlineType = 0
        idx = 0
        ' redraw the previous selection outline
        res = FrameRgn(canvas.hdc, MasterRgn, antsBrush(outlineType), 1, 1)
        canvas.Refresh
    End If
    res = FrameRgn(canvas.hdc, MasterRgn, antsBrush(outlineType), 1, 1)
    outlineType = IIf(outlineType = 0, 1, 0)
    res = FrameRgn(canvas.hdc, MasterRgn, antsBrush(outlineType), 1, 1)
    canvas.Refresh
End Sub
