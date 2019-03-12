VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20160
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   12240
      Top             =   1680
   End
   Begin VB.Shape ShapeA 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   19320
      Shape           =   2  'Oval
      Top             =   3480
      Width           =   375
   End
   Begin VB.Shape ShapeB 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   19200
      Shape           =   2  'Oval
      Top             =   4440
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const G As Single = 6.75 * 0.0000000001
Dim PlanetA As Planet, PlanetB As Planet

Private Type Vector
    x As Double
    y As Double
End Type

Private Type Planet
   Mass As Double
   Radius As Double
   Velocity As Vector
   Acceleration As Vector
End Type



Private Sub Form_Load()
 With PlanetA
  .Mass = 200000#
  .Velocity.x = -30
  ShapeA.Width = .Mass / 1000
  ShapeA.Height = .Mass / 1000
 End With
 
 With PlanetB
  .Mass = 100000#
  .Velocity.x = 30
  ShapeB.Width = .Mass / 1000
  ShapeB.Height = .Mass / 1000
 End With
End Sub

Private Sub Timer1_Timer()
    On Error GoTo 1
    Dim Distance As Double, a As Double, PlaACenter As Vector, PlaBCenter As Vector
    
    PlaACenter.x = ShapeA.Left + ShapeA.Width / 2
    PlaACenter.y = ShapeA.Top + ShapeA.Height / 2
    
    PlaBCenter.x = ShapeB.Left + ShapeB.Width / 2
    PlaBCenter.y = ShapeB.Top + ShapeB.Height / 2
    
    Distance = (Sqr((PlaACenter.x - PlaBCenter.x) ^ 2 + (PlaACenter.y - PlaBCenter.y) ^ 2)) / 100000
    
    With PlanetA
     a = (G * PlanetB.Mass) / Distance ^ 2
     
     .Acceleration.x = (PlaBCenter.x - PlaACenter.x) / Abs(PlaBCenter.x - PlaACenter.x) * a
     .Acceleration.y = (PlaBCenter.y - PlaACenter.y) / Abs(PlaBCenter.y - PlaACenter.y) * a
     
     .Velocity.x = PlanetA.Velocity.x + PlanetA.Acceleration.x
     .Velocity.y = PlanetA.Velocity.y + PlanetA.Acceleration.y
     
    End With
    
    With PlanetB

     a = (G * PlanetA.Mass) / Distance ^ 2
     
     .Acceleration.x = (PlaACenter.x - PlaBCenter.x) / Abs(PlaACenter.x - PlaBCenter.x) * a
     .Acceleration.y = (PlaACenter.y - PlaBCenter.y) / Abs(PlaACenter.y - PlaBCenter.y) * a
     
     .Velocity.x = PlanetB.Velocity.x + PlanetB.Acceleration.x
     .Velocity.y = PlanetB.Velocity.y + PlanetB.Acceleration.y
     
    End With
1:
    With ShapeA: ShapeA.Move .Left + PlanetA.Velocity.x, .Top + PlanetA.Velocity.y: End With
    With ShapeB: ShapeB.Move .Left + PlanetB.Velocity.x, .Top + PlanetB.Velocity.y: End With
    
    Me.PSet (PlaACenter.x, PlaACenter.y), RGB(250, 50, 50)
    Me.PSet (PlaBCenter.x, PlaBCenter.y), RGB(50, 50, 250)
    Me.Refresh
    
End Sub
