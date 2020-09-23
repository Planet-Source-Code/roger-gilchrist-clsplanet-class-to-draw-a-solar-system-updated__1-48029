VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlanet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ======================================================
'Based on
' Simple Solar System Simulator
' Version: 1.0
' by Peter Wilson
' Copyright © 2003 - Peter Wilson - All rights reserved.
' http://dev.midar.com/
'Class Conversion Copyright © 2003 Roger Gilchrist
'<rojagilkrist@hotmail.com>
' ======================================================
Private Enum ObType
 OStar
 OPlanet
 OMoon
 OOther
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private OStar, OPlanet, OMoon, OOther
#End If
Private O_Type                      As Integer
Private PiDiv180                    As Double
Private m_Form                      As Form
Private m_Label                     As Label
Private m_MoonLabel                 As Label
Private MoonLabels                  As New Collection
Private Type SatDesc
 SatName                           As String
 SatOrbit                          As Double
 SatDiameter                       As Double
 SatSpeed                          As Double
 SatRot                            As Currency
 SatColor                          As Long
 SatCX                             As Long
 SatCY                             As Long
 SatElip                           As Double
 SatRetro                          As Boolean
 SatIdNo                           As Long
End Type
Private CFm_MoonRotationAngle       As Currency
Private m_PlanetX                   As Currency
Private m_PlanetY                   As Currency
Private Moons()                     As SatDesc
Private MoonCount                   As Long
Private Zoom                        As Long
'these (CFm_)variables and the Properties they support were created by my Code Fixer program
Private CFm_TrueEliptic             As Double
Private CFm_Rings                   As Integer
Private CFm_DrawStyle               As FillStyleConstants
Private CFm_ShowGravArm             As Boolean
Private CFm_ShowOrbitPath           As Boolean
Private CFm_PlanetColor             As Long
Private CFm_Name                    As String
Private CFm_DispalyOrbitRadius      As Double
Private CFm_Diameter                As Double
Private CFm_OrbitSpeed              As Double
Private CFm_RotationAngle           As Currency
Private CFm_CentreX                 As Long
Private CFm_Eliptical               As Double
Private CFm_CentreY                 As Long
Private CFm_RetroGrade              As Boolean
Private CFm_Oblate                  As Double
Private CFm_ShowPlanetLabel         As Boolean
Private CFm_ShowMoonLabel           As Boolean
Private CFm_SunRadius               As Double
Private CFm_OrbitRadius             As Double

Public Property Get CentreX() As Long

 CentreX = CFm_CentreX

End Property

Public Property Let CentreX(ByVal PropVal As Long)

 CFm_CentreX = PropVal

End Property

Public Property Get CentreY() As Long

 CentreY = CFm_CentreY

End Property

Public Property Let CentreY(ByVal PropVal As Long)

 CFm_CentreY = PropVal

End Property

Private Sub Class_Initialize()

 PiDiv180 = 4 * Atn(1) / 180

End Sub

Private Sub Class_Terminate()

  'CLean up after myself

 Erase Moons
 Set MoonLabels = Nothing
 Set m_Form = Nothing
 Set m_Label = Nothing

End Sub

Private Function ConvertDeg2Rad(Degress As Double) As Double

  ' Converts Degrees to Radians

 ConvertDeg2Rad = Degress * PiDiv180

End Function

Public Property Get Diameter() As Double

 Diameter = CFm_Diameter

End Property

Public Property Let Diameter(ByVal PropVal As Double)

 CFm_Diameter = PropVal

End Property

Public Property Get DrawStyle() As FillStyleConstants

 DrawStyle = CFm_DrawStyle

End Property

Public Property Let DrawStyle(PropVal As FillStyleConstants)

 CFm_DrawStyle = PropVal

End Property

Public Property Get Eliptical() As Double

 Eliptical = CFm_Eliptical

End Property

Public Property Let Eliptical(ByVal PropVal As Double)

 CFm_Eliptical = PropVal

End Property

Private Sub GenerateNewLabel(ByVal CName As String, _
                             M_Lbl As Label)

  'if the label exists then don't make a new one
  'This idea comes form expansion of Mike James article "A Different Class' in Computer Shopper Jan 2003
  
  Dim Ctrl       As Object
  Dim CtrlExists As Boolean

 For Each Ctrl In m_Form.Controls
  If Ctrl.Name = CName Then
   CtrlExists = True
   Exit For
  End If
 Next Ctrl
 If Not CtrlExists Then
  Set M_Lbl = m_Form.Controls.Add("vb.Label", CName)
 End If

End Sub

Public Sub Init(F As Form, _
                ByVal Pname As String, _
                ByVal PSunRadius As Double, _
                ByVal POrbitRadius As Double, _
                ByVal PDiameter As Double, _
                ByVal POrbitSpeed As Double, _
                ByVal PColor As Long, _
                Optional ByVal initRot As Currency = 0, _
                Optional PRings As Integer = 0, _
                Optional CentrX As Long = 0, _
                Optional centry As Long = 0, _
                Optional eliptic As Double = 1, _
                Optional Retro As Boolean = False, _
                Optional Oblat As Double = 1)

 Set m_Form = F
 CFm_Name = Pname
 O_Type = OPlanet
 CFm_SunRadius = PSunRadius
 CFm_OrbitRadius = POrbitRadius
 CFm_DispalyOrbitRadius = CFm_SunRadius + CFm_OrbitRadius
 CFm_Diameter = PDiameter
 CFm_OrbitSpeed = POrbitSpeed
 CFm_PlanetColor = PColor
 RotationAngle = initRot
 CFm_Rings = PRings
 CFm_CentreX = CentrX
 CFm_CentreY = centry
 CFm_Eliptical = eliptic
 CFm_TrueEliptic = eliptic
 CFm_RetroGrade = Retro
 CFm_Oblate = Oblat
 If InStr(CFm_Name, "astreoid") = 0 Then
  GenerateNewLabel "Lbl" & CFm_Name, m_Label
  With m_Label
   .BackColor = vbWhite
   .ForeColor = vbBlack
   .AutoSize = True ' turning on and then off allow label to size itself then not worry again
   .Caption = CFm_Name
   .AutoSize = False
  End With 'm_Label
 End If

End Sub

Public Sub Moon(ByVal MName As String, _
                ByVal Morbit As Double, _
                ByVal Mdiam As Double, _
                ByVal MorbitSpd As Double, _
                ByVal MColor As Long, _
                Optional ByVal initRot As Currency = 0, _
                Optional cntrX As Long = 0, _
                Optional cntrY As Long = 0, _
                Optional eliptic As Double = 1, _
                Optional Retro As Boolean = False)

 MoonCount = MoonCount + 1
 ReDim Preserve Moons(MoonCount) As SatDesc
 With Moons(UBound(Moons))
  .SatIdNo = UBound(Moons)
  .SatName = MName
  .SatOrbit = Morbit
  .SatDiameter = Mdiam
  .SatSpeed = MorbitSpd
  .SatRot = initRot
  .SatColor = MColor
  .SatCX = cntrX
  .SatCY = cntrY
  If CFm_Oblate Then
   .SatElip = CFm_Oblate
   Else
   .SatElip = eliptic
  End If
  .SatRetro = False
  If InStr(.SatName, "astreoid") = 0 Then
   GenerateNewLabel "Lbl" & .SatName, m_MoonLabel
   With m_MoonLabel
    .BackColor = vbWhite
    .ForeColor = vbBlack
    .AutoSize = True ' turning on and then off allow label to size itself then not worry again
    .Caption = Moons(UBound(Moons)).SatName
    .AutoSize = False
   End With
   MoonLabels.Add m_MoonLabel
  End If
 End With

End Sub

Private Sub MoonLabel(mon As SatDesc, _
                      ByVal sngMX As Single, _
                      ByVal sngMY As Single)

 If CFm_ShowMoonLabel Then
  If InStr(mon.SatName, "astreoid") = 0 Then
   With MoonLabels(mon.SatIdNo)
    .Visible = True
    .Left = sngMX
    .Top = sngMY - mon.SatDiameter - .Height * 2
    m_Form.Line (.Left, .Top)-(sngMX, sngMY - mon.SatDiameter)
   End With 'm_Label
  End If
  Else
  If InStr(mon.SatName, "astreoid") = 0 Then
   MoonLabels(mon.SatIdNo).Visible = False
  End If
 End If

End Sub

Private Sub MoonMove(mon As SatDesc)

  Dim sngRadians          As Single
  Dim sngMoonX            As Single
  Dim sngMoonY            As Single

 With mon
  On Error Resume Next
  MoonRotationAngle = .SatRot
  .SatRot = MoonRotationAngle
  If .SatRetro Then
   sngRadians = ConvertDeg2Rad((.SatRot * -.SatSpeed) Mod 360)
   Else
   sngRadians = ConvertDeg2Rad((.SatRot * .SatSpeed) Mod 360)
  End If
  sngMoonX = (Sin(sngRadians) * .SatOrbit) + m_PlanetX
  sngMoonY = (((Cos(sngRadians) * .SatOrbit)) * CFm_Eliptical * .SatElip) + m_PlanetY
  On Error GoTo 0
 End With 'mon
 With m_Form
  .ForeColor = mon.SatColor
  If CFm_ShowOrbitPath Then
   .FillStyle = vbFSTransparent
   m_Form.Circle (m_PlanetX, m_PlanetY), mon.SatOrbit, , , , CFm_Eliptical * mon.SatElip
  End If
  If CFm_ShowGravArm Then
   .FillStyle = vbFSTransparent
   m_Form.Line (m_PlanetX, m_PlanetY)-(sngMoonX, sngMoonY)
  End If
  .FillColor = mon.SatColor
  .FillStyle = CFm_DrawStyle
  m_Form.Circle (sngMoonX, sngMoonY), mon.SatDiameter
  MoonLabel mon, sngMoonX, sngMoonY
 End With

End Sub

Private Property Get MoonRotationAngle() As Currency

 MoonRotationAngle = CFm_MoonRotationAngle

End Property

Private Property Let MoonRotationAngle(ByVal PropVal As Currency)

 CFm_MoonRotationAngle = PropVal
 CFm_MoonRotationAngle = CFm_MoonRotationAngle + 0.1 + Zoom '/ 1.3
 If CFm_MoonRotationAngle = 2 ^ 31 Then
  CFm_MoonRotationAngle = 0
 End If

End Property

Public Property Get Name() As String

 Name = CFm_Name

End Property

Public Property Let Name(ByVal PropVal As String)

 CFm_Name = PropVal

End Property

Public Property Get Oblate() As Double

 Oblate = CFm_Oblate

End Property

Public Property Let Oblate(ByVal PropVal As Double)

 CFm_Oblate = PropVal

End Property

Public Property Get OrbitRadius() As Double

 OrbitRadius = CFm_OrbitRadius

End Property

Public Property Let OrbitRadius(ByVal PropVal As Double)

 CFm_OrbitRadius = PropVal
 CFm_DispalyOrbitRadius = CFm_SunRadius + CFm_OrbitRadius

End Property

Public Property Get OrbitSpeed() As Double

 OrbitSpeed = CFm_OrbitSpeed

End Property

Public Property Let OrbitSpeed(ByVal PropVal As Double)

 CFm_OrbitSpeed = PropVal

End Property

Public Property Get PlanetColor() As Long

 PlanetColor = CFm_PlanetColor

End Property

Public Property Let PlanetColor(ByVal PropVal As Long)

 CFm_PlanetColor = PropVal

End Property

Private Sub PlanetLabel()

 If CFm_ShowPlanetLabel Then
  If InStr(CFm_Name, "astreoid") = 0 Then
   With m_Label
    .Visible = True
    .Left = CFm_CentreX + m_PlanetX - CFm_CentreX
    .Top = CFm_CentreY + m_PlanetY - CFm_Diameter - .Height * 3
    m_Form.Line (m_Label.Left, m_Label.Top)-(m_PlanetX, m_PlanetY - CFm_Diameter * CFm_Oblate)
   End With 'm_Label
  End If
  Else
  If InStr(CFm_Name, "astreoid") = 0 Then
   m_Label.Visible = False
  End If
 End If

End Sub

Public Sub PlanetMove(Optional ByVal zoomFactor As Single)

  Dim I                   As Integer
  Dim sngRadians          As Double

 Zoom = zoomFactor / 100
 RotationAngle = CFm_RotationAngle
 On Error Resume Next
 If RetroGrade Then
  sngRadians = ConvertDeg2Rad((CFm_RotationAngle * -CFm_OrbitSpeed) Mod 360)
  Else
  sngRadians = ConvertDeg2Rad((CFm_RotationAngle * CFm_OrbitSpeed) Mod 360)
 End If
 If O_Type = OStar Then
  'kludge stop sun from jumping about but assumes your sun is not moving
  m_PlanetX = ((Sin(sngRadians) * (0)) + CFm_CentreX)
  m_PlanetY = ((Cos(sngRadians) * 0) + CFm_CentreY) * CFm_Eliptical
  Else
  m_PlanetX = ((Sin(sngRadians) * (CFm_DispalyOrbitRadius)) + CFm_CentreX)
  m_PlanetY = ((Cos(sngRadians) * CFm_DispalyOrbitRadius) + CFm_CentreY) * CFm_Eliptical
 End If
 On Error GoTo 0
 With m_Form
  If MoonCount Then
   For I = 1 To MoonCount
    MoonMove Moons(I)
   Next I
  End If
  .ForeColor = CFm_PlanetColor
  If CFm_ShowOrbitPath Then
   .FillStyle = vbFSTransparent
   m_Form.Circle (CFm_CentreX, CFm_CentreY), CFm_DispalyOrbitRadius, , , , CFm_Eliptical
  End If
  If CFm_ShowGravArm Then
   .FillStyle = vbFSTransparent
   m_Form.Line (0, 0)-(m_PlanetX, m_PlanetY)
  End If
  .FillColor = .ForeColor
  .FillStyle = CFm_DrawStyle
  m_Form.Circle (m_PlanetX, m_PlanetY), CFm_Diameter, , , , CFm_Oblate
  If CFm_Rings Then
   .FillStyle = vbFSTransparent
   .DrawWidth = 2
   For I = 1 To CFm_Rings
    'these number/colours are random not real
    m_Form.Circle (m_PlanetX, m_PlanetY), CFm_Diameter + 2 + I, RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256)), , , CFm_Eliptical '/ CFm_Oblate
    'm_Form.Circle (m_PlanetX, m_PlanetY), CFm_Diameter + 2 + I, RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256)), 2.2, 0.8 + (I / 2000), CFm_Eliptical '/ CFm_Oblate
   Next I
   .DrawWidth = 1
  End If
  PlanetLabel
 End With

End Sub

Public Property Get RetroGrade() As Boolean

 RetroGrade = CFm_RetroGrade

End Property

Public Property Let RetroGrade(ByVal PropVal As Boolean)

 CFm_RetroGrade = PropVal

End Property

Public Property Get Rings() As Integer

 Rings = CFm_Rings

End Property

Public Property Let Rings(ByVal PropVal As Integer)

 CFm_Rings = PropVal

End Property

Public Property Get RotationAngle() As Currency

 RotationAngle = CFm_RotationAngle

End Property

Public Property Let RotationAngle(ByVal PropVal As Currency)

 CFm_RotationAngle = PropVal
 CFm_RotationAngle = CFm_RotationAngle + 0.1 + Zoom  '/ 1.3
 If CFm_RotationAngle = 2 ^ 31 Then
  CFm_RotationAngle = 0
  ElseIf CFm_RotationAngle < 0 Then
  CFm_RotationAngle = 2 ^ 31
 End If

End Property

Public Property Get ShowGravArm() As Boolean

 ShowGravArm = CFm_ShowGravArm

End Property

Public Property Let ShowGravArm(ByVal PropVal As Boolean)

 CFm_ShowGravArm = PropVal

End Property

Public Property Get ShowMoonLabel() As Boolean

 ShowMoonLabel = CFm_ShowMoonLabel

End Property

Public Property Let ShowMoonLabel(ByVal PropVal As Boolean)

 CFm_ShowMoonLabel = PropVal

End Property

Public Property Get ShowOrbitPath() As Boolean

 ShowOrbitPath = CFm_ShowOrbitPath

End Property

Public Property Let ShowOrbitPath(ByVal PropVal As Boolean)

 CFm_ShowOrbitPath = PropVal

End Property

Public Property Get ShowPlanetLabel() As Boolean

 ShowPlanetLabel = CFm_ShowPlanetLabel

End Property

Public Property Let ShowPlanetLabel(ByVal PropVal As Boolean)

 CFm_ShowPlanetLabel = PropVal

End Property

Public Sub Star(F As Form, _
                ByVal Pname As String, _
                ByVal PSunRadius As Double, _
                ByVal POrbitRadius As Double, _
                ByVal PDiameter As Double, _
                ByVal POrbitSpeed As Double, _
                ByVal PColor As Long, _
                Optional ByVal initRot As Currency = 0, _
                Optional PRings As Integer = 0, _
                Optional CentrX As Long = 0, _
                Optional centry As Long = 0, _
                Optional eliptic As Double = 1, _
                Optional Retro As Boolean = False, _
                Optional Oblat As Double = 1)

 Set m_Form = F
 CFm_Name = Pname
 O_Type = OStar
 CFm_SunRadius = PSunRadius
 CFm_OrbitRadius = POrbitRadius
 CFm_DispalyOrbitRadius = 0 'CFm_SunRadius + CFm_OrbitRadius
 CFm_Diameter = PDiameter
 CFm_OrbitSpeed = POrbitSpeed
 CFm_PlanetColor = PColor
 RotationAngle = initRot
 CFm_Rings = PRings
 CFm_CentreX = CentrX
 CFm_CentreY = centry
 CFm_Eliptical = eliptic
 CFm_TrueEliptic = eliptic
 CFm_RetroGrade = Retro
 CFm_Oblate = Oblat
 If InStr(CFm_Name, "astreoid") = 0 Then
  GenerateNewLabel "Lbl" & CFm_Name, m_Label
  With m_Label
   .BackColor = vbWhite
   .ForeColor = vbBlack
   .AutoSize = True ' turning on and then off allow label to size itself then not worry again
   .Caption = CFm_Name
   .AutoSize = False
  End With 'm_Label
 End If

End Sub

Public Property Get SunRadius() As Double

 SunRadius = CFm_SunRadius

End Property

Public Property Let SunRadius(ByVal PropVal As Double)

 CFm_SunRadius = PropVal
 CFm_DispalyOrbitRadius = CFm_SunRadius + CFm_OrbitRadius

End Property

Public Property Get TrueEliptic() As Double

 TrueEliptic = CFm_TrueEliptic

End Property

Public Property Let TrueEliptic(ByVal PropVal As Double)

 CFm_TrueEliptic = PropVal

End Property

':) Roja's VB Code Fixer V1.1.18 (29/08/2003 2:51:28 PM) 66 + 529 = 595 Lines Thanks Ulli for inspiration and lots of code.

