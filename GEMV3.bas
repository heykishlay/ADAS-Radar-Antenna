' ============================================================================
' CST Studio Suite 2024 - VBA Macro (CORRECTED: STAGE 1 PORT & GEOMETRY)
' Compact 77-GHz SIW-Fed 4-Element Patch Array with Single-Stub Phase Network
' ============================================================================
Option Explicit

Dim c0 As Double
Dim f0 As Double, fmin As Double, fmax As Double
Dim er As Double, hSub As Double, tCu As Double, tanD As Double
Dim fc As Double, viaD As Double, viaP As Double
Dim aSIW As Double
Dim lambda0 As Double, lambdaG As Double, elemD As Double
Dim patchW As Double, patchL As Double
Dim msW As Double, msEeff As Double, lambdaMS As Double, phaseDelta As Double
Dim outPitch As Double
Dim outX(1 To 4) As Double, px(1 To 4) As Double
Dim inputLen As Double, divArm1 As Double, divArm2 As Double
Dim transLen As Double, nTaperSteps As Integer
Dim jogLen As Double
Dim stubLead As Double, stubHairH As Double
Dim finalFeedLen As Double
Dim marginX As Double, marginY As Double
Dim y1 As Double, y2 As Double, y3 As Double, yTransEnd As Double
Dim yJogEnd As Double, yStubEnd As Double, yPatchBottom As Double, yPatchTop As Double
Dim boardX As Double, boardY As Double
Dim steerState As Integer
Dim idxArr(1 To 4) As Integer
Dim maxD As Double

' ----------------------------------------------------------------------------
' MASTER EXECUTION BLOCK
' ----------------------------------------------------------------------------
Sub Main()
    ClearOldGeometry  ' Safely removes existing shapes before rebuilding

    ComputeParameters
    DefineMaterials
    BuildSubstrateAndGround
    BuildSIWNetwork
    BuildPhaseNetworkAndPatches
    BuildInputPort
    SetupUnitsBoundarySolverMonitor
    MsgBox "Model generated successfully." & vbCrLf & _
           "The input Trunk now uses solid walls to prevent inhomogeneous mesh errors.", vbInformation
End Sub

' ----------------------------------------------------------------------------
Sub ClearOldGeometry()
    On Error Resume Next
    Component.Delete "PCB"
    Component.Delete "SIW_Metal"
    Component.Delete "SIW_Vias"
    Port.Delete "1"
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
Sub ComputeParameters()
    c0 = 299.792458 ' mm*GHz

    f0 = 77: fmin = 76: fmax = 81
    er = 2.2: hSub = 0.127: tCu = 0.017
    tanD = 0.0013

    fc = 58: viaD = 0.3: viaP = 0.5

    Dim aEq As Double
    aEq = c0 / (2 * fc * Sqr(er))
    aSIW = aEq + (viaD ^ 2) / (0.95 * viaP)

    lambda0 = c0 / f0
    lambdaG = lambda0 / Sqr(1 - (fc / f0) ^ 2)
    elemD = lambda0 / 2

    patchW = (c0 / (2 * f0)) * Sqr(2 / (er + 1))
    Dim eeffP As Double, dL As Double, Leff As Double
    eeffP = (er + 1) / 2 + (er - 1) / 2 * (1 + 12 * hSub / patchW) ^ (-0.5)
    dL = 0.412 * hSub * (eeffP + 0.3) * (patchW / hSub + 0.264) / _
         ((eeffP - 0.258) * (patchW / hSub + 0.8))
    Leff = c0 / (2 * f0 * Sqr(eeffP))
    patchL = Leff - 2 * dL

    msW = 0.38
    Dim Wh As Double
    Wh = msW / hSub
    msEeff = (er + 1) / 2 + (er - 1) / 2 * (1 / Sqr(1 + 12 / Wh))
    lambdaMS = lambda0 / Sqr(msEeff)
    phaseDelta = lambdaMS / 4

    outPitch = aSIW + 2 * viaP
    outX(1) = -1.5 * outPitch: outX(2) = -0.5 * outPitch
    outX(3) = 0.5 * outPitch:  outX(4) = 1.5 * outPitch
    px(1) = -1.5 * elemD: px(2) = -0.5 * elemD
    px(3) = 0.5 * elemD:  px(4) = 1.5 * elemD

    inputLen = 1.5: divArm1 = 1.8: divArm2 = 1.8
    transLen = 1.2: nTaperSteps = 5
    jogLen = 1#
    stubLead = 0.25: stubHairH = 0.5
    finalFeedLen = 0.35
    marginX = 1#: marginY = 1#

    y1 = inputLen
    y2 = y1 + divArm1
    y3 = y2 + divArm2
    yTransEnd = y3 + transLen
    yJogEnd = yTransEnd + jogLen
    yStubEnd = yJogEnd + (stubLead + msW + stubHairH + msW + stubLead)
    yPatchBottom = yStubEnd + finalFeedLen
    yPatchTop = yPatchBottom + patchL

    steerState = 0
    Select Case steerState
        Case 0
            idxArr(1) = 0: idxArr(2) = 0: idxArr(3) = 0: idxArr(4) = 0
        Case 1
            idxArr(1) = 0: idxArr(2) = 1: idxArr(3) = 2: idxArr(4) = 3
        Case -1
            idxArr(1) = 3: idxArr(2) = 2: idxArr(3) = 1: idxArr(4) = 0
    End Select

    maxD = 3 * phaseDelta / 2
    Dim extSIW As Double, extStub As Double
    extSIW = 1.5 * outPitch + aSIW / 2
    extStub = 1.5 * elemD + patchW / 2 + maxD
    boardX = 2 * (MaxD2(extSIW, extStub) + marginX)
    boardY = yPatchTop + marginY
End Sub

Function MaxD2(a As Double, b As Double) As Double
    If a > b Then MaxD2 = a Else MaxD2 = b
End Function
Function MinD2(a As Double, b As Double) As Double
    If a < b Then MinD2 = a Else MinD2 = b
End Function

' ----------------------------------------------------------------------------
Sub DefineMaterials()
    With Material
        .Reset
        .Name "RT5880"
        .Folder ""
        .Type "Normal"
        .Epsilon CStr(er)
        .Mu "1"
        .TanD CStr(tanD)
        .TanDFreq CStr(f0)
        .TanDGiven "True"
        .Colour "0.85", "0.9", "0.75"
        .Create
    End With

    With Material
        .Reset
        .Name "Copper"
        .Folder ""
        .Type "Lossy metal"
        .Sigma "5.8e7"
        .Colour "1", "0.55", "0.1"
        .Create
    End With
End Sub

' ----------------------------------------------------------------------------
Sub BuildSubstrateAndGround()
    With Brick
        .Reset
        .Name "Ground"
        .Component "PCB"
        .Material "Copper"
        .Xrange -boardX / 2, boardX / 2
        .Yrange 0, boardY
        .Zrange -tCu, 0
        .Create
    End With

    With Brick
        .Reset
        .Name "Substrate"
        .Component "PCB"
        .Material "RT5880"
        .Xrange -boardX / 2, boardX / 2
        .Yrange 0, boardY
        .Zrange 0, hSub
        .Create
    End With
End Sub

' ----------------------------------------------------------------------------
Sub AddTopMetal(nm As String, x1 As Double, x2 As Double, y1_ As Double, y2_ As Double)
    With Brick
        .Reset
        .Name nm
        .Component "SIW_Metal"
        .Material "Copper"
        .Xrange x1, x2
        .Yrange y1_, y2_
        .Zrange hSub, hSub + tCu
        .Create
    End With
End Sub

Sub AddVia(nm As String, cx As Double, cy As Double)
    With Cylinder
        .Reset
        .Name nm
        .Component "SIW_Vias"
        .Material "Copper"
        .OuterRadius viaD / 2
        .InnerRadius 0
        .Axis "z"
        .Zrange -tCu, hSub + tCu
        .Xcenter cx
        .Ycenter cy
        .Create
    End With
End Sub

Sub FenceVertical(tag As String, cx As Double, yA As Double, yB As Double, trimA As Double, trimB As Double)
    Dim yStart As Double, yEnd As Double, n As Integer, i As Integer, yy As Double
    yStart = yA + trimA
    yEnd = yB - trimB
    If yEnd <= yStart Then Exit Sub
    n = Int((yEnd - yStart) / viaP)
    For i = 0 To n
        yy = yStart + i * viaP
        If yy > yEnd Then Exit For
        AddVia "V_" & tag & "_L_" & i, cx - aSIW / 2, yy
        AddVia "V_" & tag & "_R_" & i, cx + aSIW / 2, yy
    Next i
End Sub

Sub FenceHorizontal(tag As String, cy As Double, xA As Double, xB As Double, trimA As Double, trimB As Double)
    Dim xStart As Double, xEnd As Double, n As Integer, i As Integer, xx As Double
    xStart = xA + trimA
    xEnd = xB - trimB
    If xEnd <= xStart Then Exit Sub
    n = Int((xEnd - xStart) / viaP)
    For i = 0 To n
        xx = xStart + i * viaP
        If xx > xEnd Then Exit For
        AddVia "V_" & tag & "_T_" & i, xx, cy - aSIW / 2
        AddVia "V_" & tag & "_B_" & i, xx, cy + aSIW / 2
    Next i
End Sub

' ----------------------------------------------------------------------------
Sub BuildSIWNetwork()
    Dim half As Double
    half = aSIW / 2
    Dim branch1X As Double
    branch1X = outPitch

    AddTopMetal "Trunk", -half, half, 0, y1 + half

    ' --- HOMOGENEOUS PORT FIX ---
    ' Replaces the first 1.5mm of vias with solid walls for the port cross-section
    With Brick
        .Reset
        .Name "Trunk_SolidWall_L"
        .Component "SIW_Vias"
        .Material "Copper"
        .Xrange -half - viaD / 2, -half + viaD / 2
        .Yrange 0, y1
        .Zrange -tCu, hSub + tCu
        .Create
    End With

    With Brick
        .Reset
        .Name "Trunk_SolidWall_R"
        .Component "SIW_Vias"
        .Material "Copper"
        .Xrange half - viaD / 2, half + viaD / 2
        .Yrange 0, y1
        .Zrange -tCu, hSub + tCu
        .Create
    End With
    ' ----------------------------

    AddTopMetal "Bus1", -branch1X - half, branch1X + half, y1 - half, y1 + half
    FenceHorizontal "Bus1L", y1, -branch1X, 0, 0, half
    FenceHorizontal "Bus1R", y1, 0, branch1X, half, 0

    AddTopMetal "Arm1L", -branch1X - half, -branch1X + half, y1 - half, y2 + half
    AddTopMetal "Arm1R", branch1X - half, branch1X + half, y1 - half, y2 + half
    FenceVertical "Arm1L", -branch1X, y1, y2, half, half
    FenceVertical "Arm1R", branch1X, y1, y2, half, half

    AddTopMetal "Bus2L", outX(1) - half, outX(2) + half, y2 - half, y2 + half
    AddTopMetal "Bus2R", outX(3) - half, outX(4) + half, y2 - half, y2 + half
    FenceHorizontal "Bus2L1", y2, outX(1), -branch1X, 0, half
    FenceHorizontal "Bus2L2", y2, -branch1X, outX(2), half, 0
    FenceHorizontal "Bus2R1", y2, outX(3), branch1X, 0, half
    FenceHorizontal "Bus2R2", y2, branch1X, outX(4), half, 0

    Dim i As Integer
    For i = 1 To 4
        AddTopMetal "Arm2_" & i, outX(i) - half, outX(i) + half, y2 - half, y3
        FenceVertical "Arm2_" & i, outX(i), y2, y3, half, 0
    Next i
End Sub

' ----------------------------------------------------------------------------
Sub BuildPhaseNetworkAndPatches()
    Dim i As Integer
    For i = 1 To 4
        BuildOneBranch i
    Next i
End Sub

Sub BuildOneBranch(i As Integer)
    Dim cxSIW As Double, cxPatch As Double
    cxSIW = outX(i)
    cxPatch = px(i)

    Dim k As Integer, w0 As Double, stepLen As Double, yy As Double
    stepLen = transLen / nTaperSteps
    For k = 0 To nTaperSteps - 1
        w0 = aSIW + (msW - aSIW) * k / nTaperSteps
        yy = y3 + k * stepLen
        AddTopMetal "Taper_" & i & "_" & k, cxSIW - w0 / 2, cxSIW + w0 / 2, yy, yy + stepLen
    Next k

    Dim yJogMid As Double
    yJogMid = yTransEnd + jogLen / 2
    AddTopMetal "Jog_" & i & "_a", cxSIW - msW / 2, cxSIW + msW / 2, yTransEnd, yJogMid
    AddTopMetal "Jog_" & i & "_b", MinD2(cxSIW, cxPatch) - msW / 2, MaxD2(cxSIW, cxPatch) + msW / 2, _
                yJogMid - msW / 2, yJogMid + msW / 2
    AddTopMetal "Jog_" & i & "_c", cxPatch - msW / 2, cxPatch + msW / 2, yJogMid, yJogEnd

    Dim n As Integer, D As Double, y As Double
    n = idxArr(i)
    D = n * phaseDelta / 2

    y = yJogEnd
    AddTopMetal "Stub_" & i & "_lead1", cxPatch - msW / 2, cxPatch + msW / 2, y, y + stubLead
    y = y + stubLead

    If D > 0.0001 Then
        AddTopMetal "Stub_" & i & "_br1", cxPatch - msW / 2, cxPatch + D + msW / 2, y, y + msW
        Dim yTop As Double
        yTop = y + msW
        AddTopMetal "Stub_" & i & "_arm", cxPatch + D - msW / 2, cxPatch + D + msW / 2, y, yTop + stubHairH
        AddTopMetal "Stub_" & i & "_br2", cxPatch - msW / 2, cxPatch + D + msW / 2, yTop + stubHairH, yTop + stubHairH + msW
        y = yTop + stubHairH + msW
    Else
        AddTopMetal "Stub_" & i & "_straight", cxPatch - msW / 2, cxPatch + msW / 2, y, y + (msW + stubHairH + msW)
        y = y + (msW + stubHairH + msW)
    End If

    AddTopMetal "Stub_" & i & "_lead2", cxPatch - msW / 2, cxPatch + msW / 2, y, y + stubLead
    y = y + stubLead

    AddTopMetal "Feed_" & i, cxPatch - msW / 2, cxPatch + msW / 2, y, yPatchBottom
    AddTopMetal "Patch_" & i, cxPatch - patchW / 2, cxPatch + patchW / 2, yPatchBottom, yPatchTop
End Sub

' ----------------------------------------------------------------------------
Sub BuildInputPort()
    Dim portX As Double
    portX = (aSIW / 2) + viaP

    With Port
        .Reset
        .PortNumber "1"
        .Label "SIW_INPUT"
        .Folder ""
        .NumberOfModes "1"
        .AdjustPolarization "False"
        .PolarizationAngle "0"
        .ReferencePlaneDistance "0"
        .Coordinates "Free"
        .Orientation "ymin"
        .PortOnBound "True"
        .ClipPickedPortToBound "False"
        .Xrange CStr(-portX), CStr(portX)
        .Yrange "0", "0"
        .Zrange CStr(-tCu), CStr(hSub + tCu)
        .XrangeAdd "0", "0"
        .YrangeAdd "0", "0"
        .ZrangeAdd "0", "0"
        .SingleEnded "False"
        .WaveguideMonitor "False"
        .Create
    End With
End Sub

' ----------------------------------------------------------------------------
Sub SetupUnitsBoundarySolverMonitor()
    With Units
        .Geometry "mm"
        .Frequency "GHz"
        .Time "ns"
    End With

    With Boundary
        .Xmin "expanded open"
        .Xmax "expanded open"
        .Ymin "open"
        .Ymax "expanded open"
        .Zmin "expanded open"
        .Zmax "expanded open"
        .ApplyInAllDirections "False"
    End With

    With Background
        .YminSpace "0.0"
    End With

    On Error Resume Next
    ChangeSolverType ("HF Time Domain")
    On Error GoTo 0

    Solver.FrequencyRange CStr(fmin), CStr(fmax)

    With Monitor
        .Reset
        .Name "farfield (f=77)"
        .Domain "Frequency"
        .FieldType "Farfield"
        .Frequency "77"
        .Create
    End With
End Sub
