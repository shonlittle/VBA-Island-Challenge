Attribute VB_Name = "mdlIslandCount"
' --
' The Boston Conulting Group
' Data and Analytics Services (DaAS)
' Shon Little
' July 12, 2016
' --

' Options
Option Explicit

' Settings
Private Const MODULE    As String = "mdlIslandCount"
Private Const DEBUGGING As Boolean = False


' Entry Point
Public Sub Main()
    Dim intIslandNum    As Integer
    Dim dblStartTime    As Double
    Dim strFinish       As String
    Dim strDebugMsg     As String
    Dim strTime         As String
        
    ' Set mode
    If Not DEBUGGING Then
        On Error GoTo catch
        Application.ScreenUpdating = False
    End If
    
    ' Initialize
    dblStartTime = Timer
    strFinish = "Finished!"
    
    ' Routine
    intIslandNum = CountIslands(ActiveSheet.Cells.SpecialCells(xlCellTypeConstants))

    ' Finish
    strFinish = strFinish & vbCrLf & "Island Count: " & CStr(intIslandNum)
    If DEBUGGING Then strDebugMsg = " with ""debugging"" turned on. " & vbCrLf & "Turning it off will improve performance."
    strTime = vbCrLf & "Time: " & Format((Timer - dblStartTime) / 86400, "hh:mm:ss")
       
finally:
    ' Reset application
    Application.ScreenUpdating = True
    ' Finish message
    MsgBox strFinish & strTime & strDebugMsg, vbOKOnly + vbInformation, "Finished"
    Exit Sub
    
catch:
    ' Error handler
    MsgBox "Error: " & Err.Description & " in " & MODULE & ".Main", vbCritical, "Error " & Err.Number
    Resume finally
End Sub

' Function to count islands
Private Function CountIslands(ByRef rngMap As Range) As Integer
    Dim found           As Boolean
    Dim intCount        As Integer
    Dim r1              As Long
    Dim c1              As Long
    Dim r2              As Long
    Dim c2              As Long
    Dim varMap()        As Variant

    ' Set mode
    If Not DEBUGGING Then On Error GoTo catch

    ' Initialize
    intCount = 1
    varMap = rngMap.Value2
    
    ' Loop rows
    For r1 = 1 To UBound(varMap, 1)
        ' Loop columns
        For c1 = 1 To UBound(varMap, 2)
            ' Check if land
            If varMap(r1, c1) = 1 Then
                ' Increment count
                intCount = intCount + 1
                ' Assign Island index
                varMap(r1, c1) = intCount
                ' Find all connecting lands
                Do
                    ' Reset found flag
                    found = False
                    ' Reloop  rows
                    For r2 = 1 To UBound(varMap, 1)
                        ' Reloop columns
                        For c2 = 1 To UBound(varMap, 2)
                            ' Check for land
                            If varMap(r2, c2) = 1 Then
                                ' Check Left
                                If c2 > LBound(varMap, 2) Then
                                    If varMap(r2, c2 - 1) = intCount Then varMap(r2, c2) = intCount
                                End If
                                ' Check Up
                                If r2 > LBound(varMap, 1) Then
                                    If varMap(r2 - 1, c2) = intCount Then varMap(r2, c2) = intCount
                                End If
                                ' Check Right
                                If c2 < UBound(varMap, 2) Then
                                    If varMap(r2, c2 + 1) = intCount Then varMap(r2, c2) = intCount
                                End If
                                ' Check Down
                                If r2 < UBound(varMap, 1) Then
                                    If varMap(r2 + 1, c2) = intCount Then varMap(r2, c2) = intCount
                                End If
                                ' Record if conneced land found
                                If varMap(r2, c2) = intCount Then found = True
                            End If
                        Next c2
                    Next r2
                    ' Loop until no more connected lands found
                Loop Until found = False
            End If
        Next c1
    Next r1
    
    ' Return
    CountIslands = intCount - 1
       
finally:
    Exit Function
    
catch:
    ' Error handler
    MsgBox "Error: " & Err.Description & " in " & MODULE & ".CountIslands", vbCritical, "Error " & Err.Number
    Resume finally
End Function
