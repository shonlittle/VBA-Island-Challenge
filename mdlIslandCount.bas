Attribute VB_Name = "mdlIslandCount"
' --
' The Boston Conulting Group
' Data and Analytics Services (DaAS)
' Shon Little
' July 12, 2016
' --

' Options
Option Explicit

' Entry Point
Public Sub IslandCount()
    Dim found           As Boolean
    Dim intCount        As Integer
    Dim r1              As Long
    Dim c1              As Long
    Dim r2              As Long
    Dim c2              As Long
    Dim dblStartTime    As Double
    Dim strFinish       As String
    Dim strTime         As String
    Dim varIsland()     As Variant
    
    ' Initialize
    dblStartTime = Timer
    strFinish = "Finished!"
    intCount = 1
    varIsland = ws_Map.Range("A1:K18").Value2
    
    ' Loop rows
    For r1 = 1 To 18
        ' Loop columns
        For c1 = 1 To 11
            ' Check if land
            If varIsland(r1, c1) = 1 Then
                ' Increment count
                intCount = intCount + 1
                ' Assign Island index
                varIsland(r1, c1) = intCount
                ' Find all connecting lands
                Do
                    ' Reset found flag
                    found = False
                    ' Reloop  rows
                    For r2 = 1 To 18
                        ' Reloop columns
                        For c2 = 1 To 11
                            ' Check for land
                            If varIsland(r2, c2) = 1 Then
                                ' Check Left
                                If c2 > LBound(varIsland, 2) Then
                                    If varIsland(r2, c2 - 1) = intCount Then varIsland(r2, c2) = intCount
                                End If
                                ' Check Up
                                If r2 > LBound(varIsland, 1) Then
                                    If varIsland(r2 - 1, c2) = intCount Then varIsland(r2, c2) = intCount
                                End If
                                ' Check Right
                                If c2 < UBound(varIsland, 2) Then
                                    If varIsland(r2, c2 + 1) = intCount Then varIsland(r2, c2) = intCount
                                End If
                                ' Check Down
                                If r2 < UBound(varIsland, 1) Then
                                    If varIsland(r2 + 1, c2) = intCount Then varIsland(r2, c2) = intCount
                                End If
                                ' Record if conneced land found
                                If varIsland(r2, c2) = intCount Then found = True
                            End If
                        Next c2
                    Next r2
                    ' Loop until no more connected lands found
                Loop Until found = False
            End If
        Next c1
    Next r1

    ' Finish
    strFinish = strFinish & " Island Count: " & CStr(intCount - 1)
    strTime = vbCrLf & "Time: " & Format((Timer - dblStartTime) / 86400, "hh:mm:ss")
       
finally:
    ' Finish message
    MsgBox strFinish & strTime, vbOKOnly + vbInformation, "Finished"
    Exit Sub
    
catch:
    ' Error handler
    MsgBox "Error: " & Err.Description & " in IslandCount", vbCritical, "Error " & Err.Number
    Resume finally
End Sub
