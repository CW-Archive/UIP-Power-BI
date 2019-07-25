Attribute VB_Name = "Functions"
Function Measurements_Lookup(MID_Input As String, MID_Column As String)
    Application.Volatile
    
    Split_MIDs = Split(MID_Input, ",")

    CountTable_Array = Range("CountTable").ListObject.DataBodyRange
    LengthTable_Array = Range("LengthTable").ListObject.DataBodyRange
    Wall_AreaTable_Array = Range("Wall_AreaTable").ListObject.DataBodyRange
    AreaTable_Array = Range("AreaTable").ListObject.DataBodyRange
    VolumeTable_Array = Range("VolumeTable").ListObject.DataBodyRange
    MasonryTable_Array = Range("MasonryTable").ListObject.DataBodyRange
    
    For Each m In Split_MIDs
        Select Case Left(m, 1)
            Case Is = "C"
                For i = 1 To UBound(CountTable_Array)
                    If CountTable_Array(i, 1) = m Then
                        If MID_Column = "Area" Or MID_Column = "Description" Then
                                If Measurements_Lookup = CountTable_Array(i, Range("CountTable").ListObject.ListColumns(MID_Column).Index) Then
                                    Else
                                    If Measurements_Lookup = "" Then
                                        Measurements_Lookup = CountTable_Array(i, Range("CountTable").ListObject.ListColumns(MID_Column).Index)
                                        Else
                                        Measurements_Lookup = Measurements_Lookup + " & " + CountTable_Array(i, Range("CountTable").ListObject.ListColumns(MID_Column).Index)
                                    End If
                                End If
                            Else
                            Measurements_Lookup = Measurements_Lookup + CountTable_Array(i, Range("CountTable").ListObject.ListColumns(MID_Column).Index)
                        End If
                    End If
                Next i
            Case Is = "L"
                For i = 1 To UBound(LengthTable_Array)
                    If LengthTable_Array(i, 1) = m Then
                        If MID_Column = "Area" Or MID_Column = "Description" Then
                                If Measurements_Lookup = LengthTable_Array(i, Range("LengthTable").ListObject.ListColumns(MID_Column).Index) Then
                                    Else
                                    If Measurements_Lookup = "" Then
                                        Measurements_Lookup = LengthTable_Array(i, Range("LengthTable").ListObject.ListColumns(MID_Column).Index)
                                        Else
                                        Measurements_Lookup = Measurements_Lookup + " & " + LengthTable_Array(i, Range("LengthTable").ListObject.ListColumns(MID_Column).Index)
                                    End If
                                End If
                            Else
                            Measurements_Lookup = Measurements_Lookup + LengthTable_Array(i, Range("LengthTable").ListObject.ListColumns(MID_Column).Index)
                        End If
                    End If
                Next i
            Case Is = "W"
                For i = 1 To UBound(Wall_AreaTable_Array)
                    If Wall_AreaTable_Array(i, 1) = m Then
                        If MID_Column = "Area" Or MID_Column = "Description" Then
                                If Measurements_Lookup = Wall_AreaTable_Array(i, Range("Wall_AreaTable").ListObject.ListColumns(MID_Column).Index) Then
                                    Else
                                    If Measurements_Lookup = "" Then
                                        Measurements_Lookup = Wall_AreaTable_Array(i, Range("Wall_AreaTable").ListObject.ListColumns(MID_Column).Index)
                                        Else
                                        Measurements_Lookup = Measurements_Lookup + " & " + Wall_AreaTable_Array(i, Range("Wall_AreaTable").ListObject.ListColumns(MID_Column).Index)
                                    End If
                                End If
                            Else
                            Measurements_Lookup = Measurements_Lookup + Wall_AreaTable_Array(i, Range("Wall_AreaTable").ListObject.ListColumns(MID_Column).Index)
                        End If
                    End If
                Next i
            Case Is = "A"
                For i = 1 To UBound(AreaTable_Array)
                    If AreaTable_Array(i, 1) = m Then
                        If MID_Column = "Area" Or MID_Column = "Description" Then
                                If Measurements_Lookup = AreaTable_Array(i, Range("AreaTable").ListObject.ListColumns(MID_Column).Index) Then
                                    Else
                                    If Measurements_Lookup = "" Then
                                        Measurements_Lookup = AreaTable_Array(i, Range("AreaTable").ListObject.ListColumns(MID_Column).Index)
                                        Else
                                        Measurements_Lookup = Measurements_Lookup + " & " + AreaTable_Array(i, Range("AreaTable").ListObject.ListColumns(MID_Column).Index)
                                    End If
                                End If
                            Else
                            Measurements_Lookup = Measurements_Lookup + AreaTable_Array(i, Range("AreaTable").ListObject.ListColumns(MID_Column).Index)
                        End If
                    End If
                Next i
            Case Is = "V"
                For i = 1 To UBound(VolumeTable_Array)
                    If VolumeTable_Array(i, 1) = m Then
                        If MID_Column = "Area" Or MID_Column = "Description" Then
                                If Measurements_Lookup = VolumeTable_Array(i, Range("VolumeTable").ListObject.ListColumns(MID_Column).Index) Then
                                    Else
                                    If Measurements_Lookup = "" Then
                                        Measurements_Lookup = VolumeTable_Array(i, Range("VolumeTable").ListObject.ListColumns(MID_Column).Index)
                                        Else
                                        Measurements_Lookup = Measurements_Lookup + " & " + VolumeTable_Array(i, Range("VolumeTable").ListObject.ListColumns(MID_Column).Index)
                                    End If
                                End If
                            Else
                            Measurements_Lookup = Measurements_Lookup + VolumeTable_Array(i, Range("VolumeTable").ListObject.ListColumns(MID_Column).Index)
                        End If
                    End If
                Next i
            Case Is = "M"
                For i = 1 To UBound(MasonryTable_Array)
                    If MasonryTable_Array(i, 1) = m Then
                        If MID_Column = "Area" Or MID_Column = "Description" Then
                                If Measurements_Lookup = MasonryTable_Array(i, Range("MasonryTable").ListObject.ListColumns(MID_Column).Index) Then
                                    Else
                                    If Measurements_Lookup = "" Then
                                        Measurements_Lookup = MasonryTable_Array(i, Range("MasonryTable").ListObject.ListColumns(MID_Column).Index)
                                        Else
                                        Measurements_Lookup = Measurements_Lookup + " & " + MasonryTable_Array(i, Range("MasonryTable").ListObject.ListColumns(MID_Column).Index)
                                    End If
                                End If
                            Else
                            Measurements_Lookup = Measurements_Lookup + MasonryTable_Array(i, Range("MasonryTable").ListObject.ListColumns(MID_Column).Index)
                        End If
                    End If
                Next i
        End Select
    Next m
End Function
