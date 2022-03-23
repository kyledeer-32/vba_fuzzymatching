Attribute VB_Name = "FuzzyMatching_Development"


'Public Function FuzzyMatching_V1(ByRef targetS As String, ByRef MatchRange As Range) As String
Public Sub FuzMat_Testing()

    'begin remove
    Dim targetS As String
    targetS = "kyle"
    'end remove
    
    'begin remove
    Dim rng As Range
    Worksheets("Sheet1").Activate
    Set rng = Range("A1:A31")
    'end remove
    
    
    Dim MatchArray() As Variant
    MatchArray = rng.Value2
    
    
    'Must add "Microsoft Scripting Runtime" library for dictionaries
    Dim EditDistance_Dict As Object
    Set EditDistance_Dict = New Scripting.Dictionary
    
    Dim S2 As String
    Dim EditD As Long
    
    Dim x As Long
    Dim y As Long

'start remove
    Dim test As Long
    test = 1
'end remove
    
    For x = LBound(MatchArray, 1) To UBound(MatchArray, 1):
        For y = LBound(MatchArray, 2) To UBound(MatchArray, 2):
            If EditDistance_Dict.Exists(MatchArray(x, y)) Then
                GoTo Skip_Add
            End If
                        
            S2 = MatchArray(x, y)
            EditD = LevD(targetS, S2)
            EditDistance_Dict.Add MatchArray(x, y), EditD
            
            'Use for Debugging
            'Debug.Print test; ": x = "; x; "   y = "; y; " MatchArray Value = "; MatchArray(x, y); " S2 = "; S2; "targetS = "; targetS
            'Debug.Print "Edit Distance = "; EditD
            'test = test + 1
            'end Debugging

Skip_Add:

        Next y
    Next x
   
   
Dim key As Variant
Dim Closest_Match As String
Dim Distance As Long

Distance = 1000

    For Each key In EditDistance_Dict.Keys
        If EditDistance_Dict(key) < Distance Then
            Closest_Match = key
            Distance = EditDistance_Dict(key)
        End If
        Debug.Print key, EditDistance_Dict(key)
    Next key
        
        
        
Debug.Print "Retrieve Dict Item using Key - kite = "; EditDistance_Dict("kite")


Debug.Print "lol"

End Sub
'End Function










'Public Function LevD(ByRef String1 As String, ByRef String2 As String) As Long
Public Sub testinglevdev()
    
    
'remove start

    Dim DebugCol As Object
    Set DebugCol = New Collection

    Dim String1 As String
    Dim String2 As String
    
    String1 = "BAASDJ FJADSGF"
    String2 = "hat"

'remove end

    String1 = Trim(String1)
    String2 = Trim(String2)
    
    String1 = LCase(String1)
    String2 = LCase(String2)
    
    Dim count_S1 As Long
    Dim count_S2 As Long
    
    count_S1 = Len(String1)
    count_S2 = Len(String2)
    
    Dim LevMatrix() As Long
    ReDim LevMatrix(0 To count_S1, 0 To count_S2)
    
    Dim i As Long
    Dim j As Long
    
        For i = 0 To count_S1:
            LevMatrix(i, 0) = i
            Next i
        
        For j = 0 To count_S2:
            LevMatrix(0, j) = j
            Next j
    
    Dim CharCol1 As Object
    Set CharCol1 = New Collection
    
    Dim CharCol2 As Object
    Set CharCol2 = New Collection
    
    Dim x As Long
    Dim y As Long
    
        For x = 1 To count_S1:
            CharCol1.Add Mid(String1, x, 1)
            Next x
        
        For y = 1 To count_S2:
            CharCol2.Add Mid(String2, y, 1)
            Next y
            
    Dim Cost As Long
    
    Dim Calc1 As Long
    Dim Calc2 As Long
    Dim Calc3 As Long
    
        For i = 1 To count_S1:
            For j = 1 To count_S2:
                If CharCol1(i) <> CharCol2(j) Then
                    Cost = 1
                Else
                    Cost = 0
                End If
                              
                Calc1 = LevMatrix(i - 1, j) + 1
                Calc2 = LevMatrix(i, j - 1) + 1
                Calc3 = LevMatrix(i - 1, j - 1) + Cost

                LevMatrix(i, j) = WorksheetFunction.Min(Calc1, Calc2, Calc3)
                
                'start remove
                
                DebugCol.Add Cost
                
                'end remove
                               
            Next j
        Next i
    
    
'Remove start
    
Debug.Print "For "; String1; " and "; String2; " : "; LevMatrix(count_S1, count_S2)
    
'Remove end
    
    'LevD = LevMatrix(count_S1, count_S2)
    
End Sub


'Public Function LevD(ByRef String1 As String, ByRef String2 As String) As Long

Public Sub LevD_test()

Dim String1 As String
Dim String2 As String

String1 = "aa"
String2 = "aaa"

Dim TextCompare As Long
Dim BinaryCompare As Long

TextCompare = StrComp(String1, String2, vbTextCompare)
Debug.Print "TextCompare returned: "; TextCompare

BinaryCompare = StrComp(String1, String2, vbBinaryCompare)
Debug.Print "BinaryCompare returned: "; BinaryCompare

End Sub

Public Sub Array_Learning()

Dim OneDim_arr(2) As Long
Dim TwoDim_arr(2, 4) As Long

TwoDim_arr(1, 3) = 10

Debug.Print "LOL"

End Sub

