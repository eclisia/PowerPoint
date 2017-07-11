Attribute VB_Name = "PowerPointShape_treatment"
'Public variable
Private w As Double
Private h As Double
Private l As Double
Private t As Double
    
Sub CopyShapeDimensions()
    'Set to Zero from previous Value
    w = 0
    h = 0
    l = 0
    t = 0
    
    Dim Sld As Slide
    Dim selShp As Shape
    Dim Shp As Shape

    With ActiveWindow.Selection.ShapeRange(1)
            w = .Width
            h = .Height
            l = .Left
            t = .Top
    End With
    Debug.Print "dimension à copier sont W=" & w & " h=" & h & " l=" & l & " t=" & t


End Sub

Sub PastShapeDimensions()

    With ActiveWindow.Selection.ShapeRange(1)
            
            Debug.Print "dimension actuelles sont W=" & .Width & " h=" & .Height & " l=" & .Left & " t=" & .Top
            
            .Width = w
            .Height = h
            .Left = l
            .Top = t
            
            Debug.Print "Nouvelles dimension sont W=" & .Width & " h=" & .Height & " l=" & .Left & " t=" & .Top
    End With

End Sub

