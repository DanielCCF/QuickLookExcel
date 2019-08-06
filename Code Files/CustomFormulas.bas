Attribute VB_Name = "CustomFormulas"

Public Function ShtNumberFromRange(ByVal rngTarget As Range) As Long
    
    'Objective: Return the number of a sheet from a given Range.
    '           Thid Function was created due to compatibility
    '           isues.
    
    ShtNumberFromRange = rngTarget.Worksheet.Index
    
End Function
