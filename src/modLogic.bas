Attribute VB_Name = "modLogic"
Public Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long

Public RGBValue As Long
Public Tick As Long

Public Function RAND(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    RAND = Int((High - Low + 1) * Rnd) + Low
End Function
