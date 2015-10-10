Attribute VB_Name = "Module1"
Public Function degf_to_degc(degf As Integer) As Integer
degf_to_degc = CInt((degf - 32) * 5 / 9)
End Function
