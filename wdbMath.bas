Option Compare Database
Option Explicit

Function pi() As Double
On Error GoTo err_handler

pi = 4 * Atn(1)

Exit Function
err_handler:
    Call handleError("wdbMath", "pi", Err.DESCRIPTION, Err.number)
End Function

Function Asin(X) As Double
On Error GoTo err_handler

Select Case X
    Case 1
        Asin = pi / 2
    Case -1
        Asin = (3 * pi) / 2
    Case Else
        Asin = Atn(X / Sqr(-X * X + 1))
End Select

Exit Function
err_handler:
    Call handleError("wdbMath", "Asin", Err.DESCRIPTION, Err.number)
End Function

Function Acos(X) As Double
On Error GoTo err_handler

Select Case X
    Case 1
        Acos = 0
    Case -1
        Acos = pi
    Case Else
        Acos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Select

Exit Function
err_handler:
    Call handleError("wdbMath", "Acos", Err.DESCRIPTION, Err.number)
End Function