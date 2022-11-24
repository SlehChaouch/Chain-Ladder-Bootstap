
'  Simulation in a spreadsheet
'  Written by Peter England 23/12/97
 
Sub Simulate()
'MsgBox "Nous sommes le " & Date
    Dim sumx(), sumx2(), outval()
    Dim mean(), sd()
    Dim sim As Integer
    'Application.ScreenUpdating = False
'    Application.Calculation = xlManual
    Range("Input1").Select
    
' if input range is empty, stop
    If IsEmpty(ActiveCell.Value) = True Then
    MsgBox ("No simulation inputs defined")
    Exit Sub
    End If
'  otherwise, count number of inputs and continue
    ninputs = 1
    ActiveCell.Offset(0, 1).Activate
    Do While IsEmpty(ActiveCell.Value) = False
        ninputs = ninputs + 1
        ActiveCell.Offset(0, 1).Activate
    Loop
  

    simnumber = Range("bsnum")

' set dimension of arrays and perform the simulation
    ReDim sumx(1 To ninputs)
    ReDim sumx2(1 To ninputs)
    ReDim mean(1 To ninputs)
    ReDim sd(1 To ninputs)
    ReDim outval(1 To simnumber, 1 To ninputs)
    For j = 1 To ninputs
        sumx(j) = 0
        sumx2(j) = 0
    Next j
    Application.DisplayStatusBar = True
    For sim = 1 To simnumber
        Calculate
        Application.StatusBar = "Simulation " & sim
        For j = 1 To ninputs
            outval(sim, j) = Range("Input1").Cells(1, j)
            sumx(j) = sumx(j) + Range("Input1").Cells(1, j)
            sumx2(j) = sumx2(j) + (Range("Input1").Cells(1, j)) ^ 2
        Next j
    Next sim
    Application.StatusBar = False
    
'  calculate mean, standard deviation etc. of all simulated inputs
    For j = 1 To ninputs
        mean(j) = sumx(j) / simnumber
        sd(j) = Sqr((simnumber / (simnumber - 1)) * (sumx2(j) / simnumber - mean(j) ^ 2))
        Range("Input1").Select
        ActiveCell.Offset(1, j - 1).Activate
        ActiveCell.Value = mean(j)
        ActiveCell.Offset(1, 0).Activate
        ActiveCell.Value = sd(j)
    Next j
    
'Output values as columns on new sheet, if applicable
    If Range("saveres") = True Then
        'Sheets.Add
        
        Sheets("SAMPLE").Cells.Clear
        Application.Calculation = xlCalculationManual
        For sim = 1 To simnumber
        Sheets("SAMPLE").Cells(sim, 1) = sim
            For j = 1 To ninputs
                Sheets("SAMPLE").Cells(sim, j + 1) = outval(sim, j)
            Next j
            'Sheets("SAMPLE").Cells(sim, 1) = outval(sim, ninputs)
        Next sim
        
        Sheets("SAMPLE").Activate
        
        Columns("K:K").Select
        Selection.Sort Key1:=Range("K1"), Order1:=xlAscending, Header:=xlNo
        Columns("B:B").Select
        Selection.Sort Key1:=Range("B1"), Order1:=xlAscending, Header:=xlNo
        Columns("C:C").Select
        Selection.Sort Key1:=Range("C1"), Order1:=xlAscending, Header:=xlNo
        Columns("D:D").Select
        Selection.Sort Key1:=Range("D1"), Order1:=xlAscending, Header:=xlNo
        Columns("E:E").Select
        Selection.Sort Key1:=Range("E1"), Order1:=xlAscending, Header:=xlNo
        Columns("F:F").Select
        Selection.Sort Key1:=Range("F1"), Order1:=xlAscending, Header:=xlNo
        Columns("G:G").Select
        Selection.Sort Key1:=Range("G1"), Order1:=xlAscending, Header:=xlNo
        Columns("H:H").Select
        Selection.Sort Key1:=Range("H1"), Order1:=xlAscending, Header:=xlNo
        Columns("I:I").Select
        Selection.Sort Key1:=Range("I1"), Order1:=xlAscending, Header:=xlNo
        Columns("J:J").Select
        Selection.Sort Key1:=Range("J1"), Order1:=xlAscending, Header:=xlNo
        
        Sheets("Bootstrap").Activate
        Application.Calculation = xlCalculationAutomatic
    End If
Reset:
'    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
End Sub


Function GRnd3(a, beta)
Application.Volatile

'Return sample from Gamma(a, b) distribution
'parameterised as mean=a*b, var=a*b^2
'Note: Gamma(a, 1) is standard Gamma
Dim e As Double
Dim b As Double
Dim U As Double
Dim U1 As Double
Dim U2 As Double
Dim p As Double
Dim y As Double
Dim x As Double
Dim ea As Double
Dim q As Double
Dim theta As Double
Dim d As Double
Dim v As Double
Dim z As Double
Dim w As Double

Do
    If (a < 1) And (a > 0) Then
        e = Exp(1)
        b = (e + a) / e
        U1 = URnd()
        p = b * U1
        If p > 1 Then
            y = -Log((b - p) / a)
            U2 = URnd()
            If (U2 < y ^ (a - 1)) Or (U2 = y ^ (a - 1)) Then
                x = y
            Else
                x = -100
            End If
        Else
            y = p ^ (1 / a)
            U2 = URnd()
            If (U2 < Exp(-y)) Or (U2 = Exp(-y)) Then
                x = y
            Else
                x = -100
            End If
        End If
    ElseIf a = 1 Then
        U = URnd()
        x = -Log(U)
    ElseIf a > 1 Then
        ea = 1 / Sqr(2 * a - 1)
        b = a - Log(4)
        q = a + 1 / ea
        theta = 4.5
        d = 1 + Log(theta)
        U1 = URnd()
        U2 = URnd()
        v = ea * Log(U1 / (1 - U1))
        y = a * Exp(v)
        z = U2 * U1 * U1
        w = b + q * v - y
        If (w + d - theta * z > 0) Or (w + d - theta * z = 0) Then
            x = y
        ElseIf (w > Log(z)) Or (w = Log(z)) Then
            x = y
        Else
            x = -100
        End If
    Else
        x = -999
    End If
    
Loop Until Not (x = -100)

GRnd3 = x * beta
 
End Function

Function URnd(Optional Seed)
' Return sample from Uniform(0,1)
'
' Note: take care to initialise seed with first call
' otherwise Last=Uni=0 on first call.
' See subroutine Initialise for setting seed.
Dim a As Double
Dim c As Double
Dim m As Double
Dim Last As Double
Static Uni As Double

a = 16807
c = 1.414
m = 2 ^ 31 - 1

If IsMissing(Seed) Then
    Last = Uni
Else
    Last = Seed
End If

Uni = (a * Last + c) - Int(a * Last + c)
URnd = Uni

End Function

Function lngamm(xx As Double)
Application.Volatile

Dim j As Integer
Dim ser As Double
Dim stp As Double
Dim tmp As Double
Dim x As Double
Dim y As Double
Dim cof() As Double
ReDim cof(1 To 6)

cof(1) = 76.1800917294715
cof(2) = -86.5053203294168
cof(3) = 24.0140982408309
cof(4) = -1.23173957245015
cof(5) = 1.20865097386618E-03
cof(6) = -5.395239384953E-06
stp = 2.506628274631
ser = 1.00000000019001

x = xx
y = x
tmp = x + 5.5
tmp = (x + 0.5) * Log(tmp) - tmp

For j = 1 To 6
    y = y + 1
    ser = ser + cof(j) / y
Next j

lngamm = tmp + Log(stp * ser / x)

End Function

Function PoiRnd(mu)
Application.Volatile

Dim xm As Double
Dim pi As Double
Dim alxm As Double
Dim em As Double
Dim g As Double
Dim sq As Double
Dim t As Double
Dim y As Double

pi = 3.14159265358979

If mu < 12 Then
    g = Exp(-mu)
    em = -1
    t = 1
    Do While t > g
        em = em + 1
        t = t * URnd
    Loop
Else
    sq = Sqr(2 * mu)
    alxm = Log(mu)
    g = mu * alxm - lngamm(mu + 1)
One:
    y = Tan(pi * URnd)
    em = sq * y + mu
    If em < 0 Then
        GoTo One
    End If
    em = Int(em)
    t = 0.9 * (1 + y ^ 2) * Exp(em * alxm - lngamm(em + 1) - g)
    If URnd > t Then
        GoTo One
    End If
End If

PoiRnd = em

End Function
