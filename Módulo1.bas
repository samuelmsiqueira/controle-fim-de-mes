Attribute VB_Name = "Módulo1"
Sub Confere()
    Dim fun As Integer
    Dim hrNtrabalhada As Date
    Dim minExtra As Double
    Dim dFalta As Integer
    Dim inss As Double
    Dim vale As Double
    Dim salario As Double
    Dim salarioTotal As Double
    Dim adiantamento As Double
    Dim calcinss As Double
    
    Sheets("Principal").Select
    fun = Range("A1").Value
    hrNtrabalhada = Range("H12").Value * 24
    minExtra = Range("J12").Value
    dFalta = Range("L12").Value
    
    
    Sheets("Funcionarios").Select
    salario = Cells(fun + 1, 3).Value
    adiantamento = Cells(fun + 1, 4).Value
    
    'CALCULO VALE
    vale = 230 / 30
    vale = vale * (30 - dFalta)
    
    'CALCULO SALARIO
    salario = salario / 220
    salario = salario * (220 - hrNtrabalhada)
    'CALCULO INSS
    If salario <= 1518 Then
        inss = salario * 0.075
    Else
        
        inss = salario - 1518
        MsgBox inss
        inss = inss * 0.09
        MsgBox inss
        MsgBox salario
        inss = inss + 113.85
    End If
    salario = salario - adiantamento
    salario = salario - inss
    salarioTotal = salario + vale
    Sheets("CONTROLE BANCO DE HORAS").Select
    Dim bdhoras As Double
    bdhoras = Cells(fun + 1, 4).Value
    Sheets("Principal").Select
    Range("C15").Value = inss
    Range("D15").Value = vale
    Range("E15").Value = salario
    Range("F15").Value = salarioTotal
    Range("C20").Value = bdhoras
    
End Sub

Sub Salvar()
    Dim fun As Integer
    Dim hrNtrabalhada As Double
    Dim minExtra As Double
    Dim dFalta As Integer
    Dim inss As Double
    Dim vale As Double
    Dim salario As Double
    Dim salarioTotal As Double
    Dim hrEx As Double
    Dim L As Integer
    
    Sheets("Principal").Select
    fun = Range("A1").Value
    hrNtrabalhada = Range("H12").Value
    minExtra = Range("J12").Value
    dFalta = Range("L12").Value
    inss = Range("C15").Value
    vale = Range("D15").Value
    salario = Range("E15").Value
    salarioTotal = Range("F15").Value
    hrEx = Range("J12").Value
    
    Sheets("CONTROLE FIM DE MÊS").Select
    L = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    L = L + 1
    Cells(L, 1).Value = fun
    Cells(L, 4).Value = hrNtrabalhada
    Cells(L, 5).Value = dFalta
    Cells(L, 2).Value = inss
    Cells(L, 3).Value = vale
    Cells(L, 6).Value = salarioTotal
    Cells(L, 7).Value = hrEx
    
    Sheets("CONTROLE BANCO DE HORAS").Select
    Cells(fun + 1, 1).Value = fun
    Dim calc As Double
    calc = Cells(fun + 1, 2).Value
    calc = calc + hrEx
    Cells(fun + 1, 2).Value = calc
    MsgBox calc
    calc = 0
    calc = Cells(fun + 1, 3).Value
    calc = calc + hrNtrabalhada
    Cells(fun + 1, 3).Value = calc
    Cells(fun + 1, 4).Value = Cells(fun + 1, 2).Value - Cells(fun + 1, 3).Value
    
    ' Ler do Excel (se célula tem hora)
    'hrNtrabalhada = Range("H12").Value * 24  ' converte hora Excel ? horas decimais
    'hrEx = Range("J12").Value * 24

    ' Somar horas:
    'calc = Cells(fun + 1, 2).Value * 24
    'calc = calc + hrEx
    'Cells(fun + 1, 2).Value = calc / 24      ' volta para formato hora
    'Cells(fun + 1, 2).NumberFormat = "[h]:mm"
    
    
    Sheets("Principal").Select
    Range("A1").Value = ""
    Range("H12").Value = ""
    Range("J12").Value = ""
    Range("L12").Value = ""
    Range("C15").Value = ""
    Range("D15").Value = ""
    Range("E15").Value = ""
    Range("F15").Value = ""
End Sub

