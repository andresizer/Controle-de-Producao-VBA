Attribute VB_Name = "Módulo_1"
Option Explicit
Dim dataHoje As Byte
Dim d As Single
Dim n As Single
Dim i As Byte
Dim j As Single
Dim m As Byte
Dim y As Integer
Dim nomeDia As Byte
Dim diaSemana As String
Dim mesExt As String
Dim novaPlanilha As Worksheet
Dim pControle As Worksheet
Dim pProducao As Worksheet
Dim Vsobra As Range
Dim VcontDiario As Range
Dim btSobra As ShapeRange
Dim btDiario As ShapeRange
Dim btPrevisoes As ShapeRange
Dim btProducao As ShapeRange
Dim btImprimir As ShapeRange
Dim vMassas As Range
Dim vContMassas As Range
Dim vProducao As Range
Dim vContProd As Range
Dim nmassadas As Integer

Sub mediaDiaria()

Dim vMediaDiaria As Range
Dim vMediaSobra As Range



m = Worksheets.Count
Set pControle = Worksheets(m)
Set vMediaDiaria = pControle.Rows(20) 'ALTERAR AQUI
Set vMediaSobra = pControle.Rows(10)  'ALTERAR  AQUI
d = WorksheetFunction.Average(vMediaDiaria)
n = WorksheetFunction.Average(vMediaSobra)
Set pProducao = Planilha2


pProducao.Range("Q6").Value = d - n

End Sub

Sub saidaOntem()

m = Worksheets.Count
Set pControle = Worksheets(m)
Set pProducao = Planilha2


i = pControle.Range("A4").CurrentRegion.Columns.Count


If WorksheetFunction.IsNumber(pControle.Cells(20, i - 1).Value) Then 'ALTERAR AQUI
    j = pControle.Cells(20, i - 1).Value 'ALTERAR AQUI
Else
    pProducao.Range("Q7").Value = 0
    Exit Sub
End If
n = pControle.Cells(11, i).Value

d = j - n

pProducao.Range("Q7").Value = d

End Sub


Sub pNova()

i = Month(Date)
mesExt = MonthName(i, True)
y = Year(Date)



m = Worksheets.Count
Set pControle = Worksheets(m)

Set novaPlanilha = Worksheets.Add
    novaPlanilha.Move , pControle
    novaPlanilha.Name = "Controle " & mesExt & y
 
       
   pControle.Range("A:A").Copy novaPlanilha.Range("A1")
   pControle.Range("B49:B61").Copy novaPlanilha.Range("B49") 'Mudan�a pra p�scoa
   novaPlanilha.Columns("A:A").EntireColumn.AutoFit
   novaPlanilha.Activate
   ActiveWindow.zoom = 80
   
novaPlanilha.Range("A1").Value = mesExt

pControle.Visible = xlSheetHidden


Call sobra

End Sub

Sub sobra()
dataHoje = Day(Date)
d = Weekday(Date)
diaSemana = WeekdayName(d, True, vbSunday)
i = Month(Date)
mesExt = MonthName(i, True)
m = Worksheets.Count

Set pProducao = Planilha2
Set pControle = Worksheets(m)

Set btSobra = Planilha2.Shapes.Range(Array("botaoSobra"))
Set btDiario = Planilha2.Shapes.Range(Array("botaoMariland"))
Set btPrevisoes = Planilha2.Shapes.Range(Array("botaoPrevisoes"))
Set btProducao = Planilha2.Shapes.Range(Array("botaoRegistros"))
Set btImprimir = Planilha2.Shapes.Range(Array("botaoImprimir"))



If pControle.Range("A1").Value <> mesExt Then
    Call pNova
    Else
    n = pControle.Range("A4").CurrentRegion.Columns.Count + 1
    
        If pControle.Cells(2, n - 1).Value = dataHoje Then
                If MsgBox("Voc� deseja substuir os valores j� adicionados hoje?", vbYesNo, "Aten��o!") = vbYes Then
                        Set Vsobra = pControle.Cells(4, n - 1)
                        pProducao.Range("vSobra").Copy Vsobra
                    Else: Exit Sub
                End If
            Else
                Set Vsobra = pControle.Cells(4, n)

                Vsobra.Offset(-3, 0).Value = diaSemana
                Vsobra.Offset(-2, 0).Value = dataHoje
                pProducao.Activate
                pProducao.Range("vSobra").Copy Vsobra

        End If
        
        MsgBox "Sobras registradas com sucesso!" & vbNewLine _
        & "Agora na coluna MAR coloque as quantidades que v�o hoje para a mariland", vbInformation, "Sucesso!"
        pProducao.Range("K9").Select

End If

Planilha2.Range(Cells(22, 9), Cells(23, 10)).Interior.Pattern = xlNone
Planilha2.Range(Cells(24, 9), Cells(25, 10)).Interior.Color = 65535

btSobra.AutoShapeType = msoShapeRectangle
btSobra.ShapeStyle = msoShapeStylePreset8

btDiario.AutoShapeType = msoShapeRectangle
btDiario.ShapeStyle = msoShapeStylePreset9

btPrevisoes.AutoShapeType = msoShapeRectangle
btPrevisoes.ShapeStyle = msoShapeStylePreset8

btProducao.AutoShapeType = msoShapeRectangle
btProducao.ShapeStyle = msoShapeStylePreset8

btImprimir.AutoShapeType = msoShapeRectangle
btImprimir.ShapeStyle = msoShapeStylePreset8


End Sub

Sub diario()

Dim vDiario As Range
Dim vSM As Range
'Dim vPDA As Range

Set btSobra = Planilha2.Shapes.Range(Array("botaoSobra"))
Set btDiario = Planilha2.Shapes.Range(Array("botaoMariland"))
Set btPrevisoes = Planilha2.Shapes.Range(Array("botaoPrevisoes"))
Set btProducao = Planilha2.Shapes.Range(Array("botaoRegistros"))
Set btImprimir = Planilha2.Shapes.Range(Array("botaoImprimir"))

dataHoje = Day(Date)

m = Worksheets.Count
Set pControle = Worksheets(m)

n = pControle.Range("A15").CurrentRegion.Columns.Count + 1
i = pControle.Range("A26").CurrentRegion.Columns.Count + 1

Set vSM = Planilha2.Range("vSomar")
Set vDiario = Planilha2.Range("vDiario")
Set vProducao = Planilha2.Range("vProducao")

If pControle.Cells(2, n - 1).Value = dataHoje Then
        If MsgBox("Voc� deseja substuir os valores j� adicionados hoje?", vbYesNo, "Aten��o!") = vbYes Then
            Set VcontDiario = pControle.Cells(14, n - 1) 'Alterar aqui
                If n = i Then
                    pControle.Activate
                    Set vContProd = pControle.Range(Cells(24, i - 2), Cells(30, i - 2)) 'Alterar aqui
                    vContProd.Copy vProducao
                
                Else
                    pControle.Activate
                    Set vContProd = pControle.Range(Cells(24, i - 1), Cells(30, i - 1)) 'Alterar aqui
                    vContProd.Copy vProducao
            
                End If
            
            Planilha2.Activate
            vSM.Select
            Selection.Copy

            vDiario.Select
            Selection.PasteSpecial xlPasteValues
            
            vDiario.Copy VcontDiario
            Else: Exit Sub
        End If
   Else

    Set VcontDiario = pControle.Cells(14, n) 'Alterar aqui

    vSM.Select
    Selection.Copy

    vDiario.Select
    Selection.PasteSpecial xlPasteValues
        
    vDiario.Copy VcontDiario

End If




MsgBox "As quantidades de brownies para o dia est�o " _
& vbNewLine & "definidas, defina agora as previs�es ao lado", vbInformation, "Aten��o"

Planilha2.Range(Cells(24, 9), Cells(25, 10)).Interior.Pattern = xlNone
Planilha2.Range(Cells(26, 9), Cells(27, 10)).Interior.Color = 65535

btSobra.AutoShapeType = msoShapeRectangle
btSobra.ShapeStyle = msoShapeStylePreset8

btDiario.AutoShapeType = msoShapeRectangle
btDiario.ShapeStyle = msoShapeStylePreset8

btPrevisoes.AutoShapeType = msoShapeRectangle
btPrevisoes.ShapeStyle = msoShapeStylePreset9

btProducao.AutoShapeType = msoShapeRectangle
btProducao.ShapeStyle = msoShapeStylePreset8

btImprimir.AutoShapeType = msoShapeRectangle
btImprimir.ShapeStyle = msoShapeStylePreset8

Call mediaDiaria
Call saidaOntem

End Sub

Sub previsoes()

Set btSobra = Planilha2.Shapes.Range(Array("botaoSobra"))
Set btDiario = Planilha2.Shapes.Range(Array("botaoMariland"))
Set btPrevisoes = Planilha2.Shapes.Range(Array("botaoPrevisoes"))
Set btProducao = Planilha2.Shapes.Range(Array("botaoRegistros"))
Set btImprimir = Planilha2.Shapes.Range(Array("botaoImprimir"))


Dim vPrevisoes As Range
Dim vProducao As Range


Set vProducao = Planilha2.Range("producao")
Set vPrevisoes = Planilha2.Range("previsao")


vPrevisoes.Select
Selection.Copy
vProducao.Select
Selection.PasteSpecial xlPasteValues

Planilha2.Range(Cells(26, 9), Cells(27, 10)).Interior.Pattern = xlNone
Planilha2.Range(Cells(28, 9), Cells(29, 10)).Interior.Color = 65535


btSobra.AutoShapeType = msoShapeRectangle
btSobra.ShapeStyle = msoShapeStylePreset8

btDiario.AutoShapeType = msoShapeRectangle
btDiario.ShapeStyle = msoShapeStylePreset8

btPrevisoes.AutoShapeType = msoShapeRectangle
btPrevisoes.ShapeStyle = msoShapeStylePreset8

btProducao.AutoShapeType = msoShapeRectangle
btProducao.ShapeStyle = msoShapeStylePreset9

btImprimir.AutoShapeType = msoShapeRectangle
btImprimir.ShapeStyle = msoShapeStylePreset8

Planilha2.Range(Cells(8, 5), Cells(9, 5)).Select


End Sub

Sub producao()

Dim vMassas As Range
Dim vContMassas As Range
Dim vProducao As Range
Dim vContProd As Range
Dim nmassadas As Integer


Set btSobra = Planilha2.Shapes.Range(Array("botaoSobra"))
Set btDiario = Planilha2.Shapes.Range(Array("botaoMariland"))
Set btPrevisoes = Planilha2.Shapes.Range(Array("botaoPrevisoes"))
Set btProducao = Planilha2.Shapes.Range(Array("botaoRegistros"))
Set btImprimir = Planilha2.Shapes.Range(Array("botaoImprimir"))

nmassadas = Planilha2.Cells(8, 5)

dataHoje = Day(Date)

m = Worksheets.Count
Set pControle = Worksheets(m)

n = pControle.Range("A26").CurrentRegion.Columns.Count + 1

Set vMassas = Planilha2.Range("massas")
Set vProducao = Planilha2.Range("producao")

    If pControle.Cells(2, n - 1).Value = dataHoje Then
        If MsgBox("Voc� deseja substuir os valores j� adicionados hoje?", vbYesNo, "Aten��o!") = vbYes Then
            Set vContMassas = pControle.Cells(34, n - 1) 'Alterar aqui
            Set vContProd = pControle.Cells(24, n - 1) 'Alterar aqui
                If MsgBox("Essa produ��o vai dar " & nmassadas & " massadas. Confirma?", vbYesNo, "Aten��o!") = vbYes Then
                        vMassas.Copy vContMassas
                        vProducao.Copy vContProd
                    Else: Exit Sub
                End If
            Else: Exit Sub
        End If
    Else

        Set vContMassas = pControle.Cells(34, n) 'Alterar aqui
        Set vContProd = pControle.Cells(24, n) 'Alterar aqui
 
    If MsgBox("Essa produ��o vai dar " & nmassadas & " massadas. Confirma?", vbYesNo, "Aten��o!") = vbYes Then
                vMassas.Copy vContMassas
                vProducao.Copy vContProd
            Else: Exit Sub
     End If


End If
        
        
Planilha2.Range(Cells(28, 9), Cells(29, 10)).Interior.Pattern = xlNone
Planilha2.Range(Cells(30, 9), Cells(30, 10)).Interior.Color = 65535
        
        
btSobra.AutoShapeType = msoShapeRectangle
btSobra.ShapeStyle = msoShapeStylePreset8

btDiario.AutoShapeType = msoShapeRectangle
btDiario.ShapeStyle = msoShapeStylePreset8

btPrevisoes.AutoShapeType = msoShapeRectangle
btPrevisoes.ShapeStyle = msoShapeStylePreset8

btProducao.AutoShapeType = msoShapeRectangle
btProducao.ShapeStyle = msoShapeStylePreset8

btImprimir.AutoShapeType = msoShapeRectangle
btImprimir.ShapeStyle = msoShapeStylePreset9

Call mediaDiaria
        
End Sub

Sub imprimir()

Dim vMassas As Range


Set btSobra = Planilha2.Shapes.Range(Array("botaoSobra"))
Set btDiario = Planilha2.Shapes.Range(Array("botaoMariland"))
Set btPrevisoes = Planilha2.Shapes.Range(Array("botaoPrevisoes"))
Set btProducao = Planilha2.Shapes.Range(Array("botaoRegistros"))
Set btImprimir = Planilha2.Shapes.Range(Array("botaoImprimir"))


Set pProducao = Planilha2

If MsgBox("Proceder com a impress�o?", vbYesNo, "Aten��o!") = vbYes Then

If MsgBox("Deseja imprimir as quantidades de insumo?", vbYesNo, "Aten��o!") = vbYes Then Call impCont

Set vMassas = pProducao.Range("massas")

For i = 25 To 35 'Alterar aqui
    If pProducao.Rows(i).Columns(5).Value = 0 Then
        pProducao.Rows(i).Columns(5).Select
        Selection.EntireRow.Hidden = True
    End If
Next
    
pProducao.PrintOut , , , , "Atelie"

MsgBox "O documento foi impresso com sucesso!", vbInformation, "Sucesso!"
    
    
    Rows("25:35").Select 'Alterar aqui
    Selection.EntireRow.Hidden = False
   

vMassas.Value = ""
Range("A1").Select

btSobra.AutoShapeType = msoShapeRectangle
btSobra.ShapeStyle = msoShapeStylePreset9

btDiario.AutoShapeType = msoShapeRectangle
btDiario.ShapeStyle = msoShapeStylePreset8

btPrevisoes.AutoShapeType = msoShapeRectangle
btPrevisoes.ShapeStyle = msoShapeStylePreset8

btProducao.AutoShapeType = msoShapeRectangle
btProducao.ShapeStyle = msoShapeStylePreset8

btImprimir.AutoShapeType = msoShapeRectangle
btImprimir.ShapeStyle = msoShapeStylePreset8

Planilha2.Range(Cells(30, 9), Cells(30, 10)).Interior.Pattern = xlNone
Planilha2.Range(Cells(22, 9), Cells(23, 10)).Interior.Color = 65535


Call mediaDiaria

EstaPastaDeTrabalho.Save
    
Else: Exit Sub
End If

End Sub


Sub impCont()

Planilha1.PrintOut , , , , "Atelie"
'MsgBox "Os controles foram impressos com sucesso na impressora em cima do microondas!", vbInformation, "Sucesso!"
End Sub
Sub verificaData()


End Sub