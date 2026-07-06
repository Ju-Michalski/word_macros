Attribute VB_Name = "NewMacros"
Sub WValorExtenso()

    On Error GoTo Erro

    Selection.MoveStartUntil cset:=" ", Count:=wdBackward
    Selection.TypeText FormatarMoedaSemArredondar(Selection.Text) & " (" & ConverterParaExtenso(Selection.Text) & ")"

    GoTo Pula

Erro:

    MsgBox "O valor deve ser informado sem ponto e sem 'R$'." & Chr$(10) & "O cursor deve estar imediatamente após o valor." _
    & Chr$(10) & "O valor não pode estar em início de parágrafo." & Chr$(10) & _
    "Exemplo: 1250,35  ou  0,00035 (até 5 casas decimais)", vbCritical, "Dados inválidos!"

    Exit Sub

Pula:

End Sub

'Formata o valor com separador de milhar, sem arredondar os decimais (mostra exatamente as casas digitadas)
Private Function FormatarMoedaSemArredondar(sValor As String) As String
Dim sInteiro As String, sDecimal As String
Dim sFormatado As String
Dim i As Integer, iCont As Integer

If InStr(1, sValor, ",") > 0 Then
    sInteiro = Mid(sValor, 1, InStr(1, sValor, ",") - 1)
    sDecimal = Right(sValor, Len(sValor) - InStr(1, sValor, ","))
    If Len(sDecimal) > 5 Then sDecimal = Left(sDecimal, 5) 'trunca sem arredondar
Else
    sInteiro = sValor
    sDecimal = ""
End If

If sInteiro = "" Then sInteiro = "0"

iCont = 0
sFormatado = ""
For i = Len(sInteiro) To 1 Step -1
    sFormatado = Mid(sInteiro, i, 1) & sFormatado
    iCont = iCont + 1
    If iCont Mod 3 = 0 And i > 1 Then sFormatado = "." & sFormatado
Next i

FormatarMoedaSemArredondar = "R$ " & sFormatado & IIf(sDecimal <> "", "," & sDecimal, "")

End Function

Public Function ConverterParaExtenso(NumeroParaConverter As String) As String
Dim sExtensoFinal As String
Dim sDecimais As String
Dim sMoedaSing As String, sMoedaPlu As String, sDecimalExtenso As String
Dim bSufMoeda As Boolean

'Separa os Decimais (aceita de 1 a 5 casas, sem arredondar)
If InStr(1, NumeroParaConverter, ",") > 0 Then
    sDecimais = Right(NumeroParaConverter, Len(NumeroParaConverter) - InStr(1, NumeroParaConverter, ","))
    NumeroParaConverter = Mid(NumeroParaConverter, 1, InStr(1, NumeroParaConverter, ",") - 1)
    If Len(sDecimais) > 5 Then sDecimais = Left(sDecimais, 5) 'trunca sem arredondar
End If

'Escreve a parte inteira por extenso (bSufMoeda indica se usa "de reais")
sExtensoFinal = NumeroExtenso(NumeroParaConverter, bSufMoeda)

'Define a moeda
sMoedaPlu = " reais"
sMoedaSing = " real"
If bSufMoeda = True And sExtensoFinal <> "" Then sMoedaPlu = " de reais"

'Escreve a parte decimal, com a denominação correta (décimos, centavos, milésimos, etc.)
sDecimalExtenso = EscreveDecimaisMonetarios(sDecimais)

'Adiciona a moeda e os decimais
sExtensoFinal = IIf((sExtensoFinal = ""), "", sExtensoFinal & IIf((sExtensoFinal = "um"), sMoedaSing, sMoedaPlu)) _
& IIf((sExtensoFinal = ""), sDecimalExtenso, IIf((sDecimalExtenso = ""), "", " e " & sDecimalExtenso))

'retorna o resultado

sExtensoFinal = Replace(sExtensoFinal, "  ", " ", 1, , vbTextCompare)

ConverterParaExtenso = Replace(sExtensoFinal, "e e ", "e ", 1, , vbTextCompare)

End Function

'Escreve qualquer número inteiro (de 1 a muitos dígitos) por extenso, sem sufixo de moeda.
'bZeraSufixoGrande sai True se os 2 últimos grupos (unidade e milhar) forem vazios,
'usado para decidir se usa "de reais" (ex: 1.000.000 de reais x 1.000.001 reais)
Private Function NumeroExtenso(ByVal sValor As String, Optional ByRef bZeraSufixoGrande As Boolean) As String
Dim sExtensoFinal As String, sExtensoAtual As String
Dim iQtdGrupos As Integer
Dim sConector As String
Dim vArrCenten As Variant
Dim i As Integer, w As Integer

bZeraSufixoGrande = False

If sValor = "" Then
    NumeroExtenso = ""
    Exit Function
End If

iQtdGrupos = Fix(Len(sValor) / 3)
If Len(sValor) Mod 3 > 0 Then iQtdGrupos = iQtdGrupos + 1

If iQtdGrupos > 2 Then bZeraSufixoGrande = True

For i = iQtdGrupos To 1 Step -1
    sExtensoAtual = DesmembraValor(sValor, i)
    If i = 1 Then
        If sExtensoAtual = "" Then
            sExtensoFinal = sExtensoFinal & sExtensoAtual
        Else
            If sExtensoFinal = "" Then
                sExtensoFinal = sExtensoFinal & sExtensoAtual
            Else

                vArrCenten = Array("cem", "duzentos", "trezentos", "quatrocentos", _
                "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos")

                sConector = ""

                For w = 0 To 8
                    If Len(sValor) >= 4 And Right(sValor, 2) = "00" _
                    And sExtensoAtual <> vArrCenten(w) Then sConector = "e "
                    Exit For
                Next w

                If Len(sValor) >= 4 And Left(Right(sValor, 3), 1) = "0" Then sConector = " e "

                If Len(sValor) >= 4 And sExtensoAtual = "cem" Then sConector = " e "

                sExtensoFinal = sExtensoFinal & sConector & sExtensoAtual
            End If
        End If
    Else
        sExtensoFinal = sExtensoFinal & sExtensoAtual
    End If

    If iQtdGrupos > 2 Then
        Select Case i
        Case 1, 2
            If sExtensoAtual <> "" Then
                bZeraSufixoGrande = False
            End If
        End Select
    End If
Next i

NumeroExtenso = sExtensoFinal

End Function

Private Function DesmembraValor(sValor As String, iGrupoDiv As Integer) As String
Dim iValor As Integer
Dim sExtenso As String
Dim iDivResto As Integer
Dim iDivInteiro As Integer
Dim iPosInicMid As Integer
Dim iTamMid As Integer
Dim sComplemento As String
Dim vArrDez1 As Variant
Dim vArrDez2 As Variant
Dim vArrCentena As Variant

vArrDez1 = Array("", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove", _
"dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", _
"dezoito", "dezenove")

vArrDez2 = Array("vinte", "trinta", "quarenta", "cinquenta", "sessenta", _
"setenta", "oitenta", "noventa")

vArrCentena = Array("cem", "cento", "duzentos", "trezentos", "quatrocentos", _
"quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos")

'Pega o Valor a ser escrito e desmembra para o grupo numérico correto
iPosInicMid = Len(sValor) - ((3 * iGrupoDiv) - 1)
If iPosInicMid <= 1 Then
iTamMid = 2 + iPosInicMid
Else
iTamMid = 3
End If

If iPosInicMid < 1 Then iPosInicMid = 1

iValor = CInt(Mid(sValor, iPosInicMid, iTamMid))

Select Case iGrupoDiv
Case 2
sComplemento = " mil "
Case 3
If iValor = 1 Then
sComplemento = " milhão "
Else
sComplemento = " milhões "
End If
Case 4
If iValor = 1 Then
sComplemento = " bilhão "
Else
sComplemento = " bilhões "
End If
Case 5
If iValor = 1 Then
sComplemento = " trilhão "
Else
sComplemento = " trilhões "
End If
End Select

Select Case iValor
Case 0 To 19
sExtenso = vArrDez1(iValor)
Case 20 To 99
iDivInteiro = Fix(iValor / 10)
iDivResto = iValor Mod 10

If iDivResto = 0 Then
sExtenso = vArrDez2(iDivInteiro - 2)
Else
sExtenso = vArrDez2(iDivInteiro - 2) & " e " & vArrDez1(iDivResto)
End If
Case 100 To 999
iDivInteiro = Fix(iValor / 100)
iDivResto = iValor Mod 100

If iDivResto = 0 Then
If iDivInteiro = 1 Then
sExtenso = vArrCentena(0)   'Cem
Else
sExtenso = vArrCentena(iDivInteiro) 'inteiro maior que 100
End If
Else
sExtenso = vArrCentena(iDivInteiro) & " e "
Select Case iDivResto
Case 0 To 19
sExtenso = sExtenso & vArrDez1(iDivResto)
Case 20 To 99
iDivInteiro2 = Fix(iDivResto / 10)
iDivResto2 = iDivResto Mod 10

If iDivResto2 = 0 Then
sExtenso = sExtenso & vArrDez2(iDivInteiro2 - 2)
Else
sExtenso = sExtenso & vArrDez2(iDivInteiro2 - 2) & " e " & vArrDez1(iDivResto2)
End If
End Select
End If

End Select

If sExtenso = "um" And sComplemento = " mil " And Len(sValor) < 7 Then
sComplemento = "mil "
sExtenso = ""
End If

smilx = Right(sValor, 6)

If sComplemento = " milhão " Then
If Left(smilx, 2) = "00" And Right(smilx, 5) <> "00000" Then sComplemento = " milhão e " Else sComplemento = " milhão "
End If

If sComplemento = " milhões " Then
If Right(smilx, 6) = "000000" Then
sComplemento = " milhões "
Else
If Left(smilx, 2) = "00" And Right(smilx, 5) <> "00000" Then sComplemento = " milhões e " Else sComplemento = " milhões "
End If
End If

DesmembraValor = sExtenso & IIf(iValor > 0, sComplemento, "")

End Function

'Escreve a parte decimal (fração do real), sem preencher com zeros, usando a denominação
'correta conforme a quantidade de casas decimais digitadas
Private Function EscreveDecimaisMonetarios(sCent As String) As String
Dim sExtenso As String
Dim sSufixo As String
Dim iCent As Long
Dim iCasas As Integer

If sCent = "" Then
    EscreveDecimaisMonetarios = ""
    Exit Function
End If

iCasas = Len(sCent)
iCent = CLng(sCent)

Select Case iCasas
Case 1
    sSufixo = IIf(iCent = 1, " décimo", " décimos")
Case 2
    sSufixo = IIf(iCent = 1, " centavo", " centavos")
Case 3
    sSufixo = IIf(iCent = 1, " milésimo", " milésimos")
Case 4
    sSufixo = IIf(iCent = 1, " décimo de milésimo", " décimos de milésimo")
Case 5
    sSufixo = IIf(iCent = 1, " centésimo de milésimo", " centésimos de milésimo")
End Select

sExtenso = NumeroExtenso(CStr(iCent))

EscreveDecimaisMonetarios = IIf(iCent > 0, sExtenso & sSufixo, "")

End Function

