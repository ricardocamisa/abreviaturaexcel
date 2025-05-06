Option Explicit

Function AbreviarNomeLongos(nomeCompleto As String, Optional limiteDeCaracteres As Long = 26) As String
    Dim partes() As String
    Dim nomeAbreviado As String
    Dim i As Integer
    Dim comprimentoTotal As Integer
    Dim palavrasIgnoradas As Object
    Dim preservarProximo As Boolean


    Set palavrasIgnoradas = CreateObject("Scripting.Dictionary")
    palavrasIgnoradas.Add "de", True
    palavrasIgnoradas.Add "da", True
    palavrasIgnoradas.Add "do", True
    palavrasIgnoradas.Add "das", True
    palavrasIgnoradas.Add "dos", True
    
    comprimentoTotal = Len(nomeCompleto)
    If comprimentoTotal <= limiteDeCaracteres Then
        AbreviarNomeLongos = nomeCompleto
        Exit Function
    End If
    partes = Split(nomeCompleto, " ")
    If UBound(partes) = 0 Then
        AbreviarNomeLongos = nomeCompleto
        Exit Function
    End If

    nomeAbreviado = partes(0) & " "
    preservarProximo = False
    For i = 1 To UBound(partes) - 1
        If partes(i) <> "" Then
            If preservarProximo Then
                nomeAbreviado = nomeAbreviado & partes(i) & " "
                preservarProximo = False
            ElseIf palavrasIgnoradas.Exists(LCase(partes(i))) Then
                nomeAbreviado = nomeAbreviado & partes(i) & " "
                preservarProximo = True
            Else
                nomeAbreviado = nomeAbreviado & Left(partes(i), 1) & ". "
            End If
        End If
    Next i
    nomeAbreviado = nomeAbreviado & partes(UBound(partes))
    AbreviarNomeLongos = nomeAbreviado
End Function
