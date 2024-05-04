Attribute VB_Name = "Módulo1"
Sub compilar_carros()

Dim resposta1 As String
Dim ult_lin, resposta As Integer
Dim unidades As range

Set unidades = Sheets("Concessionárias").range("A2:A9")
'Perguntar ao usuario se ele quer rodar a macro
resposta = MsgBox("Deseja Iniciar a macro", vbYesNo + vbInformation, "Compilar")

If resposta = 6 Then

    'Perguntar se vai ser carros novos ou usados
    resposta1 = UCase(InputBox("Qual a situação do carro?", "Situação", "Novo ou Usado"))
    
    For Each unidade In unidades
    
        'Filtrar a base de acordo
        If resposta1 = "NOVO" Then
            
            'preenche o filtro
            ult_lin = range("A1").End(xlDown).Row
            range("A1:F" & ult_lin).AutoFilter Field:=6, Criteria1:="Novo"
            range("A1:F" & ult_lin).AutoFilter Field:=1, Criteria1:=unidade
            
            'pega a ultima linha para copiar
            ult_lin = range("A1").End(xlDown).Row
            range("A1:F" & ult_lin).Copy
            
            'Tratar o nome para abrir a aba
            uni = Mid(unidade, 7)
            situacao = " - Novos"
            Sheets(uni & situacao).Activate
            range("A1").PasteSpecial
            
            'Retira os filtros
            ActiveSheet.range("$A$1:$F$1600").AutoFilter Field:=1
            ActiveSheet.range("$A$1:$F$1600").AutoFilter Field:=6
            
            'Arruma as colunas
            Cells.EntireColumn.AutoFit
            
        ElseIf resposta1 = "USADO" Then
        
            'preenche o filtro
            ult_lin = range("A1").End(xlDown).Row
            range("A1:F" & ult_lin).AutoFilter Field:=6, Criteria1:="Usado"
            range("A1:F" & ult_lin).AutoFilter Field:=1, Criteria1:=unidade
            
            'pega a ultima linha para copiar
            ult_lin = range("A1").End(xlDown).Row
            range("A1:F" & ult_lin).Copy
            
            'Tratar o nome para abrir a aba
            uni = Mid(unidade, 7)
            situacao = " - Usados"
            Sheets(uni & situacao).Activate
            range("A1").PasteSpecial
            
            'Retira os filtros
            ActiveSheet.range("$A$1:$F$1600").AutoFilter Field:=1
            ActiveSheet.range("$A$1:$F$1600").AutoFilter Field:=6
            
            'Arruma as colunas
            Cells.EntireColumn.AutoFit
        
        
        End If
    
    Sheets("Resumo").Activate
    ActiveSheet.range("$A$1:$F$1600").AutoFilter Field:=1
    ActiveSheet.range("$A$1:$F$1600").AutoFilter Field:=6
    
    Next unidade
End If

MsgBox ("Fim da macro")
End Sub


