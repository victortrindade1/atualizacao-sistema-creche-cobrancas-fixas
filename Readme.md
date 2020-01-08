# Atualização: Inserir cobranças extras

## Banco Cadastro

Inseridos os campos no banco:

	- Refeições
	- Horas Extras
	- Outros

## Navegação

### Public Function GetTabOrder() As Variant

```diff
Public Function GetTabOrder() As Variant
'  Ver 2 2014 - Dave Timms (aka DMT32) and  Jerry Sullivan
'--set the tab order of input cells - change ranges as required
'  do not use "$" in the cell addresses.

   Select Case ActiveSheet.Name
      Case "CADASTRO"
-         GetTabOrder = Array("H10", "R10", "H20", "H22", "H24", "H26", "H28", "H30", "AA20", "AA22", "AA24", _
-            "AA26", "AT20", "AT22", "AT24", "AT26", "AT28", "AT30", "H43", "H45", "H47", _
-            "H49", "H51", "H53", "Y43", "Y45", "Y47", "AJ43", "AJ45", "AJ47", "AX43", "AX45", "AX47", _
-            "H62", "H64", "Y62", "Y64", "Y66", "Y68", "AH62", "AH64", "AH66", "AT62", "H83", "H85", _
-            "Y83", "Y89", "AT83", "AT85", "H106", "H108", "H110", "Y104", "Y108", "AJ104", "H127", "Y127", "AM127", "H135", "H137")
+        GetTabOrder = Array( _
+                "H10", "R10", "H20", "H22", "H24", "H26", "H28", "H30", "AA20", "AA22", "AA24", "AA26", "AT20", "AT22", "AT24", "AT26", "AT28", "AT30", _
+                "H43", "H45", "H47", "H49", "H51", "H53", "Y43", "Y45", "Y47", "AJ43", "AJ45", "AJ47", "AX43", "AX45", "AX47", _
+                "H62", "H64", "Y62", "Y64", "Y66", "Y68", "AH62", "AH64", "AH66", "AT62", _
+                "H83", "H85", "Y83", "Y89", "AT83", "AT85", _
+                "H106", "H108", "H110", "H112", "H114", "H116", _
+                "Y104", "Y108", "AJ104", "H127", "Y127", "AM127", "H135", "H137" _
+        )
'      Case "Sheet2", "Sheet3"
'         GetTabOrder = Array("D8", "F8", "L6", "H8", "J5", "I10", "L8", "D12")
'      Case "Sheet6"
'         GetTabOrder = Array("D18", "F18", "E19", "H18", "L16", "D22")
'      Case Else
'         MsgBox "Error: Tab Order has not been specified for this sheet."
   End Select
End Function
```

## Aba Cadastro

Inseridos 3 novos campos de Cobranças Fixas.

### Sub novoAluno()

```diff
    ' Insere dados do Cadastro
    Dim i As Integer
    
-    For i = 0 To 54
+    For i = 0 To 57
        tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Matricula").Index).Offset(0, i) = planCad.Range("Cad_" & i).Value
    Next

    ...

    ' Validação: Dia do Vencimento
    ' Se estiver em branco, vencimento = dia 10
-    If planCad.Range("Cad_52").Value = "" Then
+    If planCad.Range("Cad_55").Value = "" Then
-        planCad.Range("Cad_52").Value = 10
+        planCad.Range("Cad_55").Value = 10
    End If
```

### Sub EditaAluno()

```diff
    ' Insere dados Editados
    Dim i As Integer
    
-    For i = 0 To 54
+    For i = 0 To 57
        planDCad.Cells(rowFiltrado, tbDCad.ListColumns("Matricula").Index + 1).Offset(0, i) = planCad.Range("Cad_" & i).Value
    Next
```

### Sub BuscarAluno(colTbDCad As Integer, criterioFiltro As Variant)

```diff
    If visibleRows > 0 Then
    
        Dim rowFiltrado As Integer
    
        rowFiltrado = tbDCad.DataBodyRange.Columns.SpecialCells(xlCellTypeVisible).row
        
    
        ' Zera valores
        Dim a As Integer
-        For a = 0 To 55
+        For a = 0 To 58
            planCad.Range("Cad_" & a).Value = ""
        Next


        ' Coloca valores de Dados Cadastro em Cadastro
-        For a = 0 To 55
+        For a = 0 To 58
            planCad.Range("Cad_" & a) = planDCad.Cells(rowFiltrado, tbDCad.ListColumns("Matricula").Index + 1).Offset(0, a)
        Next
```

### Sub AbrePlanInserirPagamento()

```diff
- ActiveSheet.Range("InsPag_Status").Value = Sheets("CADASTRO").Range("Cad_55").Value ' Status
+ ActiveSheet.Range("InsPag_Status").Value = Sheets("CADASTRO").Range("Cad_58").Value ' Status
```

### Sub BuscarNome()

```diff
    ' Zera valores
    Dim a As Integer
-    For a = 0 To 55
+    For a = 0 To 58
        planCad.Range("Cad_" & a).Value = ""
    Next
```

### Sub EscolheNomeListBox()

```diff
    ' Coloca valores de Dados Cadastro em Cadastro
-    For a = 0 To 55
+    For a = 0 To 58
        planCad.Range("Cad_" & a) = planDCad.Cells(rowNomeEscolhido, tbDCad.ListColumns("Matricula").Index + 1).Offset(0, a)
    Next
```

### Sub LimpaTelaCadastro()

```diff
    ' Limpa dados pessoais, dados escolares, cobranças fixas mensais
    Dim a As Integer
-    For a = 0 To 55
+    For a = 0 To 58
        Sheets("CADASTRO").Range("Cad_" & a) = ""
    Next
```

### Sub LimpaTelaCadastroTelaEditar()

```diff
    ' Limpa dados pessoais, dados escolares, cobranças fixas mensais
    Dim a As Integer
-    For a = 0 To 55
+    For a = 0 To 58
        Sheets("CADASTRO").Range("Cad_" & a) = ""
    Next
```

## Aba Inserir Pagamento

### Sub ReceberPagamento()

```diff
- planCad.Range("Cad_55").Value = statusPrinc
+ planCad.Range("Cad_58").Value = statusPrinc
```

### Sub ListaCobrancas(ByVal rowTbDCad As Integer)

```diff
    ' Cobranças Mensais:

    ' Mensalidade
    If tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Mensalidade").Index) <> "" Then
        planInsPag.Cells(rowCelCobr, colCelCobr) = "Mensalidade"
        planInsPag.Cells(rowCelCobr, colCelCobrVal + 1) = tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Mensalidade").Index).Value
        rowCelCobr = rowCelCobr + 1
    End If
    
    ' Judô
    If tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Judo").Index) <> "" Then
        planInsPag.Cells(rowCelCobr, colCelCobr) = "Judo"
        planInsPag.Cells(rowCelCobr, colCelCobrVal + 1) = tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Judo").Index).Value
        rowCelCobr = rowCelCobr + 1
    End If
    
    ' Balé
    If tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Bale").Index) <> "" Then
        planInsPag.Cells(rowCelCobr, colCelCobr) = "Bale"
        planInsPag.Cells(rowCelCobr, colCelCobrVal + 1) = tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Bale").Index).Value
        rowCelCobr = rowCelCobr + 1
    End If
  
+    ' Refeições
+    If tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Refeicoes").Index) <> "" Then
+        planInsPag.Cells(rowCelCobr, colCelCobr) = "Refeicoes"
+        planInsPag.Cells(rowCelCobr, colCelCobrVal + 1) = tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Refeicoes").Index).Value
+        rowCelCobr = rowCelCobr + 1
+    End If
+    
+    ' Horas Extras
+    If tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Horas Extras").Index) <> "" Then
+        planInsPag.Cells(rowCelCobr, colCelCobr) = "Horas Extras"
+        planInsPag.Cells(rowCelCobr, colCelCobrVal + 1) = tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Horas Extras").Index).Value
+        rowCelCobr = rowCelCobr + 1
+    End If
+    
+    ' Outros
+    If tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Outros").Index) <> "" Then
+        planInsPag.Cells(rowCelCobr, colCelCobr) = "Outros"
+        planInsPag.Cells(rowCelCobr, colCelCobrVal + 1) = tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Outros").Index).Value
+        rowCelCobr = rowCelCobr + 1
+    End If


    ' Cobranças a receber de outros meses
    Dim matricula As Long
    matricula = planInsPag.Range("InsPag_0")
```

## Aba Consultar Turmas

Inseridas as colunas `Refeições`, `Horas Extras` e `Outros` na tabela `TabelaConsultarTurmas`

### Sub SalvarAlteracoes()

```diff
        planDCad.Cells(cel.row, tbDCad.ListColumns("Judo").Index + 1).Value = _
            tbCTur.DataBodyRange.Cells(rowIn, tbCTur.ListColumns("Judô").Index).Value
            
        planDCad.Cells(cel.row, tbDCad.ListColumns("Bale").Index + 1).Value = _
            tbCTur.DataBodyRange.Cells(rowIn, tbCTur.ListColumns("Balé").Index).Value
        
+        planDCad.Cells(cel.row, tbDCad.ListColumns("Refeicoes").Index + 1).Value = _
+            tbCTur.DataBodyRange.Cells(rowIn, tbCTur.ListColumns("Refeições").Index).Value
+        
+        planDCad.Cells(cel.row, tbDCad.ListColumns("Horas Extras").Index + 1).Value = _
+            tbCTur.DataBodyRange.Cells(rowIn, tbCTur.ListColumns("Horas Extras").Index).Value
+        
+        planDCad.Cells(cel.row, tbDCad.ListColumns("Outros").Index + 1).Value = _
+            tbCTur.DataBodyRange.Cells(rowIn, tbCTur.ListColumns("Outros").Index).Value

        planDCad.Cells(cel.row, tbDCad.ListColumns("Vencimento").Index + 1).Value = _
            tbCTur.DataBodyRange.Cells(rowIn, tbCTur.ListColumns("Venc.").Index).Value

        planDCad.Cells(cel.row, tbDCad.ListColumns("Desconto Fixo").Index + 1).Value = _
            tbCTur.DataBodyRange.Cells(rowIn, tbCTur.ListColumns("Desc.").Index).Value
            
        rowIn = rowIn + 1
        
    Next
```

## Aba Financeiro

### Sub FinReceitasAReceber()

```diff
    ' TabMens:

    ' Copia colunas Mensalidade, Judo, Bale, desconto fixo
    tbDCad.ListColumns("Mensalidade").DataBodyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=tbMens.ListColumns("Mensalidade").DataBodyRange
    tbDCad.ListColumns("Judo").DataBodyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=tbMens.ListColumns("Judo").DataBodyRange
    tbDCad.ListColumns("Bale").DataBodyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=tbMens.ListColumns("Bale").DataBodyRange
+    tbDCad.ListColumns("Refeicoes").DataBodyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=tbMens.ListColumns("Refeicoes").DataBodyRange
+    tbDCad.ListColumns("Horas Extras").DataBodyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=tbMens.ListColumns("Horas Extras").DataBodyRange
+    tbDCad.ListColumns("Outros").DataBodyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=tbMens.ListColumns("Outros").DataBodyRange
    tbDCad.ListColumns("Desconto Fixo").DataBodyRange.SpecialCells(xlCellTypeVisible).Copy Destination:=tbMens.ListColumns("Desc Fixo").DataBodyRange
```

## Banco Financeiro

Inseridas colunas `Refeicoes`, `Horas Extras`, `Outros` em `TabMens`.

# Atualização: Inserir Descrição em Inserir Nova Receita (Painel Financeiro)

