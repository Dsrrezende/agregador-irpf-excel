Attribute VB_Name = "script_de_alinhamentos"
' ============================================================
' Macro de apoio: posiciona um ícone/forma em coordenadas fixas
' Útil para alinhar os botões do MENU do Agregador de IR.
' Como usar: Alt+F11 > Arquivo > Importar arquivo... > este .bas
' Salve a pasta de trabalho como .xlsm para manter a macro.
' ============================================================
Sub MoverIconeParaPosicao()
    Dim shp As Shape
    Dim ws As Worksheet
    Dim nomeIconeProcurado As String
    Dim novaPosicaoX As Double
    Dim novaPosicaoY As Double

    ' Planilha atual
    Set ws = ActiveSheet

    ' Nome exato do ícone, como aparece na Caixa de Nome do Excel
    nomeIconeProcurado = "Ícone 1" ' <-- troque pelo nome do seu ícone

    ' Posição desejada (em pontos)
    novaPosicaoX = 100
    novaPosicaoY = 50

    ' Procura o ícone na planilha
    For Each shp In ws.Shapes
        If shp.Name = nomeIconeProcurado Then
            shp.Left = novaPosicaoX
            shp.Top = novaPosicaoY
            MsgBox "Ícone '" & nomeIconeProcurado & "' movido com sucesso!", vbInformation
            Exit Sub
        End If
    Next shp

    ' Se não encontrar
    MsgBox "Ícone '" & nomeIconeProcurado & "' não encontrado.", vbExclamation
End Sub
