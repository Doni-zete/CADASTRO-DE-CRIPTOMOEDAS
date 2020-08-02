VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CRIPTOMOEDAS 
   Caption         =   "CRIPTOMOEDAS"
   ClientHeight    =   8100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12345
   OleObjectBlob   =   "CRIPTOMOEDAS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CRIPTOMOEDAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btfechar_Click()
    Unload Me
End Sub


Private Sub btsalvar_Click()
Dim i As Integer

i = Range("A1").CurrentRegion.Rows.Count + 1
    Cells(i, 1).Value = txsigla.Value
    Cells(i, 2).Value = txtmoeda.Value
    Cells(i, 3).Value = cbtipo.Value
    Cells(i, 4).Value = cbexchange.Value
    Cells(i, 5).Value = txtquantidade.Value
Call Atualizar
MsgBox "CRIPTOMOEDA CADASTRADA COM SUCESSO", vbInformation, "Informação"
End Sub

Private Sub cbAdcionar_Click()

    txsigla = ""
    txtmoeda = ""
    cbtipo = ""
    cbexchange = ""
    txtquantidade = ""
End Sub


Private Sub cbEditar_Click()




End Sub

Private Sub cbLimpar_Click()
Dim resp As Integer

    With Worksheets("Planilha1").Range("A:A")
    
    Set c = .Find(txsigla.Value, LookIn:=xlValues, Lookat:=xlPart)
    If Not c Is Nothing Then
    resp = MsgBox("CONFIRMAR EXCLUSÃO DE CRIPTOMOEDA?", vbYesNo, "CONFIRMAR")
    
        If resp = vbYes Then
        c.Select
        Selection.EntireRow.Delete
    Else
    MsgBox ("CRIPTOMOEDA NÃO EXCLUIDA!")
    End If
    Else
    MsgBox ("CRIPTOMOEDA NÃO ENCONTRADA!!!")
    End If
    End With
End Sub
Private Sub cbPesquisar_Click()

txsigla.Enabled = True

    With Worksheets("Planilha1").Range("A:A")
    Set c = .Find(txsigla.Value, LookIn:=xlValues, Lookat:=xlPart)
    If Not c Is Nothing Then

    c.Activate
    txsigla.Value = c.Value
    txtmoeda.Value = c.Offset(0, 1).Value
    cbtipo.Value = c.Offset(0, 2).Value
    cbexchange.Value = c.Offset(0, 3).Value
    txtquantidade.Value = c.Offset(0, 4).Value
    Else

    MsgBox "NENHUMA MOEDA REFERENTE A PESQUISA FOI ENCONTRADA! ", vbInformation, "AVISO"

    End If
    End With

    txtmoeda.Enabled = False
    cbtipo.Enabled = False
    cbexchange.Enabled = False
    txtquantidade.Enabled = False

End Sub
Private Sub ListBox1_Click()
Dim n As Integer

    n = ListBox1.ListIndex + 2
    txsigla.Value = Cells(n, 1).Value
    txtmoeda.Value = Cells(n, 2).Value
    cbtipo.Value = Cells(n, 3).Value
    cbexchange.Value = Cells(n, 4).Value
    txtquantidade.Value = Cells(n, 5).Value

End Sub
Private Sub SpinButton1_Change()
    txtquantidade = SpinButton1.Value
End Sub
Private Sub UserForm_Initialize()
With cbtipo
    .AddItem "MOEDA"
    .AddItem "TOKEN"
End With

With cbexchange
    .AddItem "BRAZILIEX"
    .AddItem "BINANCE"
    .AddItem "BITZ"
    .AddItem "CREX24"
    .AddItem "KUCOIN"
End With
Call Atualizar
End Sub
Sub Atualizar()
    Dim n As Integer, b As Range
    n = Range("A1").CurrentRegion.Rows.Count
    
    Set b = Range(Cells(2, 1), Cells(n, 5))
    ListBox1.RowSource = b.Address

End Sub
