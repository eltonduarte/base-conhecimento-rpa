' Retorna o primeiro dia do mês anterior
Function DataInicial()
  DataInicial = DateSerial(Year(Date), Month(Date) - 1, 1)
  DataInicial = Replace(DataInicial, "/", ".")
End function

' Retorna o último dia do mês anterior
Function DataFinal()
  DataFinal = DateSerial(Year(Date), Month(Date), 0)
  DataFinal = Replace(DataFinal, "/", ".")
End function


## Entrar comm mais de 1 parametro no vbs
# Precisa atribuir os argumentos separados por espaço numa variavel
# Exemplo: $iStrTransacaoFb70$ $iStrEmpresa$

transacao = Wscript.Arguments.Item(0)
empresa = Wscript.Arguments.Item(1)

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

'Digitando a transacao e ENTER
session.findById("wnd[0]/tbar[0]/okcd").text = transacao
session.findById("wnd[0]").sendVKey 0

'Checando a janela "Entrar Empresa", Caso exista preenche o campo e ENTER
If Not session.FindById("wnd[1]/usr/ctxtBKPF-BUKRS",False) Is Nothing Then
  	session.FindById("wnd[1]/usr/ctxtBKPF-BUKRS").text = empresa
	session.findById("wnd[0]").sendVKey 0
End If

'Checando Campo "Tp.Doc." Caso não exista efetua a configuração em "Opcões de processamento"
If session.FindById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPLFDCB:0510/cmbINVFO-BLART",False) Is Nothing Then
	session.findById("wnd[0]/tbar[1]/btn[16]").press
	session.findById("wnd[0]/usr/tabsTS/tabp1100/ssubS1100:SAPMF05O:1100/cmbRFOPTE-DMTTP").key = "4"
	session.findById("wnd[0]/tbar[0]/btn[11]").press
  	session.findById("wnd[0]/tbar[0]/okcd").text = transacao
	session.findById("wnd[0]").sendVKey 0
End If

'Digita a Empresa dentro da transacao para sempre atualizar o numero da empresa
session.findById("wnd[0]").sendVKey 7
session.findById("wnd[1]/usr/ctxtBKPF-BUKRS").text = empresa
session.findById("wnd[1]").sendVKey 0

'Digitando a transacao novamente para atualizar a tela (campos bloqueiam após troca de empresa)
session.findById("wnd[0]/tbar[0]/okcd").text = transacao
session.findById("wnd[0]").sendVKey 0
