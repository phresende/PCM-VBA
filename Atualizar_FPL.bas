Attribute VB_Name = "Atualizar_FPL"
Sub Atualizar_FPL()

Caminho.Show
Datas2.Show

Pasta = Planilha1.Range("A2").Text
Pasta2 = Planilha1.Range("A12").Text
DataPlan = Planilha1.Range("A13").Text

    Workbooks.Open Filename:="G:\Rio Verde\Informações Gerais\PCM\IDM's\Fator de Planejamento.xlsm", writerespassword:="112029", ignorereadonlyrecommended:=True


'LIMPAR DADOS E COPIAR CPM


        Windows("Fator de Planejamento.xlsm").Activate
            Sheets("IW38").Visible = True
            Sheets("IW38").Select
                Range("A3").Select
                Range(Selection, Selection.End(xlToRight)).Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.ClearContents
    

            Sheets("Dados Formulas").Visible = True
            Sheets("Dados Formulas").Select
                Range("C2:C30").Select
                Selection.Copy


'SET SAP
        
        
        Set SapGuiAuto = GetObject("SAPGUI")          'Utiliza o objeto da interface gráfica do SAP
        Set SAPApp = SapGuiAuto.GetScriptingEngine    'Conecta ao SAP que está rodando no momento
        Set SAPCon = SAPApp.Children(0)               'Encontra o primeiro sistema que está conectado
        Set session = SAPCon.Children(0)              'Encontra a primeira sessão (janela) dessa conexão


'CODIGO SAP


session.findById("wnd[0]/tbar[0]/okcd").Text = "iw38"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkDY_MAB").Selected = True
session.findById("wnd[0]/usr/ctxtDATUV").Text = ""
session.findById("wnd[0]/usr/ctxtDATUB").Text = ""
session.findById("wnd[0]/usr/ctxtIWERK-LOW").Text = "331"
session.findById("wnd[0]/usr/btn%_ILART_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "op"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_INGPR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_PLGRP_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").Select
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").Text = "m*"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/btn%_STAE1_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "inat"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "mrel"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").SetFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtGLTRS-LOW").Text = Planilha1.Range("A8").Value
session.findById("wnd[0]/usr/ctxtGLTRS-HIGH").Text = Planilha1.Range("A9").Value
session.findById("wnd[0]/usr/ctxtVARIANT").Text = "FPL RVE 2023"
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[0]/menu[11]/menu[2]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "G:\Rio Verde\Gerencia Manutenção\Manutenção PCM\tmp 2\2022\" & Pasta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "IW38 - FPL.xls"
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 12
session.findById("wnd[0]").sendVKey 12


'ABRE PLANILHA EXPORTADA


    Workbooks.Open Filename:="G:\Rio Verde\Gerencia Manutenção\Manutenção PCM\tmp 2\2022\" & Pasta & "\IW38 - FPL.xls", ignorereadonlyrecommended:=True
    
        Windows("IW38 - FPL.xls").Activate


'TRATAR TABELA QUE ABRIU


            Rows("1:3").Select
            Selection.Delete Shift:=xlUp
            Rows("2:2").Select
            Selection.Delete Shift:=xlUp
            Columns("A:A").Select
            Selection.Delete Shift:=xlToLeft
            Range("A1").Select
            Range(Selection, Selection.End(xlToRight)).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy


'VOLTA PRA PLANILHA INICIAL E COLA


    Workbooks.Open Filename:="G:\Rio Verde\Gerencia Manutenção\Manutenção PCM\tmp 2\2022\" & Pasta & "\IW38 - FPL.xls", ignorereadonlyrecommended:=True

        Windows("Fator de Planejamento.xlsm").Activate
            Sheets("IW38").Select
                Range("A2").Select
                ActiveSheet.Paste
                Columns("D:D").Select
                Selection.NumberFormat = "0"
    
      
'COLAR FORMULA
        
VlinhaIW38 = Range("D3").End(xlDown).Row    'CONTAGEM DE LINHAS
        
        
                Range("O1:P1").Select
                Selection.Copy
                Range("O3").Select
                ActiveSheet.Paste
                Selection.AutoFill Destination:=Range("O3:P" & VlinhaIW38)
                
                Range("A2:P" & VlinhaIW38).Select
                Selection.Copy


'COLAR IW38 NO CONTROLE FLP

    Workbooks.Open Filename:="G:\Rio Verde\Gerencia Manutenção\Manutenção PCM\tmp 2\2022\" & Pasta2 & "\Controle FPL.xlsx", ignorereadonlyrecommended:=True

        Windows("Controle FPL.xlsx").Activate
 
        For C = 1 To Sheets.Count
        
            If Sheets(C).Name = DataPlan Then
            
                Sheets(DataPlan).Select
                    Range("A1").Select
                    Range(Selection, Selection.End(xlToRight)).Select
                    Range(Selection, Selection.End(xlDown)).Select
                    Selection.ClearContents
                
            Windows("Fator de Planejamento.xlsm").Activate
                Sheets("IW38").Select
                    Range("A2:P" & VlinhaIW38).Select
                    Selection.Copy
                    
            Windows("Controle FPL.xlsx").Activate
                Sheets(DataPlan).Select
                    Range("A1").Select
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
         :=False, Transpose:=False
    
                Exit For
                
            End If
               
            If C = Sheets.Count Then
       
                Sheets.Add After:=ActiveSheet
                    PlanilhaN2 = ActiveSheet.Name
                    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
                Sheets(PlanilhaN2).Select
                Sheets(PlanilhaN2).Name = DataPlan
                                
                Exit For

            End If
        
        Next C
       
       
    ActiveWorkbook.Save



'FECHAR PLANILHA QUE FOI ABERTA
    
    
        Workbooks("IW38 - FPL.XLS").Close SaveChanges:=False
        Workbooks("Controle FPL.XLSX").Close SaveChanges:=False

'VOLTA PARA RESUMO


    Windows("Fator de Planejamento.xlsm").Activate
        Sheets("Dados Formulas").Select
            Range("H2").Select
            Selection.Copy
        Sheets("Resumo").Select
            Range("H33").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False

        Sheets("IW38").Visible = False
        Sheets("Dados Formulas").Visible = False
        Sheets("Resumo").Select
    
    ActiveWorkbook.RefreshAll
    ActiveWorkbook.Save
  
MsgBox ("O Fator de Planejamento foi ATUALIZADO!")


End Sub
