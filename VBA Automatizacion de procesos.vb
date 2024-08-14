CORREO DE CORTES

Sub EnviarCorreoBaseADHsql()

Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")
Dim olMail As Outlook.MailItem
Set olMail = olApp.CreateItem(olMailItem)
Dim RutaImagen As String

With olMail
.Display
.To = "lrivera@almacontactcol.co; daguarin@almacontactcol.co"
.CC = "mramos@almacontactcol.co"
.Subject = "[Privado] Informe de Cortes SFTP"
.HTMLBody = "<H6> <p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Buenos d&iacute;as,</span></span></em></strong></p></H6></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'><span style='color:#002060;'>Archivo de cortes FSTP.</span></span></em></strong></p>" & _
            "<H6><p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Actualizado:</span><span style='color:#EE1750;'>&nbsp;</span><span style='color:#C00000;'>" & Format(Date - 1, "DD/MM/YYYY") & "</span></span></em></strong></p></H6></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'><span style='color:#002060;'>Agradecemos tú ayuda calificando nuestra calidad en el siguiente enlace:  <a href='https://forms.office.com/r/4scK4ZKD1G'>https://forms.office.com/r/4scK4ZKD1G</a></span></span></em></strong></p></p></ul>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style='color:#002060;'></span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Saludos.</span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>¡Que cada día nos haga mejores personas!</span></span></em></strong><font size=3>&#129299<\font></p>" & .HTMLBody
                        
.Display
.Send
End With

End Sub


















DISTRIBUCION ARCHIVOS POC

Sub ActualizacionDeDatosPOC()

Call CargaDeDatosCLAROMDE
Call CargaDeDatosGP1
Call CargaDeDatosGP2NOVOZ
Call CargaDeDatosGP2VOZ
Call CargaDeDatosGP3
Call CargaDeDatosGP4
Call CargaDeDatosGP5
Call CargaDeDatosBOGOTA

  MsgBox ("**FINALIZADO**")

End Sub

Sub CargaDeDatosCLAROMDE()

ThisWorkbook.Worksheets("NOM_ALMACONTACT").Activate
Range("B1").AutoFilter 2, "CLARO MDE"                                             'Filtro Claro Medellin en la plantilla
Workbooks.Open Worksheets("Inputs").Range("D3"), , , , , "POCALMA2022", True      'Abre POC Claro Medellin
Worksheets("AMC_NOM").Unprotect "Planeacion2022"
If ActiveSheet.AutoFilterMode Then
ActiveSheet.AutoFilterMode = False
End If
  
  ThisWorkbook.Activate                                                           'Retornamos a la plantilla para copiar los datos
  Range("A1").Select
  SendKeys ("{DOWN}"), True
  For i = 1 To 14
  SendKeys "+{RIGHT}", True
  Next i
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy
  
Workbooks(2).Activate                                                              'Vuelve a activar el archivo de Claro Medellin
Range("A1").Select

If Range("A2") = "" Then                                                           'Condicion inicio de mes
SendKeys ("{DOWN}"), True
Else
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
End If

Selection.PasteSpecial xlPasteValues                                               'Pegamos la informacion en el archivo de POC y corremos formato y formulas
Range("A2:AV2").Select
Selection.Copy
Range(Selection, Selection.End(xlDown)).Select
Selection.PasteSpecial xlPasteFormats
Range("M1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 8
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("W1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 2
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("AB1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 17
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
 
  If Not ActiveSheet.AutoFilterMode Then
  ActiveSheet.Range("A1").AutoFilter
  End If
  ActiveWindow.Zoom = 60
  Worksheets("AMC_NOM").Protect "Planeacion2022", , , , , , , , , , , , , , True
  ActiveWorkbook.Close savechanges:=True
  
End Sub
Sub CargaDeDatosGP1()

ThisWorkbook.Worksheets("NOM_ALMACONTACT").Activate
Range("B1").AutoFilter 2, "GP1"                                                   'Filtro GP1 en la plantilla
Workbooks.Open Worksheets("Inputs").Range("D4"), , , , , "POCALMA2022", True      'Abre POC GP1
Worksheets("AMC_NOM").Unprotect "Planeacion2022"
If ActiveSheet.AutoFilterMode Then
ActiveSheet.AutoFilterMode = False
End If
  
  ThisWorkbook.Activate                                                           'Retornamos a la plantilla para copiar los datos
  Range("A1").Select
  SendKeys ("{DOWN}"), True
  For i = 1 To 14
  SendKeys "+{RIGHT}", True
  Next i
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy
  
Workbooks(2).Activate                                                              'Vuelve a activar el archivo de GP1
Range("A1").Select

If Range("A2") = "" Then                                                           'Condicion inicio de mes
SendKeys ("{DOWN}"), True
Else
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
End If

Selection.PasteSpecial xlPasteValues                                               'Pegamos la informacion en el archivo de POC y corremos formato y formulas
Range("A2:AV2").Select
Selection.Copy
Range(Selection, Selection.End(xlDown)).Select
Selection.PasteSpecial xlPasteFormats
Range("M1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 8
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("W1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 2
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("AB1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 17
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
  
  If Not ActiveSheet.AutoFilterMode Then
  ActiveSheet.Range("A1").AutoFilter
  End If
  ActiveWindow.Zoom = 60
  Worksheets("AMC_NOM").Protect "Planeacion2022", , , , , , , , , , , , , , True
  ActiveWorkbook.Close savechanges:=True
  
End Sub

Sub CargaDeDatosGP2NOVOZ()

ThisWorkbook.Worksheets("NOM_ALMACONTACT").Activate
Range("B1").AutoFilter 2, "GP2NOVOZ"                                                   'Filtro GP2NOVOZ en la plantilla
Workbooks.Open Worksheets("Inputs").Range("D5"), , , , , "POCALMA2022", True      'Abre POC GP2NOVOZ
Worksheets("AMC_NOM").Unprotect "Planeacion2022"
If ActiveSheet.AutoFilterMode Then
ActiveSheet.AutoFilterMode = False
End If
  
  ThisWorkbook.Activate                                                           'Retornamos a la plantilla para copiar los datos
  Range("A1").Select
  SendKeys ("{DOWN}"), True
  For i = 1 To 14
  SendKeys "+{RIGHT}", True
  Next i
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy
  
Workbooks(2).Activate                                                              'Vuelve a activar el archivo de GP2NOVOZ
Range("A1").Select

If Range("A2") = "" Then                                                           'Condicion inicio de mes
SendKeys ("{DOWN}"), True
Else
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
End If

Selection.PasteSpecial xlPasteValues                                               'Pegamos la informacion en el archivo de POC y corremos formato y formulas
Range("A2:AV2").Select
Selection.Copy
Range(Selection, Selection.End(xlDown)).Select
Selection.PasteSpecial xlPasteFormats
Range("M1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 8
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("W1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 2
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("AB1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 17
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
  
  If Not ActiveSheet.AutoFilterMode Then
  ActiveSheet.Range("A1").AutoFilter
  End If
  ActiveWindow.Zoom = 60
  Worksheets("AMC_NOM").Protect "Planeacion2022", , , , , , , , , , , , , , True
  ActiveWorkbook.Close savechanges:=True

End Sub

Sub CargaDeDatosGP2VOZ()

ThisWorkbook.Worksheets("NOM_ALMACONTACT").Activate
Range("B1").AutoFilter 2, "GP2VOZ"                                                   'Filtro GP2VOZ en la plantilla
Workbooks.Open Worksheets("Inputs").Range("D6"), , , , , "POCALMA2022", True      'Abre POC GP2VOZ
Worksheets("AMC_NOM").Unprotect "Planeacion2022"
If ActiveSheet.AutoFilterMode Then
ActiveSheet.AutoFilterMode = False
End If
  
  ThisWorkbook.Activate                                                           'Retornamos a la plantilla para copiar los datos
  Range("A1").Select
  SendKeys ("{DOWN}"), True
  For i = 1 To 14
  SendKeys "+{RIGHT}", True
  Next i
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy
  
Workbooks(2).Activate                                                              'Vuelve a activar el archivo de GP2VOZ
Range("A1").Select

If Range("A2") = "" Then                                                           'Condicion inicio de mes
SendKeys ("{DOWN}"), True
Else
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
End If

Selection.PasteSpecial xlPasteValues                                               'Pegamos la informacion en el archivo de POC y corremos formato y formulas
Range("A2:AV2").Select
Selection.Copy
Range(Selection, Selection.End(xlDown)).Select
Selection.PasteSpecial xlPasteFormats
Range("M1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 8
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("W1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 2
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("AB1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 17
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
  
  If Not ActiveSheet.AutoFilterMode Then
  ActiveSheet.Range("A1").AutoFilter
  End If
  ActiveWindow.Zoom = 60
  Worksheets("AMC_NOM").Protect "Planeacion2022", , , , , , , , , , , , , , True
  ActiveWorkbook.Close savechanges:=True
  
End Sub

Sub CargaDeDatosGP3()

ThisWorkbook.Worksheets("NOM_ALMACONTACT").Activate
Range("B1").AutoFilter 2, "GP3"                                                   'Filtro GP3 en la plantilla
Workbooks.Open Worksheets("Inputs").Range("D7"), , , , , "POCALMA2022", True      'Abre POC GP3
Worksheets("AMC_NOM").Unprotect "Planeacion2022"
If ActiveSheet.AutoFilterMode Then
ActiveSheet.AutoFilterMode = False
End If
  
  ThisWorkbook.Activate                                                           'Retornamos a la plantilla para copiar los datos
  Range("A1").Select
  SendKeys ("{DOWN}"), True
  For i = 1 To 14
  SendKeys "+{RIGHT}", True
  Next i
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy
  
Workbooks(2).Activate                                                              'Vuelve a activar el archivo de GP3
Range("A1").Select

If Range("A2") = "" Then                                                           'Condicion inicio de mes
SendKeys ("{DOWN}"), True
Else
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
End If

Selection.PasteSpecial xlPasteValues                                               'Pegamos la informacion en el archivo de POC y corremos formato y formulas
Range("A2:AV2").Select
Selection.Copy
Range(Selection, Selection.End(xlDown)).Select
Selection.PasteSpecial xlPasteFormats
Range("M1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 8
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("W1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 2
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("AB1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 17
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
  
  If Not ActiveSheet.AutoFilterMode Then
  ActiveSheet.Range("A1").AutoFilter
  End If
  ActiveWindow.Zoom = 60
  Worksheets("AMC_NOM").Protect "Planeacion2022", , , , , , , , , , , , , , True
  ActiveWorkbook.Close savechanges:=True

End Sub

Sub CargaDeDatosGP4()

ThisWorkbook.Worksheets("NOM_ALMACONTACT").Activate
Range("B1").AutoFilter 2, "GP4"                                                   'Filtro GP4 en la plantilla
Workbooks.Open Worksheets("Inputs").Range("D8"), , , , , "POCALMA2022", True      'Abre POC GP4
Worksheets("AMC_NOM").Unprotect "Planeacion2022"
If ActiveSheet.AutoFilterMode Then
ActiveSheet.AutoFilterMode = False
End If
  
  ThisWorkbook.Activate                                                           'Retornamos a la plantilla para copiar los datos
  Range("A1").Select
  SendKeys ("{DOWN}"), True
  For i = 1 To 14
  SendKeys "+{RIGHT}", True
  Next i
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy
  
Workbooks(2).Activate                                                              'Vuelve a activar el archivo de GP4
Range("A1").Select

If Range("A2") = "" Then                                                           'Condicion inicio de mes
SendKeys ("{DOWN}"), True
Else
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
End If

Selection.PasteSpecial xlPasteValues                                               'Pegamos la informacion en el archivo de POC y corremos formato y formulas
Range("A2:AV2").Select
Selection.Copy
Range(Selection, Selection.End(xlDown)).Select
Selection.PasteSpecial xlPasteFormats
Range("M1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 8
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("W1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 2
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("AB1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 17
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
  
  If Not ActiveSheet.AutoFilterMode Then
  ActiveSheet.Range("A1").AutoFilter
  End If
  ActiveWindow.Zoom = 60
  Worksheets("AMC_NOM").Protect "Planeacion2022", , , , , , , , , , , , , , True
  ActiveWorkbook.Close savechanges:=True

End Sub

Sub CargaDeDatosGP5()

ThisWorkbook.Worksheets("NOM_ALMACONTACT").Activate
Range("B1").AutoFilter 2, "GP5"                                                   'Filtro GP5 en la plantilla
Workbooks.Open Worksheets("Inputs").Range("D9"), , , , , "POCALMA2022", True      'Abre POC GP5
Worksheets("AMC_NOM").Unprotect "Planeacion2022"
If ActiveSheet.AutoFilterMode Then
ActiveSheet.AutoFilterMode = False
End If
  
  ThisWorkbook.Activate                                                           'Retornamos a la plantilla para copiar los datos
  Range("A1").Select
  SendKeys ("{DOWN}"), True
  For i = 1 To 14
  SendKeys "+{RIGHT}", True
  Next i
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy
  
Workbooks(2).Activate                                                              'Vuelve a activar el archivo de GP5
Range("A1").Select

If Range("A2") = "" Then                                                           'Condicion inicio de mes
SendKeys ("{DOWN}"), True
Else
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
End If

Selection.PasteSpecial xlPasteValues                                               'Pegamos la informacion en el archivo de POC y corremos formato y formulas
Range("A2:AV2").Select
Selection.Copy
Range(Selection, Selection.End(xlDown)).Select
Selection.PasteSpecial xlPasteFormats
Range("M1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 8
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("W1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 2
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("AB1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 17
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
  
  If Not ActiveSheet.AutoFilterMode Then
  ActiveSheet.Range("A1").AutoFilter
  End If
  ActiveWindow.Zoom = 60
  Worksheets("AMC_NOM").Protect "Planeacion2022", , , , , , , , , , , , , , True
  ActiveWorkbook.Close savechanges:=True

End Sub

Sub CargaDeDatosBOGOTA()

ThisWorkbook.Worksheets("NOM_ALMACONTACT").Activate
Range("B1").AutoFilter 2, Array("CLARO BOG", "SAMSUNG", "DINISSAN", "VARDI", "CONSULADO", "HISENSE", "FICOHSA", "CLARO_VIP", "SHOPEE", "WOM", "FORTINET", "WOM CHILE", "AITEL VODAFONE", "FILTROS", "OIT", "GOL"), xlFilterValues     'Filtro BOGOTA en la plantilla
Workbooks.Open Worksheets("Inputs").Range("D10"), , , , , "POCALMA2022", True      'Abre POC BOGOTA
Worksheets("AMC_NOM").Unprotect "Planeacion2022"
If ActiveSheet.AutoFilterMode Then
ActiveSheet.AutoFilterMode = False
End If
  
  ThisWorkbook.Activate                                                           'Retornamos a la plantilla para copiar los datos
  Range("A1").Select
  SendKeys ("{DOWN}"), True
  For i = 1 To 14
  SendKeys "+{RIGHT}", True
  Next i
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Copy
  
Workbooks(2).Activate                                                              'Vuelve a activar el archivo de BOGOTA
Range("A1").Select

If Range("A2") = "" Then                                                           'Condicion inicio de mes
SendKeys ("{DOWN}"), True
Else
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
End If

Selection.PasteSpecial xlPasteValues                                               'Pegamos la informacion en el archivo de POC y corremos formato y formulas
Range("A2:AV2").Select
Selection.Copy
Range(Selection, Selection.End(xlDown)).Select
Selection.PasteSpecial xlPasteFormats
Range("M1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 8
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("W1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 2
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Range("AB1").Select
Selection.End(xlDown).Select
For i = 1 To 3
SendKeys "{RIGHT}", True
Next i
For i = 1 To 17
SendKeys "+{RIGHT}", True
Next i
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
  
  If Not ActiveSheet.AutoFilterMode Then
  ActiveSheet.Range("A1").AutoFilter
  End If
  ActiveWindow.Zoom = 60
  Worksheets("AMC_NOM").Protect "Planeacion2022", , , , , , , , , , , , , , True
  ActiveWorkbook.Close savechanges:=True

End Sub
















CONSOLIDADO DE NOMINA SQL

Sub CargaDeDatos()

    ThisWorkbook.Sheets("Consolidado").Select
    Range("A3:AW3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A3").Select
    
  Workbooks.Open ("\\10.96.16.27\controlgestion$\[Privado] " & Format(Date, "YYYY") & "\[Privado] " & Format(Date, "MM") & "-" & Format(Date, "MMMM") & "\[Privado] 001-CONSOLIDADO OPERACIONES\Privado " & Format(Date, "MM") & "-AMC_Consolidado_NOM_Almacontact.xlsx"), , True    'Abrir archivo de Consolidado de nomina'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Application.Goto (ActiveWorkbook.Sheets("TEM_Carga Nomina").Range("A2:L2"))
  Range(Selection, Selection.End(xlDown)).Copy
  
ThisWorkbook.Activate                                      'Pegado en la base de Nomina SQL'
Range("A3").PasteSpecial xlPasteValues

  Workbooks(2).Activate                                    'Abrir archivo de Consolidado de nomina'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Application.Goto (ActiveWorkbook.Sheets("TEM_Carga Nomina").Range("O2:P2"))
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                      'Pegado en la base de Nomina SQL'
Range("M3").PasteSpecial xlPasteValues

  Workbooks(2).Activate                                    'Abrir archivo de Consolidado de nomina'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Application.Goto (ActiveWorkbook.Sheets("TEM_Carga Nomina").Range("L2:N2"))
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                      'Pegado en la base de Nomina SQL'
Range("O3").PasteSpecial xlPasteValues

  Workbooks(2).Activate                                    'Abrir archivo de Consolidado de nomina'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Application.Goto (ActiveWorkbook.Sheets("TEM_Carga Nomina").Range("Q2"))
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                      'Pegado en la base de Nomina SQL'
Range("O3").PasteSpecial xlPasteValues

  Workbooks(2).Activate                                    'Abrir archivo de Consolidado de nomina'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Application.Goto (ActiveWorkbook.Sheets("TEM_Carga Nomina").Range("R2"))
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                      'Pegado en la base de Nomina SQL'
Range("R3").PasteSpecial xlPasteValues

  Range("T1").Copy                                         'Valores manuales en el archivo'
  Range("S3").PasteSpecial xlPasteValues
  Range("R3").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown
  Range("T1").Copy
  Range("T3").PasteSpecial xlPasteFormulas
  Range("S3").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown
  Selection.Copy
  Selection.PasteSpecial xlPasteValues
  
Workbooks(2).Activate                                    'Abrir archivo de Consolidado de nomina'
Visible = True
WindowState = xlMaximized
ActiveSheet.AutoFilterMode = False
Application.Goto (ActiveWorkbook.Sheets("TEM_Carga Nomina").Range("S2:AU2"))
Range(Selection, Selection.End(xlDown)).Copy

  ThisWorkbook.Activate                                      'Pegado en la base de Nomina SQL'
  Range("U3").PasteSpecial xlPasteValues

Range("AX3").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
For i = 1 To 7
SendKeys "^+{LEFT}", True
Next i
Selection.Copy
SendKeys "^+{DOWN}", True
Selection.PasteSpecial xlPasteFormats
Range("AW3").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
For i = 1 To 13
SendKeys "+{RIGHT}", True
Next i
SendKeys "^+{UP}", True
Selection.FillDown
   
  Workbooks(2).Activate                                    'Abrir archivo de Consolidado de nomina'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Application.Goto (ActiveWorkbook.Sheets("TEM_Carga Nomina").Range("BJ2"))
  Range(Selection, Selection.End(xlDown)).Copy
  
ThisWorkbook.Activate                                      'Pegado en la base de Nomina SQL'
Range("BL3").PasteSpecial xlPasteValues

  Workbooks(2).Activate                                    'Abrir archivo de Consolidado de nomina'
  Application.DisplayAlerts = False
  ActiveWorkbook.Close savechanges:=False
  Application.DisplayAlerts = False
  Range("A2").Select
  ThisWorkbook.Save
  
Range("1:1").Delete
Range("AX2:BL2").Select
SendKeys "^+{DOWN}", True
Selection.Copy
Selection.PasteSpecial xlPasteValues

Application.DisplayAlerts = Falso
ThisWorkbook.Sheets("Validador").Delete
Application.DisplayAlerts = True
Application.DisplayAlerts = Falso
ThisWorkbook.Sheets("Servicios").Delete
Application.DisplayAlerts = True
Application.DisplayAlerts = Falso
ThisWorkbook.Sheets("imputs").Delete
Application.DisplayAlerts = True
Application.DisplayAlerts = Falso
ThisWorkbook.Sheets("Documentacion").Delete
Application.DisplayAlerts = True

  Application.DisplayAlerts = False
  Range("A2").Select
  ActiveWorkbook.SaveAs Filename:="\\Co0000fs0001\planeacion$\01_LATAM\01_WFM\01- CARGAS BD\8- NOVEDADES NOMINA\" & Format(Date - 1, "YYYY") & "\" & ActiveWorkbook.Name
  
  
Call EnviarCorreoLatam
Call EnviarCorreoMulticampañas
  
  Range("A2").Select
  Application.DisplayAlerts = True
  ActiveWorkbook.Close

End Sub















CORREO OPERACION LATAM

Sub EnviarCorreoLatam()

Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")
Dim olMail As Outlook.MailItem
Set olMail = olApp.CreateItem(olMailItem)
Dim RutaImagen As String
 
 RutaImagen = "\\10.96.16.27\reporting_almacontact\2022\02.Privado\01. [Privado] Analistas\02. [Privado] Julian Cardona\[Privado] Logos Clientes\Latam.png"

With olMail
.Display
.To = "jdperalta@almacontactcol.info; lfjimenez@almacontactcol.info; navilladiego@almacontactcol.info; "
.CC = "jsmanrique@almacontactcol.co; mramos@almacontactcol.co; earango@almacontactcol.co"
.Subject = "[Privado] Informe de Nomina LATAM"
.HTMLBody = "<p><img src='cid:Latam.png' height=50 width=250></p>" & _
            "<H6> <p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style='color:#002060;'>Buenos d&iacute;as,</span></span></em></strong></p></H6></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;color:red;'>&nbsp;</span></em></strong></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style='color:#002060;'>Informe de Nomina.</span></span></em></strong></p>" & _
            "<H6><p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style='color:#002060;'>Actualizado:</span><span style='color:#EE1750;'>&nbsp;</span><span style='color:#C00000;'>" & Format(Date - 1, "DD/MM/YYYY") & "</span></span></em></strong></p></H6></Body>" & _
            "<ul style='margin-bottom:0cm;margin-top:0cm;;color:#002060;' type='disc'>" & _
            "<li style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'>En caso tal de tener que validar LNR, cambios de turno y dem&aacute;s, los invito a consultar el siguiente portal WEB, all&iacute; los supervisores realizan todo este tipo de solicitudes y se les es respondido.</span></em></strong></li></ul>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <span style='color:#002060;'>Link <a href='https://sites.google.com/view/planeacin/men%C3%BA?authuser=0\'>https://sites.google.com/view/planeacin/men%C3%BA?authuser=0\</a></span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <span style='color:#002060;'>User: <a href='mailto:operaalmacontact@gmail.com'><span style='color:#002060;text-decoration:none;'>operaalmacontact@gmail.com</span></a></span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style='color:#002060;'>pass: Clave_segura2018$2</span></span></em></strong></p>" & _
            "<ul style='margin-bottom:0cm;margin-top:0cm;color:#002060;' type='disc'>" & _
            "<li style='margin-top:0cm;margin-right:0cm;margin-bottom:0cm;margin-left:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'>Link de Validaci&oacute;n para asistencia de personal a Formaci&oacute;n programada por malla</span></em></strong></li></ul>" & _
            "<p><strong><em><span style='font-size:12px;font-family:'Calibri',sans-serif;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <span style='color:#002060;'><a href='https://docs.google.com/spreadsheets/d/1-Rq6mbYmv9qfHc6AjPt3rR15nFbWDt1PHRBj6c37Ji4/edit?usp=sharing\'>https://docs.google.com/spreadsheets/d/1-Rq6mbYmv9qfHc6AjPt3rR15nFbWDt1PHRBj6c37Ji4/edit?usp=sharing\</a></span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'><span style='color:#002060;'>Agradecemos tú ayuda calificando nuestra calidad en el siguiente enlace:  <a href='https://forms.office.com/r/4scK4ZKD1G'>https://forms.office.com/r/4scK4ZKD1G</a></span></span></em></strong></p></p></ul>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style='color:#002060;'></span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Saludos.</span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>¡Que cada día nos haga mejores personas!</span></span></em></strong><font size=3>&#129299<\font></p>" & .HTMLBody
            
.Attachments.Add RutaImagen
.Display
.Send
End With

End Sub




































CORREO MULTICAMPAÑAS

Sub EnviarCorreoMulticampañas()

Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")
Dim olMail As Outlook.MailItem
Set olMail = olApp.CreateItem(olMailItem)
Dim RutaImagen As String

 RutaImagen = "\\10.96.16.27\reporting_almacontact\2022\02.Privado\01. [Privado] Analistas\02. [Privado] Julian Cardona\[Privado] Logos Clientes\Bogota.png"

With olMail
.Display
.To = "ymunoz@almacontactcol.info; nominabogota2022@gmail.com"
.CC = "jsmanrique@almacontactcol.co; mramos@almacontactcol.co; jfarbelaez@almacontactcol.co; earango@almacontactcol.co"
.Subject = "[Privado] Informe de Nomina Multicampañas"
.HTMLBody = "<p><img src='cid:Bogota.png' height=120 width=830></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Buenos d&iacute;as,</span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;color:red;'>&nbsp;</span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'><span style='color:#002060;'>Informe de Nomina.</span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Actualizado:</span><span style='color:#EE1750;'>&nbsp;</span><span style='color:#C00000;'>" & Format(Date - 1, "DD/MM/YYYY") & "</span></span></em></strong></p>" & _
                        "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'><span style='color:#002060;'>Agradecemos tú ayuda calificando nuestra calidad en el siguiente enlace:  <a href='https://forms.office.com/r/4scK4ZKD1G'>https://forms.office.com/r/4scK4ZKD1G</a></span></span></em></strong></p></p></ul>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style='color:#002060;'></span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Saludos.</span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>¡Que cada día nos haga mejores personas!</span></span></em></strong><font size=3>&#129299<\font></p>" & .HTMLBody
            
.Attachments.Add RutaImagen
.Display
.Send
End With

End Sub








CORREO BASE DE NOMINA SQL

Sub EnviarCorreoBaseNom()

Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")
Dim olMail As Outlook.MailItem
Set olMail = olApp.CreateItem(olMailItem)
Dim RutaImagen As String
 
With olMail
.Display
.To = "lrivera@almacontactcol.co; daguarin@almacontactcol.co; esoto@almacontactcol.co; jabedoya@almacontactcol.co; mlopez@almacontactcol.co; eatorres@almacontactcol.info; jchindoy@almacontactcol.co"
.CC = "mramos@almacontactcol.co"
.Subject = "[Privado] Actualización base de nomina SQL"
.HTMLBody = "<H6> <p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Buenos d&iacute;as,</span></span></em></strong></p></H6></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'><span style='color:#002060;'>Base de Nomina SQL.</span></span></em></strong></p>" & _
            "<H6><p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Actualizada:</span><span style='color:#EE1750;'>&nbsp;</span><span style='color:#C00000;'>" & Format(Date - 1, "DD/MM/YYYY") & "</span></span></em></strong></p></H6></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>En caso de tener novedades por favor contactar el analista encargado para su validación.</span></span></em></strong></p>" & _
                        "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'><span style='color:#002060;'>Agradecemos tú ayuda calificando nuestra calidad en el siguiente enlace:  <a href='https://forms.office.com/r/4scK4ZKD1G'>https://forms.office.com/r/4scK4ZKD1G</a></span></span></em></strong></p></p></ul>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style='color:#002060;'></span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Saludos.</span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>¡Que cada día nos haga mejores personas!</span></span></em></strong><font size=3>&#129299<\font></p>" & .HTMLBody
            
.Display
.Send
End With

End Sub










CARGA DE MALLAS PARA SQL

Sub CargaDatos()

'Application.ScreenUpdating = False

  ThisWorkbook.ActiveSheet.Cells.Range("B1").Select           'Borrado de los datos anteriores
  Range("B2", Range("B2").End(xlDown)).ClearContents
  Range("D2:S2", Range("D2:M2").End(xlDown)).ClearContents
  Range("C3", Range("C3").End(xlDown)).ClearContents
  Range("A3", Range("A3").End(xlDown)).ClearContents
  Range("B2", Range("B2").End(xlDown)).ClearContents

If ActiveWorkbook.Sheets("hoja2").Range("D1") = "lunes" Then                                  'Condiciones día para ejecucion de macro
Call DatosLunes.DatosLunes
End If

  If ActiveWorkbook.Sheets("hoja2").Range("D1") = "martes" Then
  Call DatosMartes.DatosMartes
  End If

If ActiveWorkbook.Sheets("hoja2").Range("D1") = "miércoles" Then
Call DatosMiercoles.DatosMiercoles
End If

  If ActiveWorkbook.Sheets("hoja2").Range("D1") = "jueves" Then
  Call DatosJueves.DatosJueves
  End If

If ActiveWorkbook.Sheets("hoja2").Range("D1") = "viernes" Then
Call DatosViernes.DatosViernes
End If

  MsgBox ("¡FINALIZADO!")

End Sub




Sub DatosEjecutivosLatam()

'Malla de turnos LATAM'

  Workbooks.Open Worksheets("Hoja2").Range("B6"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  Sheets("TURNOS").Activate
  ActiveSheet.AutoFilterMode = False
  Range("L5:N5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:G2").PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2").PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2").PasteSpecial xlPasteValues

End Sub

Sub DatosEjecutivosClaroMedellin()
 
'Malla de turnos CLARO MEDELLIN'

  Workbooks.Open Worksheets("Hoja2").Range("B7"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("L5:N5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2:H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2:K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues

End Sub

Sub DatosEjecutivosClaroBogota()

'Malla de turnos CLARO BOGOTA'

  Workbooks.Open Worksheets("Hoja2").Range("B8"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("L5:N5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2:H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2:K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  
  End Sub
  
Sub DatosEjecutivosClaroVIP()

'Malla de turnos CLARO VIP'

  Workbooks.Open Worksheets("Hoja2").Range("B9"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("L5:M5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:F2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  Range("O5", Range("O5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2:H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2:K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues

End Sub

Sub DatosEjecutivosFicohsa()

'Malla de turnos FICOHSA'

  Workbooks.Open Worksheets("Hoja2").Range("B10"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("L5:N5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2:H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2:K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  
End Sub

Sub DatosEjecutivosSamsung()

'Malla de turnos SAMSUNG'

  Workbooks.Open Worksheets("Hoja2").Range("B11"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("L5:M5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:F2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  Range("O5", Range("O5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2:H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2:K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  
End Sub

Sub DatosEjecutivosShopee()

'Malla de turnos SHOPEE'

  Workbooks.Open Worksheets("Hoja2").Range("B12"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("L5:N5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  
End Sub

Sub DatosEjecutivosHisense()

'Malla de turnos HISENSE'

  Workbooks.Open Worksheets("Hoja2").Range("B13"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("L5:N5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  
End Sub

Sub DatosEjecutivosFortinet()

'Malla de turnos FORTINET'

  Workbooks.Open Worksheets("Hoja2").Range("B14"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("L5:N5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  
End Sub

Sub DatosEjecutivosConsulado()

'Malla de turnos CONSULADO'

  Workbooks.Open Worksheets("Hoja2").Range("B15"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("L5:N5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  
End Sub

Sub DatosEjecutivosDinissan()

'Malla de turnos DINISSAN'

  Workbooks.Open Worksheets("Hoja2").Range("B16"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("O4").AutoFilter Field:=14, Criteria1:="DINISSAN"
  Range("L5:N5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  
End Sub

Sub DatosEjecutivosVardi()

'Malla de turnos VARDI'

  Workbooks.Open Worksheets("Hoja2").Range("B16"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("O4").AutoFilter Field:=14, Criteria1:="VARDI"
  Range("L5:N5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  
End Sub

Sub DatosEjecutivosWom()

'Malla de turnos WOM'

  Workbooks.Open Worksheets("Hoja2").Range("B17"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("L5:N5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  
End Sub

Sub DatosEjecutivosAitelVodafone()

'Malla de turnos AITEL'

  Workbooks.Open Worksheets("Hoja2").Range("B18"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("L5:M5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:F2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  Range("O5", Range("O5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2:H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2:K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  
End Sub

Sub DatosEjecutivosDisponible()

'Malla de turnos COLMENA'

'  Workbooks.Open Worksheets("Hoja2").Range("B19"), , True     'Abrir archivo de malla'
'  Visible = True
'  WindowState = xlMaximized
'  ActiveSheet.AutoFilterMode = False
'  Range("L5:N5").Copy            'Copiar datos de ejecutivo desde la malla'
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("E2:G2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  Range("R5").Copy                  'Copiar datos de ejecutivo desde la malla'
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("H2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  Range("P5").Copy                  'Copiar datos de ejecutivo desde la malla'
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("K2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
  
End Sub

Sub DatosEjecutivosDisponible2()

'Malla de turnos FILTROS'

'  Workbooks.Open Worksheets("Hoja2").Range("B20"), , True     'Abrir archivo de malla'
'  Visible = True
'  WindowState = xlMaximized
'  ActiveSheet.AutoFilterMode = False
'  Range("L5:N5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("E2:G2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("H2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("K2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
  
End Sub

Sub DatosEjecutivosDisponible1()

''Malla de turnos WOM CHILE'
'
'  Workbooks.Open Worksheets("Hoja2").Range("B21"), , True     'Abrir archivo de malla'
'  Visible = True
'  WindowState = xlMaximized
'  ActiveSheet.AutoFilterMode = False
'  Range("L5:N5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("E2:G2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("H2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("K2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
  
End Sub
Sub DatosEjecutivosOIT()

'Malla de turnos OIT'

  Workbooks.Open Worksheets("Hoja2").Range("B22"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("L5:M5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:F2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  Range("O5", Range("O5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2:H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2:K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues

End Sub

Sub DatosEjecutivosGol()

'Malla de turnos GOLD'

  Workbooks.Open Worksheets("Hoja2").Range("B23"), , True     'Abrir archivo de malla'
  Visible = True
  WindowState = xlMaximized
  ActiveSheet.AutoFilterMode = False
  Range("L5:N5", Range("L5:N5").End(xlDown)).Copy            'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("E2:G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("R5", Range("R5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  Range("P5", Range("P5").End(xlDown)).Copy                  'Copiar datos de ejecutivo desde la malla'
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  
End Sub
Sub DatosLunes()

'Malla de turnos LATAM'

  Call DatosEjecutivosLatam
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Range("D2").PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Range("E2:K2", Range("E2:K2").End(xlDown)).Copy             'Pegado de datos de los ejecutivos en los demas días
  ActiveSheet.Range("E2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveSheet.Paste
  ActiveSheet.Range("E2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveSheet.Paste
 
ActiveSheet.Range("D2").Select                              'Fechas faltantes en columna de fecha
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.Formula = ("=D2+1")
ActiveSheet.Range("E2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:CF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CK5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:DB5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
ActiveCell.Value = "LATAM"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroMedellinL

End Sub
 
Sub MallaClaroMedellinL()
 
'Malla de turnos CLARO MEDELLIN'

  Call DatosEjecutivosClaroMedellin
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Call DatosEjecutivosClaroMedellin
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosClaroMedellin
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:CF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CQ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:DB5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO MDE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
  
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroBogotaL
    
End Sub

Sub MallaClaroBogotaL()

'Malla de turnos CLARO BOGOTA'

  Call DatosEjecutivosClaroBogota
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Call DatosEjecutivosClaroBogota
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosClaroBogota
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:CF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CQ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:DB5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO BOG"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroVIPL
    
End Sub

Sub MallaClaroVIPL()

'Malla de turnos CLARO VIP'

  Call DatosEjecutivosClaroVIP
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Call DatosEjecutivosClaroVIP

  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosClaroVIP

  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:CF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CQ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:DB5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO_VIP"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaFicohsaL
    
End Sub

Sub MallaFicohsaL()

'Malla de turnos FICOHSA'

  Call DatosEjecutivosFicohsa
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Call DatosEjecutivosFicohsa
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosFicohsa
 
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:BZ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CK5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:CV5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "FICOHSA"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaSamsungL
    
End Sub
Sub MallaSamsungL()

'Malla de turnos SAMSUNG'

  Call DatosEjecutivosSamsung
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Call DatosEjecutivosSamsung
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosSamsung
 
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:CF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CQ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:DB5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "SAMSUNG"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaShopeeL
    
End Sub

Sub MallaShopeeL()

'Malla de turnos SHOPEE'

  Call DatosEjecutivosShopee
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Call DatosEjecutivosShopee
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosShopee
 
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:BZ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CK5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:CV5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "SHOPEE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
    Workbooks(2).Activate
    ActiveWorkbook.Close savechanges:=False
    Call MallaHisenseL
    
End Sub

Sub MallaHisenseL()

'Malla de turnos HISENSE'

  Call DatosEjecutivosHisense
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Call DatosEjecutivosHisense
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosHisense
 
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:BZ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CK5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:CV5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "HISENSE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
    Workbooks(2).Activate
    ActiveWorkbook.Close savechanges:=False
    Call MallaFortinetL
    
End Sub

Sub MallaFortinetL()

'Malla de turnos FORTINET'

  Call DatosEjecutivosFortinet
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Call DatosEjecutivosFortinet
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosFortinet
 
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:BZ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CK5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:CV5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "FORTINET"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaConsuladoL
    
End Sub


Sub MallaConsuladoL()

'Malla de turnos CONSULADO'

  Call DatosEjecutivosConsulado
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Call DatosEjecutivosConsulado
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosConsulado
 
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:BZ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CK5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:CV5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CONSULADO"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaDinissanL
    
End Sub

Sub MallaDinissanL()

'Malla de turnos DINISSAN'

  Call DatosEjecutivosDinissan
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Call DatosEjecutivosDinissan
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosDinissan
 
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:CF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CQ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:DB5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "DINISSAN"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
     
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaVardiL
    
End Sub

Sub MallaVardiL()

'Malla de turnos VARDI'

  Call DatosEjecutivosVardi
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Call DatosEjecutivosVardi
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosVardi
 
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:CF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CQ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:DB5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "VARDI"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
   
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaWomL
    
End Sub

Sub MallaWomL()

'Malla de turnos WOM'

  Call DatosEjecutivosWom
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Call DatosEjecutivosWom
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosWom
 
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:CF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CQ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:DB5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "WOM"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaAitelVodafoneL
  
End Sub

Sub MallaAitelVodafoneL()

'Malla de turnos AITEL'

  Call DatosEjecutivosAitelVodafone

Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosAitelVodafone

Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosAitelVodafone

Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:BZ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CK5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:CV5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues

Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "AITEL VODAFONE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaOITL
    
End Sub

Sub MallaDisponibleL()

'Malla de turnos COLMENA'

'  Call DatosEjecutivosColmena
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B1").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'
'  Call DatosEjecutivosColmena
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B2").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'
'  Call DatosEjecutivosColmena
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BT5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BW5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BY5:CF5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("CG5:CH5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2:J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("CJ5:CQ5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("CR5:CS5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2:J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("CU5:DB5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "COLMENA"
'Range("C2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'   Workbooks(2).Activate
'   ActiveWorkbook.Close savechanges:=False
'   Call MallaFiltrosL
    
End Sub

Sub MallaDisponible2L()

'Malla de turnos FILTROS'

'  Call DatosEjecutivosFiltros
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B1").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'ActiveSheet.Range("E1").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Call DatosEjecutivosFiltros
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B2").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'ActiveSheet.Range("E1").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Call DatosEjecutivosFiltros
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'ActiveSheet.Range("E1").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BT5", Range("BT5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BW5", Range("BW5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BY5:CF5", Range("BY5:BZ5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2:J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("CJ5:CQ5", Range("CJ5:CK5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2:J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("CU5:DB5", Range("CU5:CV5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "FILTROS"
'Range("C2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'   Workbooks(2).Activate
'   ActiveWorkbook.Close savechanges:=False
'   Call MallaOITL
    
End Sub

Sub MallaDisponible1L()

''Malla de turnos WOM CHILE'
'
'  Call DatosEjecutivosWomChile
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B1").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'ActiveSheet.Range("E1").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Call DatosEjecutivosWomChile
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B2").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'ActiveSheet.Range("E1").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Call DatosEjecutivosWomChile
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'ActiveSheet.Range("E1").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BT5", Range("BT5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BW5", Range("BW5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BY5:CF5", Range("BY5:CF5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2:J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("CJ5:CQ5", Range("CJ5:CQ5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2:J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("CU5:DB5", Range("CU5:DB5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "WOM CHILE"
'Range("C2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate
'  ActiveWorkbook.Close savechanges:=False
'  Call MallaOITL
  
End Sub

Sub MallaOITL()

'Malla de turnos OIT'

  Call DatosEjecutivosOIT
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Call DatosEjecutivosOIT

  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosOIT

  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:CF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CQ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:DB5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "OIT"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaGolL
    
End Sub

Sub MallaGolL()

'Malla de turnos GOLD'

  Call DatosEjecutivosGol
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B1").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Call DatosEjecutivosGol
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B2"))           'Seleccionar columna para escribir fecha en las celdas
Range("B2").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Call DatosEjecutivosGol
 
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B3"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BT5", Range("BT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BW5", Range("BW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BY5:CF5", Range("BY5:CF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CG5:CH5", Range("CG5:CH5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CJ5:CQ5", Range("CJ5:CQ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CR5:CS5", Range("CR5:CS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2:J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("CU5:DB5", Range("CU5:DB5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "GOL"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  
End Sub


Sub DatosMartes()

'Malla de turnos LATAM'

  Call DatosEjecutivosLatam
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Range("D2").PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
ActiveCell.Value = "LATAM"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroMedellinM

End Sub
 
Sub MallaClaroMedellinM()
 
'Malla de turnos CLARO MEDELLIN'

  Call DatosEjecutivosClaroMedellin
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO MDE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
  
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroBogotaM
    
End Sub

Sub MallaClaroBogotaM()

'Malla de turnos CLARO BOGOTA'

  Call DatosEjecutivosClaroBogota
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO BOG"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroVIPM
    
End Sub

Sub MallaClaroVIPM()

'Malla de turnos CLARO VIP'

  Call DatosEjecutivosClaroVIP
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO_VIP"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaFicohsaM
    
End Sub

Sub MallaFicohsaM()

'Malla de turnos FICOHSA'

  Call DatosEjecutivosFicohsa
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "FICOHSA"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaSamsungM
    
End Sub
Sub MallaSamsungM()

'Malla de turnos SAMSUNG'

  Call DatosEjecutivosSamsung
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "SAMSUNG"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaShopeeM
    
End Sub

Sub MallaShopeeM()

'Malla de turnos SHOPEE'

  Call DatosEjecutivosShopee
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "SHOPEE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
    Workbooks(2).Activate
    ActiveWorkbook.Close savechanges:=False
    Call MallaHisenseM
    
End Sub

Sub MallaHisenseM()

'Malla de turnos HISENSE'

  Call DatosEjecutivosHisense
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "HISENSE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
    Workbooks(2).Activate
    ActiveWorkbook.Close savechanges:=False
    Call MallaFortinetM
    
End Sub

Sub MallaFortinetM()

'Malla de turnos FORTINET'

  Call DatosEjecutivosFortinet
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "FORTINET"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaConsuladoM
    
End Sub



Sub MallaConsuladoM()

'Malla de turnos CONSULADO'

  Call DatosEjecutivosConsulado
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CONSULADO"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaDinissanM
    
End Sub

Sub MallaDinissanM()

'Malla de turnos DINISSAN'

  Call DatosEjecutivosDinissan
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "DINISSAN"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
     
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaVardiM
    
End Sub

Sub MallaVardiM()

'Malla de turnos VARDI'

  Call DatosEjecutivosVardi
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "VARDI"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
   
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaWomM
    
End Sub

Sub MallaWomM()

'Malla de turnos WOM'

  Call DatosEjecutivosWom
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "WOM"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaAitelVodafoneM
  
End Sub

Sub MallaAitelVodafoneM()

'Malla de turnos AITEL'

  Call DatosEjecutivosAitelVodafone

Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues

Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "AITEL VODAFONE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaOITM
    
End Sub

Sub MallaDisponibleM()

'Malla de turnos COLMENA'

'  Call DatosEjecutivosColmena
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("T5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("W5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("Y5:AF5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "COLMENA"
'
'   Workbooks(2).Activate
'   ActiveWorkbook.Close savechanges:=False
'   Call MallaFiltrosM
    
End Sub

Sub MallaDisponible2M()

'Malla de turnos FILTROS'

'  Call DatosEjecutivosFiltros
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'ActiveSheet.Range("E1").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("T5", Range("T5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("W5", Range("W5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "FILTROS"
'Range("C2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'   Workbooks(2).Activate
'   ActiveWorkbook.Close savechanges:=False
'   Call MallaOITM
    
End Sub

Sub MallaDisponible1M()

''Malla de turnos WOM CHILE'
'
'  Call DatosEjecutivosWomChile
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'ActiveSheet.Range("E1").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("T5", Range("T5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("W5", Range("W5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "WOM CHILE"
'Range("C2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate
'  ActiveWorkbook.Close savechanges:=False
'  Call MallaOITM
  
End Sub


Sub MallaOITM()

'Malla de turnos CLARO VIP'

  Call DatosEjecutivosOIT
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "OIT"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaGolM
    
End Sub

Sub MallaGolM()

'Malla de turnos GOLD'

  Call DatosEjecutivosGol
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("T5", Range("T5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("W5", Range("W5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("Y5:AF5", Range("Y5:AF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "GOL"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  
End Sub

Sub DatosMiercoles()

'Malla de turnos LATAM'

  Call DatosEjecutivosLatam
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Range("D2").PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
ActiveCell.Value = "LATAM"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroMedellinMR

End Sub
 
Sub MallaClaroMedellinMR()
 
'Malla de turnos CLARO MEDELLIN'

  Call DatosEjecutivosClaroMedellin
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO MDE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
  
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroBogotaMR
    
End Sub

Sub MallaClaroBogotaMR()

'Malla de turnos CLARO BOGOTA'

  Call DatosEjecutivosClaroBogota
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO BOG"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroVIPMR
    
End Sub

Sub MallaClaroVIPMR()

'Malla de turnos CLARO VIP'

  Call DatosEjecutivosClaroVIP
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO_VIP"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaFicohsaMR
    
End Sub

Sub MallaFicohsaMR()

'Malla de turnos FICOHSA'

  Call DatosEjecutivosFicohsa
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "FICOHSA"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaSamsungMR
    
End Sub
Sub MallaSamsungMR()

'Malla de turnos SAMSUNG'

  Call DatosEjecutivosSamsung
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "SAMSUNG"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaShopeeMR
    
End Sub

Sub MallaShopeeMR()

'Malla de turnos SHOPEE'

  Call DatosEjecutivosShopee
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "SHOPEE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
    Workbooks(2).Activate
    ActiveWorkbook.Close savechanges:=False
    Call MallaHisenseMR
    
End Sub

Sub MallaHisenseMR()

'Malla de turnos HISENSE'

  Call DatosEjecutivosHisense
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "HISENSE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
    Workbooks(2).Activate
    ActiveWorkbook.Close savechanges:=False
    Call MallaFortinetMR
    
End Sub

Sub MallaFortinetMR()

'Malla de turnos FORTINET'

  Call DatosEjecutivosFortinet
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "FORTINET"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaConsuladoMR
    
End Sub



Sub MallaConsuladoMR()

'Malla de turnos CONSULADO'

  Call DatosEjecutivosConsulado
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CONSULADO"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaDinissanMR
    
End Sub

Sub MallaDinissanMR()

'Malla de turnos DINISSAN'

  Call DatosEjecutivosDinissan
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "DINISSAN"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
     
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaVardiMR
    
End Sub

Sub MallaVardiMR()

'Malla de turnos VARDI'

  Call DatosEjecutivosVardi
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "VARDI"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
   
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaWomMR
    
End Sub

Sub MallaWomMR()

'Malla de turnos WOM'

  Call DatosEjecutivosWom
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "WOM"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaAitelVodafoneMR
  
End Sub

Sub MallaAitelVodafoneMR()

'Malla de turnos AITEL'

  Call DatosEjecutivosAitelVodafone

Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues

Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "AITEL VODAFONE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaOITMR
    
End Sub

Sub MallaDisponibleMR()

'Malla de turnos COLMENA'

'  Call DatosEjecutivosColmena
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AG5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AJ5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AL5:AS5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "COLMENA"
'
'   Workbooks(2).Activate
'   ActiveWorkbook.Close savechanges:=False
'   Call MallaFiltrosMR
    
End Sub

Sub MallaDisponible2MR()

'Malla de turnos FILTROS'

'  Call DatosEjecutivosFiltros
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'ActiveSheet.Range("E1").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AG5", Range("AG5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AJ5", Range("AJ5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "FILTROS"
'Range("C2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'   Workbooks(2).Activate
'   ActiveWorkbook.Close savechanges:=False
'   Call MallaOITMR
    
End Sub

Sub MallaDisponible1MR()

''Malla de turnos WOM CHILE'
'
'  Call DatosEjecutivosWomChile
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'ActiveSheet.Range("E1").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AG5", Range("AG5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AJ5", Range("AJ5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "WOM CHILE"
'Range("C2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate
'  ActiveWorkbook.Close savechanges:=False
'  Call MallaOITMR
  
End Sub

Sub MallaOITMR()

'Malla de turnos OIT'

  Call DatosEjecutivosOIT
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "OIT"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaGolMR
  
End Sub

Sub MallaGolMR()

'Malla de turnos GOLD'

  Call DatosEjecutivosGol
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AG5", Range("AG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AJ5", Range("AJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AL5:AS5", Range("AL5:AS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "GOL"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  
End Sub

Sub DatosJueves()

'Malla de turnos LATAM'

  Call DatosEjecutivosLatam
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Range("D2").PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
ActiveCell.Value = "LATAM"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroMedellinV

End Sub
 
Sub MallaClaroMedellinV()
 
'Malla de turnos CLARO MEDELLIN'

  Call DatosEjecutivosClaroMedellin
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO MDE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
  
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroBogotaV
    
End Sub

Sub MallaClaroBogotaV()

'Malla de turnos CLARO BOGOTA'

  Call DatosEjecutivosClaroBogota
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO BOG"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroVIPV
    
End Sub

Sub MallaClaroVIPV()

'Malla de turnos CLARO VIP'

  Call DatosEjecutivosClaroVIP
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO_VIP"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaFicohsaV
    
End Sub

Sub MallaFicohsaV()

'Malla de turnos FICOHSA'

  Call DatosEjecutivosFicohsa
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "FICOHSA"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaSamsungV
    
End Sub
Sub MallaSamsungV()

'Malla de turnos SAMSUNG'

  Call DatosEjecutivosSamsung
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "SAMSUNG"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaShopeeV
    
End Sub

Sub MallaShopeeV()

'Malla de turnos SHOPEE'

  Call DatosEjecutivosShopee
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "SHOPEE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
    Workbooks(2).Activate
    ActiveWorkbook.Close savechanges:=False
    Call MallaHisenseV
    
End Sub

Sub MallaHisenseV()

'Malla de turnos HISENSE'

  Call DatosEjecutivosHisense
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "HISENSE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
    Workbooks(2).Activate
    ActiveWorkbook.Close savechanges:=False
    Call MallaFortinetV
    
End Sub

Sub MallaFortinetV()

'Malla de turnos FORTINET'

  Call DatosEjecutivosFortinet
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "FORTINET"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaConsuladoV
    
End Sub


Sub MallaConsuladoV()

'Malla de turnos CONSULADO'

  Call DatosEjecutivosConsulado
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CONSULADO"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaDinissanV
    
End Sub

Sub MallaDinissanV()

'Malla de turnos DINISSAN'

  Call DatosEjecutivosDinissan
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "DINISSAN"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
     
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaVardiV
    
End Sub

Sub MallaVardiV()

'Malla de turnos VARDI'

  Call DatosEjecutivosVardi
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "VARDI"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
   
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaWomV
    
End Sub

Sub MallaWomV()

'Malla de turnos WOM'

  Call DatosEjecutivosWom
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "WOM"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaAitelVodafoneV
  
End Sub

Sub MallaAitelVodafoneV()

'Malla de turnos AITEL'

  Call DatosEjecutivosAitelVodafone

Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues

Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "AITEL VODAFONE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaOITV
    
End Sub

Sub MallaDisponibleV()

'Malla de turnos COLMENA'

  'Call DatosEjecutivosColmena
  
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas

  'Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  'Visible = True
  'WindowState = xlMaximized
  'Range("AT5").Copy
  'ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  'Range("I2").Select
  'Selection.End(xlDown).Select
  'SendKeys "{DOWN}", True
  'Selection.PasteSpecial xlPasteValues
  'Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  'Visible = True
  'WindowState = xlMaximized
  'Range("AW5").Copy
  'ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  'Range("J2").Select
  'Selection.End(xlDown).Select
  'SendKeys "{DOWN}", True
  'Selection.PasteSpecial xlPasteValues
  'Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  'Visible = True
  'WindowState = xlMaximized
  'Range("AY5:BF5").Copy
  'ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  'Range("L2:S2").Select
  'Selection.End(xlDown).Select
  'SendKeys "{DOWN}", True
  'Selection.PasteSpecial xlPasteValues
 
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "COLMENA"
    
   'Workbooks(2).Activate
   'ActiveWorkbook.Close savechanges:=False
   'Call MallaFiltrosV
    
End Sub

Sub MallaDisponible2V()

'Malla de turnos FILTROS'

'  Call DatosEjecutivosFiltros
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'ActiveSheet.Range("E1").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AT5", Range("AT5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AW5", Range("AW5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "FILTROS"
'Range("C2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'   Workbooks(2).Activate
'   ActiveWorkbook.Close savechanges:=False
'   Call MallaOITV
    
End Sub

Sub MallaDisponible1V()

''Malla de turnos WOM'
'
'  Call DatosEjecutivosWomChile
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'ActiveSheet.Range("E1").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AT5", Range("AT5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AW5", Range("AW5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "WOM CHILE"
'Range("C2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate
'  ActiveWorkbook.Close savechanges:=False
'  Call MallaOITV
  
End Sub

Sub MallaOITV()

'Malla de turnos OIT'

  Call DatosEjecutivosOIT
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "OIT"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaGolV
    
End Sub

Sub MallaGolV()

'Malla de turnos GOLD'

  Call DatosEjecutivosGol
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AT5", Range("AT5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AW5", Range("AW5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("AY5:BF5", Range("AY5:BF5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "GOL"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  
End Sub

Sub DatosViernes()

'Malla de turnos LATAM'

  Call DatosEjecutivosLatam
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Range("D2").PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
ActiveCell.Value = "LATAM"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroMedellinJ

End Sub
 
Sub MallaClaroMedellinJ()
 
'Malla de turnos CLARO MEDELLIN'

  Call DatosEjecutivosClaroMedellin
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO MDE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
  
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroBogotaJ
    
End Sub

Sub MallaClaroBogotaJ()

'Malla de turnos CLARO BOGOTA'

  Call DatosEjecutivosClaroBogota
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO BOG"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaClaroVIPJ
    
End Sub

Sub MallaClaroVIPJ()

'Malla de turnos CLARO VIP'

  Call DatosEjecutivosClaroVIP
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO_VIP"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaFicohsaJ
    
End Sub

Sub MallaFicohsaJ()

'Malla de turnos FICOHSA'

  Call DatosEjecutivosFicohsa
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "FICOHSA"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaSamsungJ
    
End Sub
Sub MallaSamsungJ()

'Malla de turnos SAMSUNG'

  Call DatosEjecutivosSamsung
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "SAMSUNG"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaShopeeJ
    
End Sub

Sub MallaShopeeJ()

'Malla de turnos SHOPEE'

  Call DatosEjecutivosShopee
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "SHOPEE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
    Workbooks(2).Activate
    ActiveWorkbook.Close savechanges:=False
    Call MallaHisenseJ
    
End Sub

Sub MallaHisenseJ()

'Malla de turnos HISENSE'

  Call DatosEjecutivosHisense
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "HISENSE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
    Workbooks(2).Activate
    ActiveWorkbook.Close savechanges:=False
    Call MallaFortinetJ
    
End Sub

Sub MallaFortinetJ()

'Malla de turnos FORTINET'

  Call DatosEjecutivosFortinet
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "FORTINET"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaConsuladoJ
    
End Sub



Sub MallaConsuladoJ()

'Malla de turnos CONSULADO'

  Call DatosEjecutivosConsulado
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CONSULADO"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaDinissanJ
    
End Sub

Sub MallaDinissanJ()

'Malla de turnos DINISSAN'

  Call DatosEjecutivosDinissan
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "DINISSAN"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
     
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaVardiJ
    
End Sub

Sub MallaVardiJ()

'Malla de turnos VARDI'

  Call DatosEjecutivosVardi
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "VARDI"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
   
   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaWomJ
    
End Sub

Sub MallaWomJ()

'Malla de turnos WOM'

  Call DatosEjecutivosWom
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "WOM"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaAitelVodafoneJ
  
End Sub

Sub MallaAitelVodafoneJ()

'Malla de turnos AITEL'

  Call DatosEjecutivosAitelVodafone

Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues

Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "AITEL VODAFONE"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

   Workbooks(2).Activate
   ActiveWorkbook.Close savechanges:=False
   Call MallaOITJ
    
End Sub

Sub MallaDisponibleJ()

'Malla de turnos COLMENA'

'  Call DatosEjecutivosColmena
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BG5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BJ5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BL5:BS5").Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "COLMENA"
'
'   Workbooks(2).Activate
'   ActiveWorkbook.Close savechanges:=False
'   Call MallaFiltrosJ
    
End Sub

Sub MallaDisponible2J()

'Malla de turnos FILTROS'

'  Call DatosEjecutivosFiltros
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'ActiveSheet.Range("E1").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BG5", Range("BG5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BJ5", Range("BJ5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "FILTROS"
'Range("C2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'   Workbooks(2).Activate
'   ActiveWorkbook.Close savechanges:=False
'   Call MallaOITJ
    
End Sub

Sub MallaDisponible1J()

''Malla de turnos WOM CHILE'
'
'  Call DatosEjecutivosWomChile
'
'Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
'Range("B3").Copy
'Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'Selection.PasteSpecial Paste:=xlPasteFormulas
'ActiveSheet.Range("E1").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BG5", Range("BG5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("I2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BJ5", Range("BJ5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("J2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
'  Visible = True
'  WindowState = xlMaximized
'  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
'  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
'  Range("L2:S2").Select
'  Selection.End(xlDown).Select
'  SendKeys "{DOWN}", True
'  Selection.PasteSpecial xlPasteValues
'
'Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("D2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'Range("B2").Select
'Selection.End(xlDown).Select
'SendKeys "{DOWN}", True
'ActiveCell.Value = "WOM CHILE"
'Range("C2").Select
'Selection.End(xlDown).Select
'SendKeys "{LEFT}", True
'SendKeys "^+{UP}", True
'Selection.FillDown
'
'  Workbooks(2).Activate
'  ActiveWorkbook.Close savechanges:=False
'  Call MallaOITJ
  
End Sub

Sub MallaOITJ()

'Malla de turnos OIT'

  Call DatosEjecutivosOIT
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "OIT"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  Call MallaGolJ
  
End Sub

Sub MallaGolJ()

'Malla de turnos GOLD'

  Call DatosEjecutivosGol
  
Application.Goto (ActiveWorkbook.Sheets("Hoja2").Range("B1"))           'Seleccionar columna para escribir fecha en las celdas
Range("B3").Copy
Application.Goto (ActiveWorkbook.Sheets("DB").Range("D2"))
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial Paste:=xlPasteFormulas
ActiveSheet.Range("E1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
 
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BG5", Range("BG5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BJ5", Range("BJ5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Vuelve a activar el archivo de mallas'
  Visible = True
  WindowState = xlMaximized
  Range("BL5:BS5", Range("BL5:BS5").End(xlDown)).Copy
  ThisWorkbook.Activate                                      'Pegado en la base de mallas'
  Range("L2:S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
 
Range("D2").Select                                         'Continuidad de formulas y descripcion de operacion
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("B2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "GOL"
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
    
  Workbooks(2).Activate
  ActiveWorkbook.Close savechanges:=False
  
End Sub







































CARGA ARCHIVO DE KPIS NO VOZ

Sub CargaDeDatosKPIS()

If Format(Date, "DDDD") = "lunes" Then
Call CargaDeDatosKPISLunes
End If

If Format(Date, "DDDD") <> "lunes" Then
Call CargaDeDatosKPISMartesViernes
End If

End Sub

Sub CargaDeDatosKPISLunes()

Workbooks.Open ("\\10.96.16.27\controlgestion$\[Privado] " & Format(Date - 1, "YYYY") & "\[Privado] " & Format(Date - 1, "MM") & "-" & Format(Date - 1, "MMMM") & "\[Privado] 001-CONSOLIDADO OPERACIONES\Privado " & Format(Date - 1, "MM") & "-" & "AMC_Consolidado_NOM_Almacontact.xlsx"), , True    'Abrir archivo de Consolidado de nomina'
Visible = True
WindowState = xlMaximized
ActiveSheet.AutoFilterMode = False
Worksheets("TEM_Carga Nomina").Activate
Range("B1").AutoFilter 2, Array(Date - 3, Date - 2, Date - 1), xlFilterValues 'Filtro de fechas a copiar
Range("A1").AutoFilter 1, Array("GP1", "GP2NOVOZ", "GP2VOZ", "GP3", "GP4", "GP5"), xlFilterValues                         'Filtro de cliente LATAM
Range("F1").AutoFilter 6, Array("HVC AMC", "Latam Travel AMC", "AGENCIAS PORTUGUES", "AGENCIAS TARGET ENG", "CORPORATE PYME", "WPP EQUIPAJES AMC", "WPP EQUIPAJES AMC ING", "AGENCIAS TARGET ES", "BO_WAIVERS", "BO LUA AMC", "BO AGENCIAS TARGET", "BO ANTIFRAUDE AMC", "BO_CORPORATE", "BO RECLAMOS AMC", "BO LUA AMC ING", "CHAT AGENCIAS ESP", "DEVOLUCIONES", "DT FFP AMC", "RRSS AMC ING", "RRSS AMC", "WPP LUA AMC", "CHAT AGENCIAS ENG", "CARGO BOOKING", "CARGO CC", "BO AGENCIAS TARGET ENG", "AG CHECK IN", "BO RECLAMOS AMC ING", "BO HVC"), xlFilterValues    'Filtro de servicios a copiar
Columns("D").Select
Selection.TextToColumns DataType:=xlDelimited, _
ConsecutiveDelimiter:=True, Space:=True
  
  Range("B1").Activate                                      'Aqui comenzamos a trasladar los datos desde POC a KPIS NO VOZ
  SendKeys "{DOWN}", True
  Range(ActiveCell, ActiveCell.End(xlDown)).Copy
  ThisWorkbook.Sheets("BD_TL_No_Voz").Activate
  Range("F1").Select
  If Range("F2") = "" Then
  SendKeys ("{DOWN}"), True
  Else
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  End If
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Abrir archivo de Consolidado de nomina'
  Visible = True
  WindowState = xlMaximized
  Range("D1").Activate
  SendKeys "{DOWN}", True
  SendKeys "+{RIGHT}", True
  SendKeys "+{RIGHT}", True
  SendKeys "^+{DOWN}", True
  Selection.Copy
  ThisWorkbook.Activate
  Range("G1").Select
  If Range("G2") = "" Then
  SendKeys ("{DOWN}"), True
  Else
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  End If
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Abrir archivo de Consolidado de nomina'
  Visible = True
  WindowState = xlMaximized
  Range("J1").Activate
  SendKeys "{DOWN}", True
  SendKeys "^+{DOWN}", True
  Selection.Copy
  ThisWorkbook.Activate
  Range("J1").Select
  If Range("J2") = "" Then
  SendKeys ("{DOWN}"), True
  Else
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  End If
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Abrir archivo de Consolidado de nomina'
  Visible = True
  WindowState = xlMaximized
  Range("R1").Activate
  SendKeys "{DOWN}", True
  SendKeys "^+{DOWN}", True
  Selection.Copy
  ThisWorkbook.Activate
  Range("L1").Select
  If Range("L2") = "" Then
  SendKeys ("{DOWN}"), True
  Else
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  End If
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Abrir archivo de Consolidado de nomina'
  SendKeys "{ESCAPE}", True
  ActiveWorkbook.Close savechanges:=False
  
ThisWorkbook.Activate                                           'Damos continuidad a formatos y formulas
Range("L1").Activate
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
For i = 1 To 4
SendKeys "{RIGHT}", True
Next i
SendKeys "^+{UP}", True
For i = 1 To 75
SendKeys "+{RIGHT}", True
Next i
Selection.FillDown
Range("F2:L2").Select
Range("F2:L2").Copy
SendKeys "^+{DOWN}", True
Selection.PasteSpecial xlPasteFormats
Range("F1").Activate
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
For i = 1 To 4
SendKeys "+{LEFT}", True
Next i
SendKeys "^+{UP}", True
Selection.FillDown
ActiveWorkbook.RefreshAll
ActiveWorkbook.Close savechanges:=True

End Sub

Sub CargaDeDatosKPISMartesViernes()

Workbooks.Open ("\\10.96.16.27\controlgestion$\[Privado] " & Format(Date - 1, "YYYY") & "\[Privado] " & Format(Date - 1, "MM") & "-" & Format(Date - 1, "MMMM") & "\[Privado] 001-CONSOLIDADO OPERACIONES\Privado " & Format(Date - 1, "MM") & "-" & "AMC_Consolidado_NOM_Almacontact.xlsx"), , True    'Abrir archivo de Consolidado de nomina'
Visible = True
WindowState = xlMaximized
ActiveSheet.AutoFilterMode = False
Worksheets("TEM_Carga Nomina").Activate
Range("B1").AutoFilter 2, Array(Date - 1), xlFilterValues      'Filtro de fechas a copiar
Range("A1").AutoFilter 1, Array("GP1", "GP2NOVOZ", "GP2VOZ", "GP3", "GP4", "GP5"), xlFilterValues            'Filtro de cliente LATAM
Range("F1").AutoFilter 6, Array("HVC AMC", "Latam Travel AMC", "AGENCIAS PORTUGUES", "AGENCIAS TARGET ENG", "CORPORATE PYME", "WPP EQUIPAJES AMC", "WPP EQUIPAJES AMC ING", "AGENCIAS TARGET ES", "BO_WAIVERS", "BO LUA AMC", "BO AGENCIAS TARGET", "BO ANTIFRAUDE AMC", "BO_CORPORATE", "BO RECLAMOS AMC", "BO LUA AMC ING", "CHAT AGENCIAS ESP", "DEVOLUCIONES", "DT FFP AMC", "RRSS AMC ING", "RRSS AMC", "WPP LUA AMC", "CHAT AGENCIAS ENG", "CARGO BOOKING", "CARGO CC", "BO AGENCIAS TARGET ENG", "AG CHECK IN", "BO RECLAMOS AMC ING", "BO HVC"), xlFilterValues    'Filtro de servicios a copiar
Columns("D").Select
Selection.TextToColumns DataType:=xlDelimited, _
ConsecutiveDelimiter:=True, Space:=True
  
  Range("B1").Activate                                      'Aqui comenzamos a trasladar los datos desde POC a KPIS NO VOZ
  SendKeys "{DOWN}", True
  Range(ActiveCell, ActiveCell.End(xlDown)).Copy
  ThisWorkbook.Sheets("BD_TL_No_Voz").Activate
  Range("F1").Select
  If Range("F2") = "" Then
  SendKeys ("{DOWN}"), True
  Else
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  End If
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Abrir archivo de Consolidado de nomina'
  Visible = True
  WindowState = xlMaximized
  Range("D1").Activate
  SendKeys "{DOWN}", True
  SendKeys "+{RIGHT}", True
  SendKeys "+{RIGHT}", True
  SendKeys "^+{DOWN}", True
  Selection.Copy
  ThisWorkbook.Activate
  Range("G1").Select
  If Range("G2") = "" Then
  SendKeys ("{DOWN}"), True
  Else
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  End If
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Abrir archivo de Consolidado de nomina'
  Visible = True
  WindowState = xlMaximized
  Range("J1").Activate
  SendKeys "{DOWN}", True
  SendKeys "^+{DOWN}", True
  Selection.Copy
  ThisWorkbook.Activate
  Range("J1").Select
  If Range("J2") = "" Then
  SendKeys ("{DOWN}"), True
  Else
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  End If
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Abrir archivo de Consolidado de nomina'
  Visible = True
  WindowState = xlMaximized
  Range("R1").Activate
  SendKeys "{DOWN}", True
  SendKeys "^+{DOWN}", True
  Selection.Copy
  ThisWorkbook.Activate
  Range("L1").Select
  If Range("L2") = "" Then
  SendKeys ("{DOWN}"), True
  Else
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  End If
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate     'Abrir archivo de Consolidado de nomina'
  SendKeys "{ESCAPE}", True
  ActiveWorkbook.Close savechanges:=False
  
ThisWorkbook.Activate                                           'Damos continuidad a formatos y formulas
Range("L1").Activate
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
For i = 1 To 4
SendKeys "{RIGHT}", True
Next i
SendKeys "^+{UP}", True
For i = 1 To 75
SendKeys "+{RIGHT}", True
Next i
Selection.FillDown
Range("F2:L2").Select
Range("F2:L2").Copy
SendKeys "^+{DOWN}", True
Selection.PasteSpecial xlPasteFormats
Range("F1").Activate
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
For i = 1 To 4
SendKeys "+{LEFT}", True
Next i
SendKeys "^+{UP}", True
Selection.FillDown
ActiveWorkbook.RefreshAll
ActiveWorkbook.Close savechanges:=True

End Sub

































CONSOLIDACION DE KPIS SQL

Sub CargaKPISsql()

Range("A3").EntireRow.Activate                             'Limpieza de los datos
Range(Selection, Selection.End(xlDown)).Select
Selection.EntireRow.Delete

  Workbooks.Open ("\\10.96.16.27\reporting_almacontact\2022\02.Privado\01. [Privado] Analistas\02. [Privado] Julian Cardona\[Privado] Carpetas diarias AMC\[Privado] 11-Archivos KPIS\01-Plantillas\" & Format(Date - 1, "YYYY") & "\" & Format(Date - 1, "MM") & "-" & Format(Date - 1, "MMMM") & "\" & Format(Date - 1, "MMMM") & "_" & Format(Date - 1, "YYYY") & " Genesys.xlsx"), True 'Abre archivo de kpis GENESYS
  Visible = True
  WindowState = xlMaximized
  If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
  Range("I5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  
ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("C2").PasteSpecial xlPasteValues

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("L5:T5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  
ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("D2").PasteSpecial xlPasteValues

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("AJ5:AL5").Select
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("M2").PasteSpecial xlPasteValues

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("BG5:BO5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  
ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("P2").PasteSpecial xlPasteValues
  
  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("AY5").Select
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("Y2").PasteSpecial xlPasteValues

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("BR5:BS5").Select
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("AA2").PasteSpecial xlPasteValues

  Range("Z2").Activate                                       'Nombre de la base que se estrajeron los datos
  ActiveCell.Value = "GENESYS"
  Range("Y2").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown
  
Workbooks(2).Activate                                        'Cierre archivo de KPIS GENESYS
ActiveWorkbook.Close savechanges:=False

  Workbooks.Open ("\\10.96.16.27\reporting_almacontact\2022\02.Privado\01. [Privado] Analistas\02. [Privado] Julian Cardona\[Privado] Carpetas diarias AMC\[Privado] 11-Archivos KPIS\01-Plantillas\" & Format(Date - 1, "YYYY") & "\" & Format(Date - 1, "MM") & "-" & Format(Date - 1, "MMMM") & "\" & Format(Date - 1, "MMMM") & "_" & Format(Date - 1, "YYYY") & " CMS.xlsx"), True 'Abre archivo de kpis CMS
  Visible = True
  WindowState = xlMaximized
  If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
  Range("I5").Select
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("K5:S5").Select
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("AI5:AK5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  
ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("M2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("BF5:BN5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  
ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("P2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("AX5").Select
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("Y2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues

  Range("Z1").Select                                         'Nombre de la base que se estrajeron los datos
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell.Value = "CMS"
  Range("Y1").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown

Range("AA1").Select                                          'Llenado de campos con 0 para KPIS CMS
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "0"
Range("AB1").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "0"
Range("Z1").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "+{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate                                        'Cierre archivo de KPIS CMS
  ActiveWorkbook.Close savechanges:=False
  
Workbooks.Open ("\\10.96.16.27\reporting_almacontact\2022\02.Privado\01. [Privado] Analistas\02. [Privado] Julian Cardona\[Privado] Carpetas diarias AMC\[Privado] 11-Archivos KPIS\01-Plantillas\" & Format(Date - 1, "YYYY") & "\" & Format(Date - 1, "MM") & "-" & Format(Date - 1, "MMMM") & "\" & "Conexiones Aplicativo Click " & Format(Date - 1, "MMMM") & "_" & Format(Date - 1, "YYYY") & "_V3.xlsm"), True  'Abre archivo de kpis CLICK
Visible = True
WindowState = xlMaximized
If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
Range("L1").AutoFilter 12, ">0:00"
Range("F2:G2").Select
Range(Selection, Selection.End(xlDown)).Copy

  ThisWorkbook.Activate                                        'Pegamos los datos en la base
  Range("C2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues

Range("E1").Select                                          'Llenado de campos con 0 para KPIS CLICK
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "0"
Range("F1").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "0"
Range("G1").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "0"
Range("H1").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "0"
Range("D1").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "+{RIGHT}", True
SendKeys "+{RIGHT}", True
SendKeys "+{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("AF2").Select
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("I2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Range("J1").Select                                           'Llenado de campos con 0 para KPIS CLICK
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "0"
Range("I1").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("BO2").Select
  Range(Selection, Selection.End(xlDown)).Copy
  
ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("K2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Range("L1").Select                                          'Llenado de campos con 0 para KPIS CLICK
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "0"
Range("M1").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "0"
Range("N1").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "0"
Range("O1").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "0"
Range("K1").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "+{RIGHT}", True
SendKeys "+{RIGHT}", True
SendKeys "+{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("BD2").Select
  Range(Selection, Selection.End(xlDown)).Copy
  
ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("P2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("BL2").Select
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("Q2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("BE2").Select
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("R2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("BI2").Select
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("S2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("BG2").Select
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("T2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("BF2").Select
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("U2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues

  Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("BH2").Select
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("V2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues

Workbooks(2).Activate                                      'Activamos de nuevo el archivo original de KPIS
  Range("BJ2").Select
  Range(Selection, Selection.End(xlDown)).Copy

ThisWorkbook.Activate                                        'Pegamos los datos en la base
Range("W2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues

  Range("X1").Select                                           'Llenado de campos con 0 para KPIS CLICK
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell.Value = "0"
  Range("W1").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown
  
Range("Y1").Select                                           'Llenado de campos con 0 para KPIS CLICK
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "OK"
Range("X1").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("Z1").Select                                           'Llenado de campos con 0 para KPIS CLICK
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "CLICK"
Range("Y1").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("AA1").Select                                           'Llenado de campos con 0 para KPIS CLICK
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "0"
Range("Z1").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("AB1").Select                                           'Llenado de campos con 0 para KPIS CLICK
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
ActiveCell.Value = "0"
Range("AA1").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown
  
  Workbooks(2).Activate                                        'Cierre archivo de KPIS CMS
  ActiveWorkbook.Close savechanges:=False

ThisWorkbook.Activate                                          'Damos continuidad a formatos y formulas
Range("C2:AB2").Copy
Range("C2:AB2").Select
Range(Selection, Selection.End(xlDown)).PasteSpecial xlPasteFormats
Range("C1").Select
Selection.End(xlDown).Select
SendKeys "{LEFT}", True
SendKeys "+{LEFT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("A1").Select

  ThisWorkbook.Activate                                        'Cierre la plantilla para SQL guardando cambios
  ActiveWorkbook.Close savechanges:=True

End Sub













ENVIO CORREO KPIS SQL

Sub EnviarCorreoBaseKPISsql()

Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")
Dim olMail As Outlook.MailItem
Set olMail = olApp.CreateItem(olMailItem)
Dim RutaImagen As String

With olMail
.Display
.To = "daguarin@almacontactcol.co"
.CC = "mramos@almacontactcol.co"
.Subject = "[Privado] Actualización base de KPIS SQL"
.HTMLBody = "<H6> <p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Buenos d&iacute;as,</span></span></em></strong></p></H6></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'><span style='color:#002060;'>Base de KPIS SQL.</span></span></em></strong></p>" & _
            "<H6><p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Actualizada:</span><span style='color:#EE1750;'>&nbsp;</span><span style='color:#C00000;'>" & Format(Date - 1, "DD/MM/YYYY") & "</span></span></em></strong></p></H6></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>En caso de tener novedades por favor contactar el analista encargado para su validación.</span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'><span style='color:#002060;'>Agradecemos tú ayuda calificando nuestra calidad en el siguiente enlace:  <a href='https://forms.office.com/r/4scK4ZKD1G'>https://forms.office.com/r/4scK4ZKD1G</a></span></span></em></strong></p></p></ul>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style='color:#002060;'></span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Saludos.</span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>¡Que cada día nos haga mejores personas!</span></span></em></strong><font size=3>&#129299<\font></p>" & .HTMLBody
            
.Display
.Send
End With

End Sub












ENVIO DE CORREO PRODCUTIVIDAD NO VOZ SQL

Sub EnviarCorreoProductSQL()

Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")
Dim olMail As Outlook.MailItem
Set olMail = olApp.CreateItem(olMailItem)
Dim RutaImagen As String

With olMail
.Display
.To = "daguarin@almacontactcol.co"
.CC = "mramos@almacontactcol.co"
.Subject = "[Privado] Actualización base de Productividad No Voz SQL"
.HTMLBody = "<H6> <p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Buenos d&iacute;as,</span></span></em></strong></p></H6></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'><span style='color:#002060;'>Base de Productividad No Voz SQL.</span></span></em></strong></p>" & _
            "<H6><p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Actualizada:</span><span style='color:#EE1750;'>&nbsp;</span><span style='color:#C00000;'>" & Format(Date - 1, "DD/MM/YYYY") & "</span></span></em></strong></p></H6></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>En caso de tener novedades por favor contactar el analista encargado para su validación.</span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'><span style='color:#002060;'>Agradecemos tú ayuda calificando nuestra calidad en el siguiente enlace:  <a href='https://forms.office.com/r/4scK4ZKD1G'>https://forms.office.com/r/4scK4ZKD1G</a></span></span></em></strong></p></p></ul>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style='color:#002060;'></span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Saludos.</span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>¡Que cada día nos haga mejores personas!</span></span></em></strong><font size=3>&#129299<\font></p>" & .HTMLBody
            
.Display
.Send
End With

End Sub











ENVIO CORREO PRODUCTIVIDAD POWER BI

Sub EnviarCorreoProductNoVoz()

Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")
Dim olMail As Outlook.MailItem
Set olMail = olApp.CreateItem(olMailItem)
Dim RutaImagen As String

With olMail
.Display
.To = "jgutierrez@almacontactcol.co; ibarrera@almacontactcol.co; daguarin@almacontactcol.co"
.CC = "mramos@almacontactcol.co"
.Subject = "[Privado] Actualización Informe Productividad No Voz"
.HTMLBody = "<H6> <p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Buenos d&iacute;as,</span></span></em></strong></p></H6></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'><span style='color:#002060;'>Informe de Productividad No Voz.</span></span></em></strong></p>" & _
            "<H6><p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Actualizado:</span><span style='color:#EE1750;'>&nbsp;</span><span style='color:#C00000;'>" & Format(Date - 1, "DD/MM/YYYY") & "</span></span></em></strong></p></H6></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Link de acceso: <a href='https://app.powerbi.com/view?r=eyJrIjoiMTU5NTMwNjctYWM5Mi00YjI2LTgzNjYtM2RkNTg0M2E5ZDE4IiwidCI6ImE1M2VkNGQ4LTAyODMtNDMxMy1iYjIwLWRjMzUwNDI4ZDg1OSIsImMiOjR9'>Informe de Productividad Power BI</a></span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>En el anterior link encontraran el informe de productividades relacionadas a continuación, cualquier novedad por favor validar con el analista a cargo: </span></span></em></strong></p>" & _
            "<ul style='margin-bottom:0cm;margin-top:0cm;;color:#002060;' type='disc'>" & _
            "<li style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'><span style='color:#002060;'>BACK OFFICE &#10003</span></span></em></strong></p></li></ul>" & _
            "<ul style='margin-bottom:0cm;margin-top:0cm;;color:#002060;' type='disc'>" & _
            "<li style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'><span style='color:#002060;'>BACK OFFICE CUS &#10003</span></span></em></strong></p></li></ul>" & _
            "<ul style='margin-bottom:0cm;margin-top:0cm;;color:#002060;' type='disc'>" & _
            "<li style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'><span style='color:#002060;'>DEVOLUCIONES &#10003</span></span></em></strong></p></li></ul>" & _
            "<ul style='margin-bottom:0cm;margin-top:0cm;;color:#002060;' type='disc'>" & _
            "<li style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'><span style='color:#002060;'>DREAM TEAM &#10003</span></span></em></strong></p></li></ul>" & _
            "<ul style='margin-bottom:0cm;margin-top:0cm;;color:#002060;' type='disc'>" & _
            "<li style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'><span style='color:#002060;'>SOCIAL MEDIA &#10003</span></span></em></strong></p></li></ul>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'><span style='color:#002060;'></span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;color:red;'>&nbsp;</span></em></strong></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'><span style='color:#002060;'>Agradecemos tú ayuda calificando nuestra calidad en el siguiente enlace:  <a href='https://forms.office.com/r/4scK4ZKD1G'>https://forms.office.com/r/4scK4ZKD1G</a></span></span></em></strong></p></p></ul>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style='color:#002060;'></span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Saludos.</span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>¡Que cada día nos haga mejores personas!</span></span></em></strong><font size=3>&#129299<\font></p>" & .HTMLBody
            
.Display
.Send
End With

End Sub







































CARGA DE DATOS HORAS FACTURABLES

Sub ActualizacionHorasFacturacion()

Worksheets("BD").Activate           'Limpieza de datos
Range("A2:G2").Select
Range(Selection, Selection.End(xlDown)).ClearContents

  Workbooks.Open Worksheets("BD").Range("N2"), True          'Actualizacion de los datos desde el archivo de kpis
  Worksheets("BD_TL_No_Voz").Activate
  If ActiveSheet.FilterMode = True Then
  ActiveSheet.ShowAllData
  End If
  Range("F2:K2").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  Worksheets("BD").Activate
  Range("A2:F2").PasteSpecial xlPasteValues
  Workbooks(2).Activate
  Range("P2").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  Range("G2").PasteSpecial xlPasteValues
  Worksheets("Horas").Activate
  ActiveWorkbook.RefreshAll
  Workbooks(2).Activate
  SendKeys "{ESCAPE}", True
  ActiveWorkbook.Close Savechanges:=False
  ThisWorkbook.Activate
  ActiveWorkbook.Save
  
Call EnviarCorreo
 
  ThisWorkbook.Close Savechanges = False
  
End Sub





















ENVIO CORREO HORAS FACTURABLES


Sub EnviarCorreo()

Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")
Dim olMail As Outlook.MailItem
Set olMail = olApp.CreateItem(olMailItem)
Const RutaImagen As String = "\\10.96.16.27\reporting_almacontact\2022\02.Privado\01. [Privado] Analistas\02. [Privado] Julian Cardona\[Privado] Carpetas diarias AMC\[Privado] 06-Horas facturacion\Imagen Envio\Tabla.jpg"

Range("C12:D12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
ActiveSheet.ChartObjects.Add(0, 0, Selection.Width, Selection.Height).Select
ActiveChart.Paste
ActiveChart.Export Filename:=RutaImagen
ActiveChart.Parent.Delete

With olMail
.Display
.To = "jamejia@almacontactcol.co"
.CC = "mramos@almacontactcol.co; ictabares@almacontactcol.co; ibarrera@almacontactcol.co; afurrego@almacontactcol.co; jamendoza@almacontactcol.co; oroldan@almacontactcol.co; jgutierrez@almacontactcol.co; japerez@almacontactcol.co; dfalzate@almacontactcol.co; mgjacobo@almacontactcol.co; crmartinez@almacontactcol.co; dmoura@almacontactcol.co"
.Subject = "[Privado] HORAS PARA FACTURACION"
.HTMLBody = "<Body><pstyle='margin:0cm;font-size:15px;font-family:Calibri,sans-serif;'</p><em><p><span style=color:#002060;>Buenos días,</span></p></em></Body>" & _
            "<Body><pstyle='margin:0cm;font-size:15px;font-family:Calibri,sans-serif;'</p><em><p><span style=color:#002060;>Comparto las horas productivas para el mes de " & Application.WorksheetFunction.Proper(Format(Date - 1, "Mmmm")) & "." & "</span></p></em></Body>" & _
            "<Body><pstyle='font-size:15px;font-family:Calibri,sans-serif;'</p><em><p><span style=color:red;>Corte: </span>" & "<span style=color:#002060;>" & Format(Date - 1, "DD/MM/YYYY") & "</span></p></Body>" & _
            "<img src='\\10.96.16.27\reporting_almacontact\2022\02.Privado\01. [Privado] Analistas\02. [Privado] Julian Cardona\[Privado] Carpetas diarias AMC\[Privado] 06-Horas facturacion\Imagen Envio\Tabla.jpg'height=414 width=203>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'><span style='color:#002060;'>Agradecemos tú ayuda calificando nuestra calidad en el siguiente enlace:  <a href='https://forms.office.com/r/4scK4ZKD1G'>https://forms.office.com/r/4scK4ZKD1G</a></span></span></em></strong></p></p></ul>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style='color:#002060;'></span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Saludos.</span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>¡Que cada día nos haga mejores personas!</span></span></em></strong><font size=3>&#129299<\font></p>" & .HTMLBody
            
.Attachments.Add ActiveWorkbook.FullName
.Display
.Send
End With

End Sub
























































CARGA DE DATOS BASE DE NOMINA PROYECCION FINANCIERA

Sub CargaBaseNomina()

ThisWorkbook.Activate                                       'Limpieza de los datos
Worksheets("Consolidado").Activate
Range("A3:AW3").Select
Range(Selection, Selection.End(xlDown)).ClearContents

  Workbooks.Open Worksheets("Link VBA").Range("C2"), True   'Apertura del libro consolidado SQL y copiado de los datos
  Worksheets("Consolidado").Activate
  Range("B1").AutoFilter 2, Array(Date - 1), xlFilterValues
  Range("A1").Select
  SendKeys "{DOWN}", True
  Selection.EntireRow.Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.EntireRow.Delete
  ActiveSheet.AutoFilterMode = False
  Range("A2:AW2").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  Range("A3:AW3").PasteSpecial xlPasteValues
  Range("AX3").Select
  Selection.End(xlDown).Select
  SendKeys "{LEFT}", True
  Selection.EntireRow.Copy
  Selection.EntireRow.Select
  Range(Selection, Selection.End(xlDown)).PasteSpecial xlPasteFormats
  Range("AW3").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  For i = 1 To 13
  SendKeys "+{RIGHT}", True
  Next i
  SendKeys "^+{UP}", True
  Selection.FillDown
  Workbooks(2).Activate
  Range("BL2").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  Range("BL3").PasteSpecial xlPasteValues
  Columns("D").Select
  Selection.TextToColumns DataType:=xlDelimited, ConsecutiveDelimiter:=True, Space:=True
  Workbooks(2).Activate
  SendKeys "{ESCAPE}", True
  ActiveWorkbook.Close savechanges:=False

Workbooks.Open ("\\Co0000fs0001\planeacion$\01_LATAM\01_WFM\01- CARGAS BD\7- SOCIODEMOGRAFICO\SOCIO SQL AMC.xlsx"), , True                  'Cruce con la informacion en sociodemografico
Worksheets("CONSOLIDADO").Activate
Range("BD1").Activate
ActiveCell.Value = "Validador"
Range("BD2").FormulaLocal = "=BUSCARV(B2;'[03-NOMINA ALMACONTACT-Marzo.xlsm]Consolidado'!$D:$D;1;0)"
Range("BC1").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown
Range("A2").AutoFilter
Range("BD1").AutoFilter 56, "#N/D", xlFilterValues
Range("G1").AutoFilter 7, "Activo", xlFilterValues
Range("H1").AutoFilter 8, "*ASESOR*"

  Range("A1").Select                                            'Copiado de datos personal no encontrado en la base de Nomina
  SendKeys "{DOWN}", True
  SendKeys "+{RIGHT}", True
  SendKeys "+{RIGHT}", True
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  Range("C2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate
  Range("I1").Select
  SendKeys "{DOWN}", True
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  Range("F2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate
  Range("G1").Select
  SendKeys "{DOWN}", True
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  Range("G2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate
  Range("BC1").Select
  SendKeys "{DOWN}", True
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  Range("A2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Range("BL2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Workbooks(2).Activate
  Range("S1").Select
  SendKeys "{DOWN}", True
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Range("B2").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  Selection.EntireRow.Copy
  Selection.EntireRow.Select
  Range(Selection, Selection.End(xlDown)).PasteSpecial xlPasteFormats
  Range("A2").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown
  Range("H2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "No Prog"
  SendKeys "{LEFT}", True
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown
  Range("I2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  SendKeys "{LEFT}", True
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown
  Range("K2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("L2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("M2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("N2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("O2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("R2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("S2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("T2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("U2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "OK"
  Range("X2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AB2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  SendKeys "{LEFT}", True
  SendKeys "{LEFT}", True
  ActiveCell = "OK"
  SendKeys "{RIGHT}", True
  ActiveCell = "R_J"
  SendKeys "{RIGHT}", True
  ActiveCell = "0"
  SendKeys "{RIGHT}", True
  ActiveCell = "0"
  Range("AF2").Select
  Selection.End(xlDown).Select
  Selection.Copy
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Range("AG2").Select
  Selection.End(xlDown).Select
  Selection.Copy
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Range("AH2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AI2").Select
  Selection.End(xlDown).Select
  Selection.Copy
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Range("AJ2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AK2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AL2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AM2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AN2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AO2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AP2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AQ2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AR2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AS2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AT2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AU2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AV2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AW2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  ActiveCell = "0"
  Range("AX2").Select
  Selection.End(xlDown).Select
  For i = 1 To 13
  SendKeys "+{RIGHT}", True
  Next i
  Selection.Copy
  SendKeys "{Down}", True
  ActiveCell.PasteSpecial xlPasteAll
  Range("J2").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "^+{RIGHT}", True
  SendKeys "+{LEFT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown

Range("BL2").Select                                   'Alineacion de operaciones para la base
Selection.AutoFilter 64, "*ALLIANCE* "
SendKeys "{DOWN}", True
ActiveCell.Value = "ALLIANCE"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown
Range("BL2").Select
Selection.AutoFilter 64, Array("CLARO PREMIUM", "CLARO SWAT", "CLARO_VIP"), xlFilterValues
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO_VIP"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown
Range("BL2").Select
Selection.AutoFilter 64, Array("CLARO TMK BOG", "CLARO TMK MED", "CLARO"), xlFilterValues
SendKeys "{DOWN}", True
ActiveCell.Value = "CLARO"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown
Range("BL2").Select
Selection.AutoFilter 64, "*CONSULADO*"
SendKeys "{DOWN}", True
ActiveCell.Value = "CONSULADO"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown
Range("BL2").Select
Selection.AutoFilter 64, "*DINISSAN*"
SendKeys "{DOWN}", True
ActiveCell.Value = "DINISSAN"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown
Range("BL2").Select
Selection.AutoFilter 64, "*FICOHSA*"
SendKeys "{DOWN}", True
ActiveCell.Value = "FICOHSA"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown
Range("BL2").Select
Selection.AutoFilter 64, "*FILTROS*"
SendKeys "{DOWN}", True
ActiveCell.Value = "FILTROS"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown
Range("BL2").Select
Selection.AutoFilter 64, "*FORTINET*"
SendKeys "{DOWN}", True
ActiveCell.Value = "FORTINET"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown
Range("BL2").Select
Selection.AutoFilter 64, "*HISENSE*"
SendKeys "{DOWN}", True
ActiveCell.Value = "HISENSE"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown
Range("BL2").Select
Selection.AutoFilter 64, "*LATAM*"
SendKeys "{DOWN}", True
ActiveCell.Value = "LATAM"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown
Range("BL2").Select
Selection.AutoFilter 64, "*SAMSUNG*"
SendKeys "{DOWN}", True
ActiveCell.Value = "SAMSUNG"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown
Range("BL2").Select
Selection.AutoFilter 64, "*SHOPEE*"
SendKeys "{DOWN}", True
ActiveCell.Value = "SHOPEE"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown
Range("BL2").Select
Selection.AutoFilter 64, "*VARDI*"
SendKeys "{DOWN}", True
ActiveCell.Value = "VARDI"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown
Range("BL2").Select
Selection.AutoFilter 64, "WOM", xlFilterValues
SendKeys "{DOWN}", True
ActiveCell.Value = "WOM"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown
Range("BL2").Select
Selection.AutoFilter 64, "WOM CHILE", xlFilterValues
SendKeys "{DOWN}", True
ActiveCell.Value = "WOM CHILE"
Range(Selection, Selection.End(xlDown)).Select
Selection.FillDown

  If ActiveSheet.FilterMode = True Then               'Limpieza de los filtros
  ActiveSheet.ShowAllData
  End If

Workbooks(2).Activate                               'Cierre del sociodemografico
ActiveWorkbook.Close savechanges:=False

  Range("Z2").AutoFilter 26, ""                     'Correccion de justificaciones en blanco POC
  Range("U2").AutoFilter 21, "OK", xlFilterValues
  Range("Z2").Select
  SendKeys "{DOWN}", True
  ActiveCell.Value = Worksheets("imputs").Range("P2").Value
  Range("X2").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "{RIGHT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown
  Range("AA2").Select
  SendKeys "{DOWN}", True
  ActiveCell.Value = Worksheets("imputs").Range("R2").Value
  Range("AB2").Select
  Selection.End(xlDown).Select
  SendKeys "{LEFT}", True
  SendKeys "^+{UP}", True
  SendKeys "+{DOWN}", True
  SendKeys "+{DOWN}", True
  Selection.FillDown
  Range("U2").AutoFilter 21, " Excede ", xlFilterValues
  Range("W2").AutoFilter 23, "<=00:20", xlFilterValues
  Range("Z2").Select
  SendKeys "{DOWN}", True
  ActiveCell.Value = Worksheets("imputs").Range("P2").Value
  Range("X2").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "{RIGHT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown
  Range("AA2").Select
  SendKeys "{DOWN}", True
  ActiveCell.Value = Worksheets("imputs").Range("R2").Value
  Range("AB2").Select
  Selection.End(xlDown).Select
  SendKeys "{LEFT}", True
  SendKeys "^+{UP}", True
  SendKeys "+{DOWN}", True
  SendKeys "+{DOWN}", True
  Selection.FillDown
  Range("W2").AutoFilter 23, ">00:20", xlFilterValues
  Range("Z2").Select
  SendKeys "{DOWN}", True
  ActiveCell.Value = Worksheets("imputs").Range("P15").Value
  Range("X2").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "{RIGHT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown
  Range("AA2").Select
  SendKeys "{DOWN}", True
  ActiveCell.Value = Worksheets("imputs").Range("R15").Value
  Range("AB2").Select
  Selection.End(xlDown).Select
  SendKeys "{LEFT}", True
  SendKeys "^+{UP}", True
  SendKeys "+{DOWN}", True
  SendKeys "+{DOWN}", True
  Selection.FillDown
  If ActiveSheet.FilterMode = True Then               'Limpieza de los filtros
  ActiveSheet.ShowAllData
  End If
  Range("Z2").AutoFilter 26, ""
  Range("U2").AutoFilter 21, " Falta", xlFilterValues
  Range("V2").AutoFilter 22, "<=00:10", xlFilterValues
  Range("Z2").Select
  SendKeys "{DOWN}", True
  ActiveCell.Value = Worksheets("imputs").Range("P2").Value
  Range("X2").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "{RIGHT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown
  Range("AA2").Select
  SendKeys "{DOWN}", True
  ActiveCell.Value = Worksheets("imputs").Range("R2").Value
  Range("AB2").Select
  Selection.End(xlDown).Select
  SendKeys "{LEFT}", True
  SendKeys "^+{UP}", True
  SendKeys "+{DOWN}", True
  SendKeys "+{DOWN}", True
  Selection.FillDown
  Range("V2").AutoFilter 22, ">=04:00", xlFilterValues
  Range("Z2").Select
  SendKeys "{DOWN}", True
  ActiveCell.Value = Worksheets("imputs").Range("P5").Value
  Range("X2").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "{RIGHT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown
  Range("AA2").Select
  SendKeys "{DOWN}", True
  ActiveCell.Value = Worksheets("imputs").Range("R5").Value
  Range("AB2").Select
  Selection.End(xlDown).Select
  SendKeys "{LEFT}", True
  SendKeys "^+{UP}", True
  SendKeys "+{DOWN}", True
  SendKeys "+{DOWN}", True
  Selection.FillDown
  If ActiveSheet.FilterMode = True Then               'Limpieza de los filtros
  ActiveSheet.ShowAllData
  End If
  Range("Z2").AutoFilter 26, ""
  Range("Z2").Select
  SendKeys "{DOWN}", True
  ActiveCell.Value = Worksheets("imputs").Range("P14").Value
  Range("X2").Select
  Selection.End(xlDown).Select
  SendKeys "{RIGHT}", True
  SendKeys "{RIGHT}", True
  SendKeys "^+{UP}", True
  Selection.FillDown
  Range("AA2").Select
  SendKeys "{DOWN}", True
  ActiveCell.Value = Worksheets("imputs").Range("R14").Value
  Range("AB2").Select
  Selection.End(xlDown).Select
  SendKeys "{LEFT}", True
  SendKeys "^+{UP}", True
  SendKeys "+{DOWN}", True
  SendKeys "+{DOWN}", True
  Selection.FillDown
  If ActiveSheet.FilterMode = True Then               'Limpieza de los filtros
  ActiveSheet.ShowAllData
  End If
  
Range("A2").Select
ThisWorkbook.Close savechanges:=True

  
End Sub






















































ENVIO NOTIFICACION BASE DE PRIYECCION FINANCIERA

Sub EnviarNotificacion()

Dim olApp As Outlook.Application
Set olApp = CreateObject("Outlook.Application")
Dim olMail As Outlook.MailItem
Set olMail = olApp.CreateItem(olMailItem)
Dim RutaImagen As String

With olMail
.Display
.To = "dmoura@almacontactcol.co; jarubio@almacontactcol.co"
.CC = "mramos@almacontactcol.co"
.Subject = "[Privado] Proyección Financiera"
.HTMLBody = "<H6> <p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Buenos d&iacute;as,</span></span></em></strong></p></H6></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'><span style='color:#002060;'>Actualización base de nomina Share point.</span></span></em></strong></p>" & _
            "<H6><p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Actualizado:</span><span style='color:#EE1750;'>&nbsp;</span><span style='color:#C00000;'>" & Format(Date - 2, "DD/MM/YYYY") & "</span></span></em></strong></p></H6></Body>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Comparto base de nomina para proyeccion financiera, la misma se toma de acuerdo a las mallas de cada operacion y se puede encontrar en el link agregado.</span></span></em></strong></p>" & _
            "<ul style='margin-bottom:0cm;margin-top:0cm;;color:#002060;' type='disc'>" & _
            "<li style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:12px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <span style='color:#002060;'>Link para acceso: <a href='https://almacontact-my.sharepoint.com/:f:/g/personal/jcardona_almacontactcol_co/ElPC7CTAi4tMvxAVU6jQPzAB6DASz1NQwtEyWCn3VIu12A?e=zqTQz9'>Base de Nomina</a></span></span></em></strong></p></li></ul>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:16px;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style='color:#002060;'></span></span></em></strong></p>" & _
            "<p style='margin:0cm;font-size:15px;font-family:'Calibri',sans-serif;'><strong><em><span style='font-size:13px;'><span style='color:#002060;'>Saludos.</span></span></em></strong></p>" & .HTMLBody

.Display
.Send
End With

End Sub












CARGA DE DATOS BASE DE MALLAS SP SQL

Sub CargaPrincipal()

Dim Respuesta As Integer
Respuesta = MsgBox("¿Ejecutar semana completa?", vbYesNo + vbQuestion)

Worksheets("DATA").Activate                  'Limpieza de datos
Range("A4").EntireRow.Activate
Range(Selection, Selection.End(xlDown)).Select
Selection.EntireRow.Delete
Range("A3:D3").ClearContents

  Select Case Respuesta
  Case vbYes
  Call RespuestaYes
  Case vbNo
  Call RespuestaNo
  End Select

  Range("A3").EntireRow.Copy                                'Continuidad de formatos
  Range("A3").EntireRow.Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.PasteSpecial xlPasteFormats

Range("A2").Select                                          'Continuidad de formulas
Selection.End(xlDown).Select
For i = 1 To 4
SendKeys "{RIGHT}", True
Next i
For i = 1 To 7
SendKeys "+{RIGHT}", True
Next i
SendKeys "^+{UP}", True
Selection.FillDown
  
  Sheets("CARGA").Activate                             'Limpia los datos en la hoja de carga
  Range("A2").EntireRow.Activate
  Range(Selection, Selection.End(xlDown)).Select
  Selection.EntireRow.Delete
  
Sheets("DATA").Activate                                'Copia de datos a hoja de carga
Range("F3:L3").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Sheets("CARGA").Activate
Range("A2").PasteSpecial xlPasteValues
Range("A2").PasteSpecial xlPasteFormats
  
  MsgBox ("¡CONSOLIDACIÓN FINALIZADA CON EXITO!")
  
End Sub

Sub CodigoEjecucionLunes()

Workbooks(2).Sheets("Turnos").Activate
Range("L5").Select
Range(Selection, Selection.End(xlDown)).Copy
ThisWorkbook.Activate

  If Range("A3") = "" Then                'Pega las avayas o usuarios
  Range("A2").Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Else
  Range("A2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  End If

Workbooks(2).Sheets("Turnos").Activate                    'Pegas las fechas de los lunes
Range("W3").Copy
ThisWorkbook.Activate

  If Range("B3") = "" Then
  Range("B2").Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Else
  Range("B2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  End If

Range("A2").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Sheets("Turnos").Activate                     'Pegas los horarios de inicio de los lunes
  Range("Y5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate

If Range("C3") = "" Then
Range("C2").Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Else
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
End If
  
  Workbooks(2).Sheets("Turnos").Activate                     'Pegas los horarios de fin de los lunes
  Range("Z5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  
If Range("D3") = "" Then
Range("D2").Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Else
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
End If


End Sub

Sub CodigoEjecucionMartes()

Workbooks(2).Sheets("Turnos").Activate
Range("L5").Select
Range(Selection, Selection.End(xlDown)).Copy
ThisWorkbook.Activate

  If Range("A3") = "" Then                'Pega las avayas o usuarios
  Range("A2").Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Else
  Range("A2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  End If

Workbooks(2).Sheets("Turnos").Activate                    'Pegas las fechas de los martes
Range("AJ3").Copy
ThisWorkbook.Activate

  If Range("B3") = "" Then
  Range("B2").Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Else
  Range("B2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  End If

Range("A2").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Sheets("Turnos").Activate                     'Pegas los horarios de inicio de los martes
  Range("AL5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate

If Range("C3") = "" Then
Range("C2").Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Else
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
End If
  
  Workbooks(2).Sheets("Turnos").Activate                     'Pegas los horarios de fin de los martes
  Range("AM5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  
If Range("D3") = "" Then
Range("D2").Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Else
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
End If


End Sub

Sub CodigoEjecucionMiercoles()

Workbooks(2).Sheets("Turnos").Activate
Range("L5").Select
Range(Selection, Selection.End(xlDown)).Copy
ThisWorkbook.Activate

  If Range("A3") = "" Then                'Pega las avayas o usuarios
  Range("A2").Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Else
  Range("A2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  End If

Workbooks(2).Sheets("Turnos").Activate                    'Pegas las fechas de los miercoles
Range("AW3").Copy
ThisWorkbook.Activate

  If Range("B3") = "" Then
  Range("B2").Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Else
  Range("B2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  End If

Range("A2").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Sheets("Turnos").Activate                     'Pegas los horarios de inicio de los miercoles
  Range("AY5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate

If Range("C3") = "" Then
Range("C2").Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Else
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
End If
  
  Workbooks(2).Sheets("Turnos").Activate                     'Pegas los horarios de fin de los miercoles
  Range("AZ5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  
If Range("D3") = "" Then
Range("D2").Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Else
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
End If


End Sub

Sub CodigoEjecucionJueves()

Workbooks(2).Sheets("Turnos").Activate
Range("L5").Select
Range(Selection, Selection.End(xlDown)).Copy
ThisWorkbook.Activate

  If Range("A3") = "" Then                'Pega las avayas o usuarios
  Range("A2").Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Else
  Range("A2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  End If

Workbooks(2).Sheets("Turnos").Activate                    'Pegas las fechas de los jueves
Range("BJ3").Copy
ThisWorkbook.Activate

  If Range("B3") = "" Then
  Range("B2").Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Else
  Range("B2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  End If

Range("A2").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Sheets("Turnos").Activate                     'Pegas los horarios de inicio de los jueves
  Range("BL5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate

If Range("C3") = "" Then
Range("C2").Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Else
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
End If
  
  Workbooks(2).Sheets("Turnos").Activate                     'Pegas los horarios de fin de los jueves
  Range("BM5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  
If Range("D3") = "" Then
Range("D2").Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Else
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
End If


End Sub

Sub CodigoEjecucionViernes()

Workbooks(2).Sheets("Turnos").Activate
Range("L5").Select
Range(Selection, Selection.End(xlDown)).Copy
ThisWorkbook.Activate

  If Range("A3") = "" Then                'Pega las avayas o usuarios
  Range("A2").Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Else
  Range("A2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  End If

Workbooks(2).Sheets("Turnos").Activate                    'Pegas las fechas de los viernes
Range("BW3").Copy
ThisWorkbook.Activate

  If Range("B3") = "" Then
  Range("B2").Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Else
  Range("B2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  End If

Range("A2").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Sheets("Turnos").Activate                     'Pegas los horarios de inicio de los viernes
  Range("BY5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate

If Range("C3") = "" Then
Range("C2").Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Else
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
End If
  
  Workbooks(2).Sheets("Turnos").Activate                     'Pegas los horarios de fin de los viernes
  Range("BZ5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  
If Range("D3") = "" Then
Range("D2").Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Else
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
End If


End Sub

Sub CodigoEjecucionSabado()

Workbooks(2).Sheets("Turnos").Activate
Range("L5").Select
Range(Selection, Selection.End(xlDown)).Copy
ThisWorkbook.Activate

  If Range("A3") = "" Then                'Pega las avayas o usuarios
  Range("A2").Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Else
  Range("A2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  End If

Workbooks(2).Sheets("Turnos").Activate                    'Pegas las fechas de los sabado
Range("CH3").Copy
ThisWorkbook.Activate

  If Range("B3") = "" Then
  Range("B2").Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Else
  Range("B2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  End If

Range("A2").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Sheets("Turnos").Activate                     'Pegas los horarios de inicio de los sabado
  Range("CJ5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate

If Range("C3") = "" Then
Range("C2").Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Else
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
End If
  
  Workbooks(2).Sheets("Turnos").Activate                     'Pegas los horarios de fin de los sabado
  Range("CK5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  
If Range("D3") = "" Then
Range("D2").Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Else
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
End If


End Sub

Sub CodigoEjecucionDomingo()

Workbooks(2).Sheets("Turnos").Activate
Range("L5").Select
Range(Selection, Selection.End(xlDown)).Copy
ThisWorkbook.Activate

  If Range("A3") = "" Then                'Pega las avayas o usuarios
  Range("A2").Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Else
  Range("A2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  End If

Workbooks(2).Sheets("Turnos").Activate                    'Pegas las fechas de los domingo
Range("CS3").Copy
ThisWorkbook.Activate

  If Range("B3") = "" Then
  Range("B2").Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  Else
  Range("B2").Select
  Selection.End(xlDown).Select
  SendKeys "{DOWN}", True
  Selection.PasteSpecial xlPasteValues
  End If

Range("A2").Select
Selection.End(xlDown).Select
SendKeys "{RIGHT}", True
SendKeys "^+{UP}", True
Selection.FillDown

  Workbooks(2).Sheets("Turnos").Activate                     'Pegas los horarios de inicio de los domingo
  Range("CU5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate

If Range("C3") = "" Then
Range("C2").Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Else
Range("C2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
End If
  
  Workbooks(2).Sheets("Turnos").Activate                     'Pegas los horarios de fin de los domingo
  Range("CV5").Select
  Range(Selection, Selection.End(xlDown)).Copy
  ThisWorkbook.Activate
  
If Range("D3") = "" Then
Range("D2").Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
Else
Range("D2").Select
Selection.End(xlDown).Select
SendKeys "{DOWN}", True
Selection.PasteSpecial xlPasteValues
End If


End Sub

Sub RespuestaYes()

Dim Acrhivos As String
Archivos = Dir("\\Co0000fs0001\planeacion$\16_SISTEMA DE PUNTO\01.Mallas de turno\*.xlsx")
Do While Archivos <> “”

  Workbooks.Open ("\\Co0000fs0001\planeacion$\16_SISTEMA DE PUNTO\01.Mallas de turno\" & Archivos), , True

Call CodigoEjecucionLunes
Call CodigoEjecucionMartes
Call CodigoEjecucionMiercoles
Call CodigoEjecucionJueves
Call CodigoEjecucionViernes
Call CodigoEjecucionSabado
Call CodigoEjecucionDomingo


Workbooks(2).Activate
SendKeys "{ESCAPE}", True
ActiveWorkbook.Close savechanges:=False


Archivos = Dir
Loop

End Sub

Sub RespuestaNo()

Dim Acrhivos As String
Archivos = Dir("\\Co0000fs0001\planeacion$\16_SISTEMA DE PUNTO\01.Mallas de turno\*.xlsx")
Do While Archivos <> “”

  Workbooks.Open ("\\Co0000fs0001\planeacion$\16_SISTEMA DE PUNTO\01.Mallas de turno\" & Archivos), , True

If Format(Date, "DDDD") = "lunes" Then
Call CodigoEjecucionMartes
Call CodigoEjecucionMiercoles
Call CodigoEjecucionJueves
Call CodigoEjecucionViernes
Call CodigoEjecucionSabado
Call CodigoEjecucionDomingo
Else
If Format(Date, "DDDD") = "martes" Then
Call CodigoEjecucionMiercoles
Call CodigoEjecucionJueves
Call CodigoEjecucionViernes
Call CodigoEjecucionSabado
Call CodigoEjecucionDomingo
Else
If Format(Date, "DDDD") = "miércoles" Then
Call CodigoEjecucionJueves
Call CodigoEjecucionViernes
Call CodigoEjecucionSabado
Call CodigoEjecucionDomingo
Else
If Format(Date, "DDDD") = "jueves" Then
Call CodigoEjecucionViernes
Call CodigoEjecucionSabado
Call CodigoEjecucionDomingo
Else
If Format(Date, "DDDD") = "viernes" Then
Call CodigoEjecucionSabado
Call CodigoEjecucionDomingo
Else
End If
End If
End If
End If
End If


Workbooks(2).Activate
SendKeys "{ESCAPE}", True
ActiveWorkbook.Close savechanges:=False


Archivos = Dir
Loop

End Sub