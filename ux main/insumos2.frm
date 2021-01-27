VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} insumos2 
   Caption         =   "REGISTRO DE INSUMOS: USUARIO ADMIN."
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8025
   OleObjectBlob   =   "insumos2.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "insumos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'--CÓDIGO DEL BOTÓN BUSCAR--
Private Sub BtnSearch_Click()
'CÓDIGO PARA BUSCAR A TRAVÉS DE UN FOR CON INCREMENTO
Dim numerodedatos As Integer
Dim fila As Variant
Dim nombre As String
Dim racion As String
Dim departamento As String
Dim descripcion As String
Dim precio As Variant
Dim Y As Integer

'BUSQUEDA POR NOMBRE
If OptNom = True Then
numerodedatos = Hoja1.Range("A" & Rows.Count).End(xlUp).Row
 LisTab.RowSource = ""
Y = 0

For fila = 1 To numerodedatos

 nombre = ActiveSheet.Cells(fila, 2).Value
 
 If nombre Like "*" & Me.TxtSearch.Value & "*" Then
    Me.LisTab.AddItem
    Me.LisTab.List(Y, 0) = ActiveSheet.Cells(fila, 1).Value
    Me.LisTab.List(Y, 1) = ActiveSheet.Cells(fila, 2).Value
    Me.LisTab.List(Y, 2) = ActiveSheet.Cells(fila, 3).Value
    Me.LisTab.List(Y, 3) = ActiveSheet.Cells(fila, 4).Value
    Me.LisTab.List(Y, 4) = ActiveSheet.Cells(fila, 5).Value
    Me.LisTab.List(Y, 5) = ActiveSheet.Cells(fila, 6).Value
    Me.LisTab.List(Y, 6) = ActiveSheet.Cells(fila, 7).Value
    Y = Y + 1
 End If
 
Next
    End If

'BUSQUEDA POR RACIÓN

    If OptRac = True Then
numerodedatos = Hoja1.Range("A" & Rows.Count).End(xlUp).Row
 LisTab.RowSource = ""
Y = 0

For fila = 1 To numerodedatos

 racion = ActiveSheet.Cells(fila, 3).Value
 
 If racion Like "*" & Me.TxtSearch.Value & "*" Then
    Me.LisTab.AddItem
    Me.LisTab.List(Y, 0) = ActiveSheet.Cells(fila, 1).Value
    Me.LisTab.List(Y, 1) = ActiveSheet.Cells(fila, 2).Value
    Me.LisTab.List(Y, 2) = ActiveSheet.Cells(fila, 3).Value
    Me.LisTab.List(Y, 3) = ActiveSheet.Cells(fila, 4).Value
    Me.LisTab.List(Y, 4) = ActiveSheet.Cells(fila, 5).Value
    Me.LisTab.List(Y, 5) = ActiveSheet.Cells(fila, 6).Value
    Me.LisTab.List(Y, 6) = ActiveSheet.Cells(fila, 7).Value
    Y = Y + 1
 End If
 
Next
    End If
    
'BUSQUEDA POR DEPARTAMENTO
 If OptDep = True Then
numerodedatos = Hoja1.Range("A" & Rows.Count).End(xlUp).Row
 LisTab.RowSource = ""
Y = 0

For fila = 1 To numerodedatos

 departamento = ActiveSheet.Cells(fila, 4).Value
 
 If departamento Like "*" & Me.TxtSearch.Value & "*" Then
    Me.LisTab.AddItem
    Me.LisTab.List(Y, 0) = ActiveSheet.Cells(fila, 1).Value
    Me.LisTab.List(Y, 1) = ActiveSheet.Cells(fila, 2).Value
    Me.LisTab.List(Y, 2) = ActiveSheet.Cells(fila, 3).Value
    Me.LisTab.List(Y, 3) = ActiveSheet.Cells(fila, 4).Value
    Me.LisTab.List(Y, 4) = ActiveSheet.Cells(fila, 5).Value
    Me.LisTab.List(Y, 5) = ActiveSheet.Cells(fila, 6).Value
    Me.LisTab.List(Y, 6) = ActiveSheet.Cells(fila, 7).Value
    Y = Y + 1
 End If
 
Next
    End If

'BUSQUEDA POR DESCRIPCIÓN
 If OptDes = True Then
numerodedatos = Hoja1.Range("A" & Rows.Count).End(xlUp).Row
 LisTab.RowSource = ""
Y = 0

For fila = 1 To numerodedatos

 descripcion = ActiveSheet.Cells(fila, 5).Value
 
 If descripcion Like "*" & Me.TxtSearch.Value & "*" Then
    Me.LisTab.AddItem
    Me.LisTab.List(Y, 0) = ActiveSheet.Cells(fila, 1).Value
    Me.LisTab.List(Y, 1) = ActiveSheet.Cells(fila, 2).Value
    Me.LisTab.List(Y, 2) = ActiveSheet.Cells(fila, 3).Value
    Me.LisTab.List(Y, 3) = ActiveSheet.Cells(fila, 4).Value
    Me.LisTab.List(Y, 4) = ActiveSheet.Cells(fila, 5).Value
    Me.LisTab.List(Y, 5) = ActiveSheet.Cells(fila, 6).Value
    Me.LisTab.List(Y, 6) = ActiveSheet.Cells(fila, 7).Value
    Y = Y + 1
 End If
 
Next
    End If
    
'BUSQUEDA POR PRECIO
If OptPre = True Then
numerodedatos = Hoja1.Range("A" & Rows.Count).End(xlUp).Row
 LisTab.RowSource = ""
Y = 0

For fila = 1 To numerodedatos

 precio = ActiveSheet.Cells(fila, 7).Value
 
 If precio Like "*" & Me.TxtSearch.Value & "*" Then
    Me.LisTab.AddItem
    Me.LisTab.List(Y, 0) = ActiveSheet.Cells(fila, 1).Value
    Me.LisTab.List(Y, 1) = ActiveSheet.Cells(fila, 2).Value
    Me.LisTab.List(Y, 2) = ActiveSheet.Cells(fila, 3).Value
    Me.LisTab.List(Y, 3) = ActiveSheet.Cells(fila, 4).Value
    Me.LisTab.List(Y, 4) = ActiveSheet.Cells(fila, 5).Value
    Me.LisTab.List(Y, 5) = ActiveSheet.Cells(fila, 6).Value
    Me.LisTab.List(Y, 6) = ActiveSheet.Cells(fila, 7).Value
    Y = Y + 1
 End If
 
Next
    End If

End Sub
'--END BOTÓN BUSCAR--

'--CÓDIGO DE LOS COMBOBOX--
Private Sub CmbRac_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'CÓDIGO PARA CAMBIAR LA PROPIEDAD LISTBOX AL INICIAR EL FORM
CmbRac.Style = 2
End Sub

Private Sub CmbDep_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'CÓDIGO PARA CAMBIAR LA PROPIEDAD LISTBOX AL INICIAR EL FORM
CmbDep.Style = 2
End Sub

Private Sub CmbUni_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'CÓDIGO PARA CAMBIAR LA PROPIEDAD LISTBOX AL INICIAR EL FORM
CmbUni.Style = 2
End Sub
'--END COMBOBOX--

'--CÓDIGO BOTÓN REGISTRAR--
Private Sub BtnReg_Click()
Dim Uf As Long

'CÓDIGO PARA VERIFICAR QUE LOS CAMPOS ESTÉN LLENOS
If TxtNom = Empty Or CmbRac = Empty Or CmbDep = Empty Or TxtDes = Empty Or CmbUni = Empty Or TxtCos = Empty Then
MsgBox ("Ingrese todos los datos")
Else
'CÓDIGO PARA INSERTAR VALORES DESDE EL FORMULARIO HACIA LA HOJA DE EXCEL
Label9.Caption = Uf - 1
Worksheets("insumos").Range("B1").End(xlDown).Offset(1, 0).Value = Remove(Me.TxtNom.Value)
Worksheets("insumos").Range("C1").End(xlDown).Offset(1, 0).Value = Me.CmbRac.Value
Worksheets("insumos").Range("D1").End(xlDown).Offset(1, 0).Value = Me.CmbDep.Value
Worksheets("insumos").Range("E1").End(xlDown).Offset(1, 0).Value = Remove(Me.TxtDes.Value)
Worksheets("insumos").Range("F1").End(xlDown).Offset(1, 0).Value = Me.CmbUni.Value
Worksheets("insumos").Range("G1").End(xlDown).Offset(1, 0).Value = Remove(Me.TxtCos.Value)
MsgBox ("Registrado Satisfactoriamente")
End If
'CÓDIGO PARA LIMPIAR CAMPOS DESPUES DE REGISTRAR
TxtNom = Empty
CmbRac = Empty
CmbDep = Empty
TxtDes = Empty
CmbUni = Empty
TxtCos = Empty
Label9 = Empty

UserForm_Initialize

End Sub

Private Sub CommandButton2_Click()

'CÓDIGO DEL BOTÓN ELIMINAR
Unload Me

End Sub

Private Sub BtnLim_Click()

'CÓDIGO PARA BOTÓN LIMPIAR
TxtNom = Empty
CmbRac = Empty
CmbDep = Empty
TxtDes = Empty
CmbUni = Empty
TxtCos = Empty
Label9 = Empty
LisTab = Empty

End Sub
'--END BOTÓN REGISTRAR


'--CÓDIGO ELIMINAR--
Private Sub BtnEli_Click()
'CÓDIGO PARA VERIFICAR SI LOS CAMPOS ESTÁN VACÍOS
If LisTab.ListIndex = -1 Then
MsgBox ("Seleccione un registro")
Else
    
Dim fila As Object
Dim linea As Integer
Dim valor_buscado As Variant
Dim pregunta As String
valor_buscado = UserForm1.TxtBoxID.Value
     
         pregunta = MsgBox("Deseas continuar", vbOKCancel + vbQuestion, "ELIMINAR")
         If pregunta = vbCancel Then
         Else
         'CÓDIGO PARA DETERMINAR LA ÚLTIMA FILA
         Set fila = Sheets("insumos").Range("A:A").Find(valor_buscado, LookAt:=xlWhole)
         linea = fila.Row
         'CÓDIGO PARA BORRAR EL REGISTRO SELECCIONADO
         Range("A" & linea).EntireRow.Delete
         End If
                
End If

End Sub
'--END ELIMINAR--

'--CÓDIGO BOTÓN MODIFICAR--
Private Sub BtnMod_Click()
'CÓDIGO PARA VERIFICAR SI LOS CAMPOS ESTÁN VACÍOS
If LisTab.ListIndex = -1 Then
MsgBox ("Seleccione un registro")
Else
'CÓDIGO PARA SOBREESCRIBIR UN REGISTRO

Dim fila As Object
Dim linea As Integer
Dim valor_buscado As Variant
valor_buscado = Me.TxtBoxID
Set fila = Sheets("insumos").Range("A:A").Find(valor_buscado, LookAt:=xlWhole)
linea = fila.Row
Range("B" & linea).Value = Me.TxtNom.Value
Range("C" & linea).Value = Me.CmbRac.Value
Range("D" & linea).Value = Me.CmbDep.Value
Range("E" & linea).Value = Me.TxtDes.Value
Range("F" & linea).Value = Me.CmbUni.Value
Range("G" & linea).Value = Me.TxtCos.Value
MsgBox ("Modificado Exitosamente")
End If

End Sub
'--END BOTÓN ELIMINAR--

Private Sub Image1_Click()

End Sub

'--CÓDIGO LISTBOX--
Private Sub LisTab_Click()
Dim codigo As Integer
codigo = LisTab.List(LisTab.ListIndex, 0)
Me.TxtBoxID.Value = codigo
Me.Label9.Caption = codigo
End Sub
'--END LISTBOX--

'--CÓDIGO TEXTBOX--
Private Sub TxtNom_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'CÓDIGO PARA QUE EN EL TxtNom SE ESCRIBA EN MAYÚSCULA
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If
End Sub

Private Sub TxtDes_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Dim mayus As String
'CÓDIGO PARA QUE EN EL TxtDes SE ESCRIBA EN MAYÚSCULA
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If
End Sub

Private Sub TxtCos_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'CÓDIGO PARA QUE EN EL TxtCos SE ESCRIBA EN MAYÚSCULA
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If
End Sub
'--END TEXTBOX--

'--CÓDIGO TxtBoxID--
Private Sub TxtBoxID_Change()
Dim codigo As Integer
codigo = TxtBoxID.Value
On Error Resume Next
Me.TxtNom = Application.WorksheetFunction.VLookup(codigo, Sheets("insumos").Range("A:G"), 2, 0)
Me.CmbRac = Application.WorksheetFunction.VLookup(codigo, Sheets("insumos").Range("A:G"), 3, 0)
Me.CmbDep = Application.WorksheetFunction.VLookup(codigo, Sheets("insumos").Range("A:G"), 4, 0)
Me.TxtDes = Application.WorksheetFunction.VLookup(codigo, Sheets("insumos").Range("A:G"), 5, 0)
Me.CmbUni = Application.WorksheetFunction.VLookup(codigo, Sheets("insumos").Range("A:G"), 6, 0)
Me.TxtCos = Application.WorksheetFunction.VLookup(codigo, Sheets("insumos").Range("A:G"), 7, 0)

End Sub
'--END TxtBoxID--'

'--CÓDIGO TXTBOX SEARCH--
Private Sub TxtSearch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
   If KeyCode = vbKeyReturn Then
     BtnSearch_Click
   End If

   If KeyCode = vbKeySeparator Then
     BtnSearch_Click
   End If
End Sub

Private Sub TxtSearch_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'CÓDIGO PARA QUE EN EL TxtCos SE ESCRIBA EN MAYÚSCULA
If (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 225) Or (KeyAscii = 233) Or (KeyAscii = 237) Or (KeyAscii = 241) Or (KeyAscii = 243) Or (KeyAscii = 250) Then
KeyAscii = VBA.Asc(VBA.UCase(VBA.Chr(KeyAscii)))
End If
End Sub
'--END TXTBOX SEARCH--

'--CÓDIGO USERFORM--
Private Sub UserForm_Initialize()

'CÓDIGO PARA CERRAR EL DELETEFORM AL HACER LOGIN
Unload DeleteForm

'CÓDIGO PARA MOSTRAR LOS BOTONES DE COLOR
BtnReg.BackColor = RGB(90, 50, 176)
BtnLim.BackColor = RGB(90, 50, 176)
BtnMod.BackColor = RGB(90, 50, 176)
BtnEli.BackColor = RGB(90, 50, 176)
BtnSearch.BackColor = RGB(90, 50, 176)

'CODIGO PARA VISUALIZAR EL LISTBOX ARRIBA
LisTab.Top = 354

'CODIGO PARA OCULTAR
    
    'EL TxtBoxID
    TxtBoxID.Visible = False

    'BOTÓN MODIFICAR
    BtnMod.Visible = True
  
'CÓDIGO PARA INSERTAR DATOS EN EL LISTBOX
Dim MiRango As Range
Dim MiRango2 As Range
Dim Columnas As Integer

Set MiRango = Sheets("insumos").Range("A1").CurrentRegion
Set MiRango2 = MiRango.Offset(1, 0).Resize(MiRango.Rows.Count - 1, MiRango.Columns.Count)

MiRango2.Name = "MiTabla"
Columnas = MiRango2.Columns.Count

With Me.LisTab
    .ColumnCount = Columnas
    .ColumnWidths = "20 pt; 60pt; 60pt; 60pt; 40pt; 30pt; 30pt"
    .ColumnHeads = True
    .RowSource = "MiTabla"
End With

'CÓDIGO PARA MOSTRAR EL EVENTO SUCCESS
Label9 = ""

'CÓDIGO PARA SELECCIONAR LA OPCIÓN NOMBRE POR DEFECTO
OptNom.Value = True

'CÓDIGO PARA LLENAR EL COMBOBOX DESDE EL PRINCIPIO
CmbRac.AddItem "RACIÓN CALIENTE"
CmbRac.AddItem "RACIÓN FRÍA"
CmbDep.AddItem "CARNES, HUEVO Y EMBUTIDO"
CmbDep.AddItem "DERIVADOSY LACTEOS"
CmbDep.AddItem "ABARROTES"
CmbDep.AddItem "FRUTAS Y VERDURAS"
CmbUni.AddItem "KG"
CmbUni.AddItem "LT"
CmbUni.AddItem "PZA"
CmbUni.AddItem "PQTE"

End Sub
'--END USERFORM--










