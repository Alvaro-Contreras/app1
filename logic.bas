Attribute VB_Name = "logic"
Option Explicit

Sub procesar_archivos()
      UserForm1.Show
End Sub

Function main(textBoxFecha, opcionPuc, opcionCuif, TextBox_RutaOrigen, TextBox_RutaDestino)
    Dim res, numero_archivos, indice, listado_archivos_origen, nombre_archivo, numero_archivos_destino
    Dim is_substring, fecha, listado_archivos_destino, indice_destino, existe, respuesta, timestamp
    Dim archivo As String, ruta_origen As String, ruta_destino As String
    Dim copiados: copiados = 0
    
    res = validar_formulario(textBoxFecha, opcionPuc, opcionCuif, TextBox_RutaOrigen, TextBox_RutaDestino)
    If res(0) Then
        UserForm1.Hide
        ruta_origen = res(3) & "\"
        ruta_destino = res(4) & "\"
        listado_archivos_origen = listar_archivos(ruta_origen)
        listado_archivos_destino = listar_archivos(ruta_destino)
        fecha = Split(res(1), "/")(2) & Split(res(1), "/")(1) & Split(res(1), "/")(0)

        numero_archivos = UBound(listado_archivos_origen)
        If numero_archivos > 0 Then
            For indice = 0 To numero_archivos
                archivo = ruta_origen & (listado_archivos_origen(indice))
                nombre_archivo = UCase(listado_archivos_origen(indice))
                If InStr(1, nombre_archivo, fecha, vbTextCompare) > 0 And _
                    InStr(1, nombre_archivo, res(2), vbTextCompare) > 0 And _
                    InStr(1, nombre_archivo, ".xls", vbTextCompare) > 0 Then
                    numero_archivos_destino = UBound(listado_archivos_destino)
                    existe = False
                    For indice_destino = 0 To numero_archivos_destino
                        If nombre_archivo = listado_archivos_destino(indice_destino) Then
                            existe = True
                            respuesta = MsgBox("El archivo " & nombre_archivo & " ya existe en la ruta destino. ¿Desea crear una copia?", vbYesNo)
                            If respuesta = vbNo Then
                                MsgBox "El archivo no fue copiado."
                                Exit Function
                            End If
                        End If
                    Next
                    If existe Then
                        timestamp = Format(now, "yyyymmdd_hhmmss")
                        nombre_archivo = Split(nombre_archivo, ".")(0) & "_" & timestamp & "." & Split(nombre_archivo, ".")(1)
                    End If
                    FileCopy archivo, ruta_destino & "\" & nombre_archivo
                    copiados = copiados + 1
                End If
            Next
        Else
            MsgBox "No hay archivos en la ruta de origen"
        End If
    Else
        MsgBox "Todos los campos deben estar diligenciados"
    End If
    MsgBox "Total archivos copiados: " & copiados
End Function
  
Function listar_archivos(ruta As String) As Variant
    Dim archivos() As Variant
    Dim i As Integer
    Dim archivo As String
    
    i = 0
    ReDim archivos(0)
    
    If Right(ruta, 1) <> "\" Then
        ruta = ruta & "\"
    End If
    
    archivo = Dir(ruta & "*.*")
    
    Do While archivo <> ""
        If Not (archivo = "." Or archivo = "..") Then
            ReDim Preserve archivos(i)
            archivos(i) = archivo
            i = i + 1
        End If
        archivo = Dir
    Loop
    
    listar_archivos = archivos
End Function
 
  
Function validar_formulario(textBoxFecha, opcionPuc, opcionCuif, TextBox_RutaOrigen, TextBox_RutaDestino) As Variant
    Dim res(4)
    Dim opcion As String
    Dim fecha_valida As Boolean:    fecha_valida = False
    Dim opcion_valida As Boolean:   opcion_valida = False
    Dim opcion_rutas As Boolean:    opcion_rutas = False
    
    If textBoxFecha.value <> "" Then
        fecha_valida = True
    End If
    
    If opcionPuc.value = True Or opcionCuif.value = True Then
        If opcionPuc.value = True Then
            opcion = "PUC"
        Else
            opcion = "CUIF"
        End If
        opcion_valida = True
    End If
            
    If TextBox_RutaOrigen <> "" And TextBox_RutaDestino <> "" Then
        opcion_rutas = True
    End If
    
    If fecha_valida And opcion_valida And opcion_rutas Then
        res(0) = True
        res(1) = textBoxFecha.value
        res(2) = opcion
        res(3) = TextBox_RutaOrigen.value
        res(4) = TextBox_RutaDestino.value
    Else
        res(0) = False
    End If
    
    validar_formulario = res
End Function
