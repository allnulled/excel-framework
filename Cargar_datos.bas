Attribute VB_Name = "Cargar_datos"
Public Sub cargar_datos_por_ajax()

    ' 1. Definimos parámetros iniciales
    Dim ajax_metodo As String
    Dim ajax_url As String
    Dim ajax_data As Variant
    Dim hoja_destino_string As String
    ajax_metodo = "POST"
    ajax_url = "http://192.168.1.40"
    ajax_data = "{""operation"":""select"",""table"":""Usuario""}"
    hoja_destino_string = "Hoja1"
    
    ' 2. Enviamos petición XMLHTTP
    Dim xhr As Object
    Dim xhr_resultado As Variant
    Set xhr = CreateObject("MSXML2.XMLHTTP")
    xhr.Open ajax_metodo, ajax_url, False
    xhr.Send ajax_data
    
    ' 3. Extraemos la respuesta
    Dim respuesta As String
    respuesta = xhr.responseText
    
    ' 4. La convertimos en JSON. Para esto se requiere de la librería de aquí: https://raw.githubusercontent.com/VBA-tools/VBA-JSON/master/JsonConverter.bas
    Dim datos As Object
    Set datos = JsonConverter.ParseJson(respuesta)
    
    ' 5. Ponemos los datos en la hoja nuevamente:
    Dim rango_inicial As Range
    Dim datos_indice As Integer
    Set rango_inicial = Sheets(hoja_destino_string).Range("A1")
    rango_inicial.Offset(0, 0).Value = "Nombre"
    rango_inicial.Offset(0, 1).Value = "Contraseña"
    rango_inicial.Offset(0, 2).Value = "Correo"
    For datos_indice = 1 To datos.Count - 1
        Dim dato As Variant
        Set dato = datos.Item(datos_indice)
        Dim nombre As Variant
        Dim contrasenya As Variant
        Dim correo As Variant
        nombre = dato.Item("nombre")
        contrasenya = dato.Item("contrasenya")
        correo = dato.Item("correo")
        rango_inicial.Offset(datos_indice, 0).Value = nombre
        rango_inicial.Offset(datos_indice, 1).Value = contrasenya
        rango_inicial.Offset(datos_indice, 2).Value = correo
    Next datos_indice
    
End Sub

