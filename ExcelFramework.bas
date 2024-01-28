Attribute VB_Name = "ExcelFramework"
Public Function REGEX_TEST(texto, regex_texto)
    '
    ' @name REGEX_TEST
    ' @description Function to test regular expressions against text
    ' @parameter texto String - Text to test
    ' @parameter regex_texto String - Regular expression to test
    ' @returns Boolean - Weather or not the regular expression matches somehow
    '
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = regex_texto
    REGEX_TEST = regex.Test(texto)
End Function

Public Function REGEX_MATCHES(texto, regex_texto, delimitador)
    '
    ' @name REGEX_MATCHES
    ' @description Function to extract regular expression matches from a text
    ' @parameter texto String - Text to test
    ' @parameter regex_texto String - Regular expression to test
    ' @parameter delimitador String - Delimiter to use when joining the matches
    ' @returns String - Matches of first level joined by the 'delimitador' parameter
    '
    Dim coincidencias As Object
    Dim regex As Object
    Dim es_global As Boolean
    Dim ignorar_mayusculas As Boolean
    es_global = True
    ignorar_mayusculas = False
    Set regex = CreateObject("VBScript.RegExp")
    Set coincidencias = CreateObject("System.Collections.ArrayList")
    regex.Pattern = regex_texto
    Dim matches As Object
    If es_global = False Then
        regex.Global = es_global
    End If
    If ignorar_mayuscula = True Then
        regex.IgnoreCase = ignorar_mayuscula
    End If
    Set matches = regex.Execute(texto)
    For Each Match In matches
        coincidencias.Add Match.Value
    Next Match
    REGEX_MATCHES = Join(coincidencias.ToArray, delimitador)
End Function

Public Function HEX_TO_RGB(color_hex_param As String)
    '
    ' @name HEX_TO_RGB
    ' @description Function to pass color values from HEX to RGB VBA valid value
    ' @parameter texto color_hex_param - Color in hexadecimal
    ' @returns String - The same color in RGB format that VBA accepts
    '
    Dim R As Integer
    Dim G As Integer
    Dim B As Integer
    Dim celda As Range
    Dim color_hex As String
    color_hex = Replace(color_hex_param, "#", "")
    color_hex = Right$("000000" & color_hex_param, 6)
    R = Val("&H" & Mid(color_hex, 1, 2))
    G = Val("&H" & Mid(color_hex, 3, 2))
    B = Val("&H" & Mid(color_hex, 5, 2))
    HEX_TO_RGB = RGB(R, G, B)
End Function

Public Function HEX_TO_RGB_TEXT(color_hex_param As String)
    '
    ' @name HEX_TO_RGB_TEXT
    ' @description Function to pass color values from HEX to RGB notation
    ' @parameter texto color_hex_param - Color in hexadecimal
    ' @returns String - The same color in RGB format like: rgb({R}, {G}, {B})
    '
    Dim R As Integer
    Dim G As Integer
    Dim B As Integer
    Dim celda As Range
    Dim color_hex As String
    color_hex = Replace(color_hex_param, "#", "")
    color_hex = Right$("000000" & color_hex_param, 6)
    R = Val("&H" & Mid(color_hex, 1, 2))
    G = Val("&H" & Mid(color_hex, 3, 2))
    B = Val("&H" & Mid(color_hex, 5, 2))
    HEX_TO_RGB_TEXT = "rgb(" & R & "," & G & "," & B & ")"
End Function

Function RANDOM_INTEGER(minimo, maximo)
    '
    ' @name RANDOM_INTEGER
    ' @description Function to get a random integer number
    ' @parameter minimo Integer - Minimum number of the randomization
    ' @parameter maximo Integer - Maximum number of the randomization
    ' @returns Integer - A random value between 'minimo' and 'maximo'
    '
    RANDOM_INTEGER = Int(Rnd * (maximo - minimo + 1))
End Function

Function RANDOM_STRING(longitud As Integer)
    '
    ' @name RANDOM_STRING
    ' @description Function to get a random string of variable length
    ' @parameter longitud Integer - Length of the resultant string
    ' @returns String - A random string
    '
    Dim banco_de_caracteres As Variant
    Dim indice As Integer
    Dim texto As String
    If longitud < 1 Then
        MsgBox "Length variable must be greater than 0"
        Exit Function
    End If
    banco_de_caracteres = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
        "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
        "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
        "A", "B", "C", "D", "E", "F", "G", "H", _
        "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
        "W", "X", "Y", "Z")
    For indice = 1 To longitud
        texto = texto & banco_de_caracteres(Int((UBound(banco_de_caracteres) - LBound(banco_de_caracteres) + 1) * Rnd + LBound(banco_de_caracteres)))
    Next indice
    RANDOM_STRING = texto
End Function

Function RANDOM_ITEM_FROM_RANGE(rango_de_items)
    '
    ' @name RANDOM_ITEM_FROM_RANGE
    ' @description Function to get a random item from an array
    ' @parameter lista_de_items Array - Items of the list
    ' @returns String - A random item from the passed array
    '
    Dim lista_de_items As Variant
    lista_de_items = rango_de_items.Value
    Dim item As Variant
    item = lista_de_items(Application.WorksheetFunction.RandBetween(LBound(lista_de_items), UBound(lista_de_items)), 1)
    RANDOM_ITEM_FROM_RANGE = item
End Function

Public Sub autoajustar_documento()
    Application.ActiveSheet.Cells.EntireColumn.AutoFit
    Application.ActiveSheet.Cells.EntireRow.AutoFit
End Sub

Public Sub marcar_seleccion_con_color()
    Dim color As String
    color = InputBox("¿Qué color deseas usar? (Usando notación #HEX)", "Macro de 'marcar_con_color'", "#00FF00")
    For Each celda In Selection
        celda.Interior.color = HEX_TO_RGB(color)
    Next celda
End Sub

Public Sub aleatorizar_seleccion_con_valores_de_entero()
    Dim minimo As Integer
    Dim maximo As Integer
    minimo = InputBox("¿Mínimo?", "Macro de 'aleatorizar_valores_enteros'", "0")
    maximo = InputBox("¿Máximo?", "Macro de 'aleatorizar_valores_enteros'", "100")
    For Each celda In Selection
        Dim aleatorio As Integer
        aleatorio = RANDOM_INTEGER(minimo, maximo)
        celda.Value = aleatorio
    Next celda
End Sub

Public Sub aleatorizar_seleccion_con_valores_de_texto()
    Dim longitud As Integer
    longitud = Int(InputBox("¿Longitud?", "Macro de 'aleatorizar_valores_de_texto'", "10"))
    For Each celda In Selection
        Dim aleatorio As String
        aleatorio = RANDOM_STRING(longitud)
        celda.Value = aleatorio
    Next celda
End Sub

Public Sub pintar_celda_de_color_en_hex()
    For Each celda In Selection
        celda.Interior.color = HEX_TO_RGB(celda.Value)
    Next celda
End Sub




