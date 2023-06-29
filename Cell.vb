Imports System.Text.RegularExpressions
Imports System.Drawing

Public Class Cell
#Region "Parámetros"
    Public XlsApp As Microsoft.Office.Interop.Excel.Application
    Public infoSheet As DataWorksheet
#End Region

#Region "Constantes"
    'Info extra: https://learn.microsoft.com/en-us/office/vba/api/excel.constants
    'FONTS
    Const xlUnderlineStyleDouble As Short = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleDouble
    Const xlUnderlineStyleDoubleAccounting As Short = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleDoubleAccounting
    Const xlUnderlineStyleNone As Short = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
    Const xlUnderlineStyleSingle As Short = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle

    'ALIGNMENT
    Const xlHAlignLeft As Short = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
    Const xlHAlignRight As Short = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
    Const xlHAlignCenter As Short = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
    Const xlHAlignJustify As Short = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignJustify
    Const xlVAlignTop As Short = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop
    Const xlVAlignBottom As Short = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom
    Const xlVAlignCenter As Short = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
    Const xlHAlignFill As Short = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignFill

    'BORDERS
    Const xlEdgeLeft As Short = Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft
    Const xlEdgeRight As Short = Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight
    Const xlEdgeTop As Short = Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop
    Const xlEdgeBottom As Short = Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom
    Const xlInsideHorizontal As Short = Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal
    Const xlInsideVertical As Short = Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical
    '---
    Const xlContinuous As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    Const xlDash As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlDash
    Const xlDashDot As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlDashDot
    Const xlDashDotDot As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlDashDotDot
    Const xlDot As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot
    Const xlDouble As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
    Const xlLineStyleNone As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone
    Const xlSlantDashDot As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlSlantDashDot
#End Region

#Region "Funciones"

    ''' <summary>
    ''' Crea una celda (VERSIÓN 1)
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <param name="row1"></param>
    ''' <param name="column1"></param>
    ''' <param name="row2"></param>
    ''' <param name="column2"></param>
    ''' <param name="Configurations"></param>
    Public Sub Cell(ByVal Value As String, ByVal row1 As Short, ByVal column1 As Short, Optional ByVal row2 As Short = 0, Optional ByVal column2 As Short = 0, Optional Configurations As String = "")
        Dim Cell_Init As String = L(column1) & row1        'Celda de inicio
        Dim Cell_Final As String = If(row2 = 0 Or column2 = 0, "", L(column2) & row2)        'Celda final

        Cell(Value, Cell_Init, Cell_Final, Configurations)      'Apuntando a la Versión 2
    End Sub

    ''' <summary>
    ''' Crea una celda (VERSIÓN 2)
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <param name="Cell_Init"></param>
    ''' <param name="Cell_Final"></param>
    ''' <param name="Configurations"></param>
    Public Sub Cell(ByVal Value As String, ByVal Cell_Init As String, Cell_Final As String, ByVal Optional Configurations As String = "")
        With XlsApp.Worksheets(infoSheet.SheetName)
            Dim IsMerge As Boolean = True       'es celda convinada?

            If Value IsNot Nothing Then
                Value = Trim(Value)                                 'Quitamos espacios 
            End If

            'Verificar si habrá convinación de celdas o no
            Dim rangeNotMergeConfigurationNivel As String = "range"       'Nivel de configuración de rango sin convinación de celda: none | range | cell   (se hace mayormente por el comportamiento de los bordes)
            If Regex.IsMatch(Cell_Init, "^<[i]?>.*$") Then
                IsMerge = False

                If Regex.IsMatch(Cell_Init, "<[i]>") Then
                    rangeNotMergeConfigurationNivel = "cell"
                End If

                Cell_Init = Regex.Replace(Cell_Init, "<[i]?>", "")  'Eliminar ese sufijo del Cell_Init tenga el parámetro i o no lo tenga (<> | <i>)
            End If

            Cell_Init = Cell_Init.ToUpper()                     'se trabajarán en mayúsculas
            Cell_Final = Cell_Final.ToUpper()                   'se trabajarán en mayúsculas

            '--------Sacamos el valor del rango para el caso
            Dim _rangeMerge As String = "", _rangeValue As String = ""
            If Cell_Final = "" Then
                'si sólo se menciona la celda_inicial
                _rangeMerge = Cell_Init
                _rangeValue = Cell_Init
            ElseIf Regex.IsMatch($"{Cell_Init}:{Cell_Final}", "[A-Z]+:[A-Z]+") And Not Regex.IsMatch($"{Cell_Init}:{Cell_Final}", "\d+") Then
                'si se refiere a una columna entera... G:G
                _rangeMerge = $"{Cell_Init}:{Cell_Final}"
                _rangeValue = $"{Cell_Init}1"
            Else
                'caso general
                _rangeMerge = $"{Cell_Init}:{Cell_Final}"
                _rangeValue = Cell_Init
            End If

            If Value IsNot Nothing Then
                '--------Es o no Fórmula(si tiene el signo "=" al inicio(index 0) => sí)
                If (Value.IndexOf("=") = 0) Then
                    .Range(_rangeValue).FormulaLocal = Value
                Else
                    .Range(_rangeValue).Value = Value
                End If
            End If


            '--------Aplicación
            If IsMerge Then
                If Cell_Final <> "" Then
                    .Range(_rangeMerge).Merge() 'Convinación de múltiples celdas
                End If

                CellConfigurations(.Range(_rangeValue), Configurations)
            Else
                CellConfigurations(.Range(_rangeMerge), Configurations, rangeNotMergeConfigurationNivel)
            End If
        End With
    End Sub

    ''' <summary>
    ''' Crea una celda (VERSIÓN 3)
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <param name="RangeCell"></param>
    ''' <param name="Configurations"></param>
    Public Sub Cell(ByVal Value As String, ByVal RangeCell As String, ByVal Optional Configurations As String = "")
        Dim Cells() As String = RangeCell.Split(":")        'Ejemplo A1:A20
        Dim Cell_Final As String = If(Cells.Length > 1, Cells(1), "")

        Cell(Value, Cells(0), Cell_Final, Configurations)   'Apuntando a la versión 2
    End Sub

    ''' <summary>
    ''' L = convert Number To Letter / convertidor de números a letras 
    ''' </summary>
    ''' <param name="ColumnNumber"></param>
    ''' <returns></returns>
    Public Shared Function L(ColumnNumber As Long) As String
        Dim a As Long
        Dim b As Long
        L = ""
        Do While ColumnNumber > 0
            a = Int((ColumnNumber - 1) / 26)
            b = (ColumnNumber - 1) Mod 26
            L = Chr(b + 65) & L
            ColumnNumber = a
        Loop
    End Function

    ''' <summary>
    ''' Convertir Letra de Columna a número de Columna de Excel :: N = Convert Letter to Number / Convierte la letra de columna de la hoja a número
    ''' </summary>
    ''' <param name="ColumnLetter"></param>
    ''' <returns></returns>
    Public Shared Function N(ColumnLetter As String) As Short
        Dim letter As String = UCase(ColumnLetter)

        Dim letterArray() As Char = letter.ToCharArray()                'Separar los caracteres

        Dim abc As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"                'Alfabeto a utilizar que toma Excel (Inglés EU)
        Dim constant As Short = abc.Length                              'Total de letras en el alfabeto

        Dim columnNumber As Short = 0
        Dim nivelLetter As Short = letterArray.Length - 1               'Nivel máximo como índice

        'Disminuyendo nivel hasta llegar a 0 (todo número elevado a la 0 potencia es igual a 1)
        For Each lett In letterArray
            Dim positionLetter As Short = abc.IndexOf(lett) + 1         'Posición de la letra, multiplicado por los niveles necesarios
            columnNumber += positionLetter * (constant) ^ nivelLetter   'Fórmula para calcular el número de Columna

            nivelLetter -= 1                                            'Bajamos de nivel
        Next

        N = columnNumber                                                'Retornando Número de columna
    End Function

    ''' <summary>
    ''' Configuraciones Generales para la celda (DICCIONARIO DE CONFIGURACIONES)
    ''' FORMATO DE CADENA:::>
    '''     key1: value1; key2:value2; ... ; keyN:valueN
    '''     
    '''     key1: valueMultipleOption1 valueMultipleOption2 valueMultipleOption3; ... ; keyN: valueMultipleOptionN          (divididos por espacios) la key tiene múltiples configuraciones
    '''                                                                                                                     
    '''     key1:  valueOptionAvailable1, valueOptionAvailable2, valueOptionAvailableN; ... ; keyN: valueOptionAvailableN   (divididos por comas): la primera configuración en la lista es el primero en validarse
    '''     
    '''     key1: 'value String 1'; ... ; keyN: 'value String N'
    ''' </summary>
    ''' <param name="CellRange"></param>
    ''' <param name="Configurations"></param>
    Private Sub CellConfigurations(ByVal CellRange As Microsoft.Office.Interop.Excel.Range, Configurations As String, Optional RangeNotMergeConfigurationNivel As String = "range")
        'Obtenemos nueva configuración predefinida para la celda
        Dim _CellConfig As Dictionary(Of String, String) = Dictionary.CellConfigurations

        'Interpretador Léxico de las configuraciones definidas/predefinidas
        LexicalInterpreter(_CellConfig, Configurations)

        'Ejecutando las configuraciones predefinidas de la celda en la hoja de Excel
        ConfigurationExecution(_CellConfig, CellRange, RangeNotMergeConfigurationNivel)

    End Sub

    ''' <summary>
    ''' Interpretador Léxico de las configuraciones para la celda: lenguaje tipo CSS
    ''' </summary>
    ''' <param name="_CellConfig"></param>
    ''' <param name="Configurations"></param>
    Private Sub LexicalInterpreter(_CellConfig As Dictionary(Of String, String), Optional Configurations As String = "")
        Try
            If (Configurations <> "") Then
                Dim _regx_bracket As String = "a?rgb\s*\((\s*\d{1,3}\s*,?)+\)"                  'Configuración entre paréntesis para no afectar los splits con comas y otra manera
                Dim _regx_string As String = "([""'])(.*?)\1"      'Expresión regular para la sustitución de los valores tipo string (entre comillas sencillas o dobles)

                'reunir coincidencias dentro de "" ó ''
                Dim outputConfigBrackets As MatchCollection = Regex.Matches(Configurations, _regx_bracket)
                Dim outputConfigStrings As MatchCollection = Regex.Matches(Configurations, _regx_string)

                'almacenarnos en una lista con sus futuras sustituciones (palabras reservadas)
                Dim ConfigStringsDictionary As New Dictionary(Of String, String)

                Dim i As Integer = 1            'iteración para las sustituciones
                Dim substitution As String, _configBracket As String, _configString As String

                'Sustitución de los paréntesis
                For Each configBracket As Match In outputConfigBrackets
                    substitution = $"_confgbrack_{i}_"
                    _configBracket = configBracket.Value
                    ConfigStringsDictionary.Add(substitution, _configBracket)

                    'sustituímos la cadena por una palabra de sustitución
                    Configurations = Replace(Configurations, _configBracket, substitution)
                    i += 1
                Next

                'Sustitución de las cadenas de texto
                i = 1
                For Each configString As Match In outputConfigStrings
                    substitution = $"_confgstr_{i}_"
                    _configString = configString.Value
                    ConfigStringsDictionary.Add(substitution, _configString)

                    'sustituímos la cadena por una palabra de sustitución
                    Configurations = Replace(Configurations, _configString, substitution)
                    i += 1
                Next

                Configurations = Configurations.ToLower()       'con las configuraciones ya seteadas ya podemos convertirlo a minúsculas

                'SEPARADOR - NIVEL 1 (;)
                Dim nivel_1() As String = Configurations.Split(";")

                Dim nivel_2() As String
                Dim key As String, value As String
                For Each nivel1 In nivel_1

                    'SEPARADOR - NIVEL 2 (:)
                    nivel_2 = nivel1.Split(":")

                    key = Trim(nivel_2(0))
                    value = Trim(nivel_2(1))

                    '-----------SEPARADORES - NIVEL 3-----------
                    LexicalKey(key, value, _CellConfig, ConfigStringsDictionary)
                    '---------FIN SEPARADORES - NIVEL 3-----------


                    ''Seteamos todas las configuraciones key:value
                    'If _CellConfig.ContainsKey(key) Then
                    '    value = If(ConfigStringsDictionary.ContainsKey(value), ConfigStringsDictionary(value), value)
                    '    value = Replace(Replace(value, """", ""), "'", "")
                    '    _CellConfig(key) = value
                    'Else
                    '    Console.WriteLine($"key not found: {key}")
                    'End If

                Next
            End If
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Ejecución de las configuraciones definidas/predefinidas en el rango de la Celda seleccionada
    ''' </summary>
    ''' <param name="_CellConfig"></param>
    ''' <param name="CellRange"></param>
    Private Sub ConfigurationExecution(_CellConfig As Dictionary(Of String, String), CellRange As Microsoft.Office.Interop.Excel.Range, Optional RangeNotMergeConfigurationNivel As String = "range")
        Try
            Dim regxARGB As String = "^(a?rgb)?\(?([01]?\d\d?|2[0-4]\d|25[0-5])(\W+)([01]?\d\d?|2[0-4]\d|25[0-5])\W+(([01]?\d\d?|2[0-4]\d|25[0-5])\)?)$"

            'BORDERS
            Dim BorderEnumerations As Dictionary(Of Short, String) = Dictionary.NameAndEnumerationBorders
            If RangeNotMergeConfigurationNivel = "cell" Then
                'Setear los bordes internos por default
                _CellConfig("border-inside-horizontal") = If(BorderType(_CellConfig("border-inside-horizontal")) = xlLineStyleNone, _CellConfig("border-bottom"), _CellConfig("border-inside-horizontal"))
                _CellConfig("border-inside-vertical") = If(BorderType(_CellConfig("border-inside-vertical")) = xlLineStyleNone, _CellConfig("border-right"), _CellConfig("border-inside-vertical"))
            Else
                'Para agregar sólo borde del rango (sin los bordes internos del rango)
                BorderEnumerations.Remove(xlInsideVertical)
                BorderEnumerations.Remove(xlInsideHorizontal)
            End If

            Dim borderWidth As String = ""
            Dim borderStyle As String = ""
            Dim borderColor As String = ""
            Dim allBorder As Boolean = _CellConfig("border") <> "none"
            If allBorder Then
                borderWidth = "border-width"
                If _CellConfig(borderWidth) <> "none" Then
                    CellRange.Borders.Weight = Dictionary.BorderWeightEnumerations(_CellConfig(borderWidth))
                End If

                borderStyle = "border-style"
                If _CellConfig(borderStyle) <> "none" Then
                    CellRange.Borders.LineStyle = BorderType(_CellConfig(borderStyle))
                End If

                borderColor = "border-color"
                If _CellConfig(borderColor) <> "none" Then
                    If (_CellConfig(borderColor).IndexOf("#") = 0) Then      'https://stackoverflow.com/questions/7423456/changing-an-excel-cells-backcolor-using-hex-results-in-excel-displaying-complet
                        CellRange.Borders.Color = FillType(_CellConfig(borderColor), "color-hex")
                    ElseIf Regex.IsMatch(_CellConfig(borderColor), regxARGB) Then
                        CellRange.Borders.Color = FillType(_CellConfig(borderColor), "color-rgb")
                    Else
                        CellRange.Borders.ColorIndex = FillType(_CellConfig(borderColor), "color-palette")
                    End If
                End If
            End If

            'configuración individual de los borders
            For Each border In BorderEnumerations   'Name = Value And Enumeration = Key
                borderWidth = $"{border.Value}-width"
                If _CellConfig(borderWidth) <> "none" Then
                    CellRange.Borders(border.Key).Weight = Dictionary.BorderWeightEnumerations(_CellConfig(borderWidth))
                End If

                borderStyle = $"{border.Value}-style"
                If _CellConfig(borderStyle) <> "none" Then
                    CellRange.Borders(border.Key).LineStyle = BorderType(_CellConfig(borderStyle))
                End If

                borderColor = $"{border.Value}-color"
                If _CellConfig(borderColor) <> "none" Then
                    If (_CellConfig(borderColor).IndexOf("#") = 0) Then      'https://stackoverflow.com/questions/7423456/changing-an-excel-cells-backcolor-using-hex-results-in-excel-displaying-complet
                        CellRange.Borders(border.Key).Color = FillType(_CellConfig(borderColor), "color-hex")
                    ElseIf Regex.IsMatch(_CellConfig(borderColor), regxARGB) Then
                        CellRange.Borders(border.Key).Color = FillType(_CellConfig(borderColor), "color-rgb")
                    Else
                        CellRange.Borders(border.Key).ColorIndex = FillType(_CellConfig(borderColor), "color-palette")
                    End If
                End If
            Next

            'FONTS
            CellRange.Font.Name = _CellConfig("font-family")
            CellRange.Font.Bold = FontStyleType(_CellConfig("font-style"), "bold", _CellConfig) 'Boolean.Parse(If(_CellConfig("font-style") = "bold", True, _CellConfig("bold")))
            CellRange.Font.Italic = FontStyleType(_CellConfig("font-style"), "italic", _CellConfig)  'Boolean.Parse(If(_CellConfig("font-style") = "italic", True, _CellConfig("italic")))
            CellRange.Font.Underline = FontStyleType(_CellConfig("font-style"), "underline", _CellConfig)
            CellRange.Font.Size = Short.Parse(_CellConfig("font-size"))
            If (_CellConfig("color") <> "none") Then
                If (_CellConfig("color").IndexOf("#") = 0) Then      'https://stackoverflow.com/questions/7423456/changing-an-excel-cells-backcolor-using-hex-results-in-excel-displaying-complet
                    CellRange.Font.Color = FillType(_CellConfig("color"), "color-hex")
                ElseIf Regex.IsMatch(_CellConfig("color"), regxARGB) Then
                    CellRange.Font.Color = FillType(_CellConfig("color"), "color-rgb")
                Else
                    CellRange.Font.ColorIndex = FillType(_CellConfig("color"), "color-palette")
                End If
            End If

            'ALIGNMENT
            CellRange.HorizontalAlignment = AlignmentType(_CellConfig("text-align"))
            CellRange.VerticalAlignment = AlignmentType(_CellConfig("vertical-align"), "vertical-align")
            CellRange.WrapText = _CellConfig("text-wrap")

            'NUMBER - TIPO DE FORMATO DE CELDA
            If (_CellConfig("number-format") <> "General") Then
                CellRange.NumberFormat = NumberFormatType(_CellConfig("number-format"))
            End If

            'FILL
            If (_CellConfig("background-color") <> "none") Then
                If (_CellConfig("background-color").IndexOf("#") = 0) Then      'https://stackoverflow.com/questions/7423456/changing-an-excel-cells-backcolor-using-hex-results-in-excel-displaying-complet
                    CellRange.Interior.Color = FillType(_CellConfig("background-color"), "color-hex")
                ElseIf Regex.IsMatch(_CellConfig("background-color"), regxARGB) Then
                    CellRange.Interior.Color = FillType(_CellConfig("background-color"), "color-rgb")
                Else
                    CellRange.Interior.ColorIndex = FillType(_CellConfig("background-color"), "color-palette")
                End If
            End If

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Retorna un tipo de configuración que se utiliza en la sección de Bordes (LineStyle)
    ''' </summary>
    ''' <param name="keyName"></param>
    ''' <returns></returns>
    Private Shared Function BorderType(keyName As String) As Object
        Return If(Dictionary.BorderTypes.ContainsKey(keyName),
                  Dictionary.BorderTypes(keyName),
                  Dictionary.BorderTypes("none")     'default
        )
    End Function

    ''' <summary>
    ''' Retorna un tipo de configuración que se utiliza en la seccíón de Alineamiento
    ''' </summary>
    ''' <param name="keyName"></param>
    ''' <param name="type"></param>
    ''' <returns></returns>
    Private Shared Function AlignmentType(keyName As String, Optional type As String = "text-align") As Object
        Dim result As Object = Nothing

        If type = "text-align" Then
            result = If(Dictionary.AlignmentTypes_Text.ContainsKey(keyName),
                        Dictionary.AlignmentTypes_Text(keyName),
                        Dictionary.AlignmentTypes_Text("none")       'default
            )
        ElseIf type = "vertical-align" Then
            result = If(Dictionary.AlignmentTypes_Vertical.ContainsKey(keyName),
                        Dictionary.AlignmentTypes_Vertical(keyName),
                        Dictionary.AlignmentTypes_Vertical("none")   'default
            )
        End If

        Return result
    End Function

    ''' <summary>
    ''' Retorna un tipo de configuración que se utiliza en la sección de Estilo de Fuentes
    ''' </summary>
    ''' <param name="keyValue"></param>
    ''' <param name="type"></param>
    ''' <param name="_CellConfig"></param>
    ''' <returns></returns>
    Private Shared Function FontStyleType(keyValue As String, Optional type As String = "underline", Optional _CellConfig As Dictionary(Of String, String) = Nothing) As Object
        Dim result As Object = Nothing
        Dim _keyName As String = ""

        If _CellConfig IsNot Nothing Then

            '====>BOLD
            If type = "bold" Then
                'font-style: bold
                result = If(keyValue = "bold", True, _CellConfig("bold"))
            End If
            '=========================================

            '====>ITALIC
            If type = "italic" Then
                'font-style: italic
                result = If(keyValue = "italic", True, _CellConfig("italic"))
            End If
            '=========================================

            '====>UNDERLINE
            If type = "underline" Then
                Dim underlineTypeList As New List(Of String)
                For Each ut In Dictionary.UnderlineTypes.Keys
                    underlineTypeList.Add($"underline-{ut}")
                Next

                'valuekey única en el key ::: font-style: underline | font-style: underline-single | font-style: underline-double | font-style: underline-doubleaccount
                If underlineTypeList.Contains(keyValue) Or keyValue = "underline" Then
                    _CellConfig(If(keyValue = "underline", "underline-single", keyValue)) = True
                End If

                '<> underline: none => underline: single | underline: double | underline: doubleaccount
                If _CellConfig("underline") <> "none" Then
                    _keyName = _CellConfig("underline")
                End If

                'font-style: ... underline-single | font-style: ... underline-double | font-style: ... underline-doubleaccount
                For Each underlinetype In underlineTypeList
                    If _CellConfig.ContainsKey(underlinetype) Then  'si existe (?)
                        If _CellConfig(underlinetype) Then
                            _keyName = underlinetype.Split("-")(1)
                        End If
                    End If
                Next

                'resultado
                result = If(Dictionary.UnderlineTypes.ContainsKey(If(_keyName <> "", _keyName, keyValue)),
                      Dictionary.UnderlineTypes(If(_keyName <> "", _keyName, keyValue)),
                      Dictionary.UnderlineTypes("none")
                )

            End If
            '=========================================

        End If

        Return result
    End Function

    ''' <summary>
    ''' Retorna el tipo de Número de Formato | se puede personalizar o nombrar y referenciar un formato para sólo mencionarlo y traer el formato querido
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    Private Shared Function NumberFormatType(value As String) As String
        Select Case value
            Case "hora_minuto"  'configuración personalizada
                Return "hh:mm"
            Case "text"         'configuración referenciada
                Return "@"
            Case "@"            'configuración excluida
                Return "@"
            Case Else
                Return value  'formato personalizado en un string. Ej. "$##.##"
        End Select
    End Function

    ''' <summary>
    ''' Retorna el tipo de relleno para la celda
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="type"></param>
    ''' <returns></returns>
    Private Shared Function FillType(value As String, Optional type As String = "color-palette") As Object
        Dim result As Object = Nothing

        If type = "color-palette" Then
            result = If(Dictionary.PaletteColorTypes.ContainsKey(value),
                        Dictionary.PaletteColorTypes(value),
                        Dictionary.PaletteColorTypes("none")        'Default
            )
        ElseIf type = "color-hex" Then
            result = New ColorConverter().ConvertFromString(value)
        ElseIf type = "color-rgb" Then
            Dim regx As String = "[0-9]+"
            Dim R As Short = 0
            Dim G As Short = 0
            Dim B As Short = 0
            Dim colectionRGB As MatchCollection = Regex.Matches(value, regx) 'Sacando los números rgb

            Dim position As Short = 1
            For Each number In colectionRGB
                If position = 1 Then
                    R = number.Value
                ElseIf position = 2 Then
                    G = number.Value
                ElseIf position = 3 Then
                    B = number.Value
                End If

                position += 1
            Next

            result = RGB(R, G, B)
        End If

        Return result
    End Function

    Private Shared Function GetOptionConfigAvailableValue(ByVal key As String, ByVal valueList() As String) As String
        Dim OptionAvailable As String = ""

        If key = "font-family" Then
            OptionAvailable = GetOptionAvailable_FontFamily(valueList)
        End If

        GetOptionConfigAvailableValue = ""
    End Function

    ''' <summary>
    ''' Obtiene la fuente de letra que esté disponible
    ''' Ejemplo: Helvetica, San serif, Arial(disponible)        ::: retornará Arial porque es la siguiente opción que sí está disponible
    ''' </summary>
    ''' <param name="FontOptions"></param>
    ''' <returns></returns>
    Private Shared Function GetOptionAvailable_FontFamily(ByVal FontOptions() As String) As String
        'TODO code here! (incomplete)
        GetOptionAvailable_FontFamily = "Arial"
    End Function

    ''' <summary>
    ''' Obtiene el valor por defecto del keyConfig
    ''' </summary>
    ''' <param name="keyConfig"></param>
    ''' <returns></returns>
    Private Shared Function GetKeyConfigDefaultValue(keyConfig As String, key As String) As String
        Dim result As String = Nothing
        'UNDERLINE
        If key = "font-style" Then
            If keyConfig = "underline" Then
                result = "single"
            Else
                result = True
            End If
        End If

        Return result
    End Function

    Private Shared Sub LexicalKey(key As String, value As String, _CellConfig As Dictionary(Of String, String), ConfigStringsDictionary As Dictionary(Of String, String))
        Dim valueKeyConfig() As String = Regex.Split(value, "\s+").ToArray()
        Dim totalParameters As Short = valueKeyConfig.Length

        If key = "border-width" Then
            If totalParameters = 1 Then
                _CellConfig("border-top-width") = valueKeyConfig(0)
                _CellConfig("border-right-width") = valueKeyConfig(0)
                _CellConfig("border-bottom-width") = valueKeyConfig(0)
                _CellConfig("border-left-width") = valueKeyConfig(0)
            ElseIf totalParameters = 2 Then
                _CellConfig("border-top-width") = valueKeyConfig(0)
                _CellConfig("border-right-width") = valueKeyConfig(1)
                _CellConfig("border-bottom-width") = valueKeyConfig(0)
                _CellConfig("border-left-width") = valueKeyConfig(1)
            ElseIf totalParameters = 3 Then
                _CellConfig("border-top-width") = valueKeyConfig(0)
                _CellConfig("border-right-width") = valueKeyConfig(1)
                _CellConfig("border-bottom-width") = valueKeyConfig(2)
                _CellConfig("border-left-width") = valueKeyConfig(1)
            ElseIf totalParameters = 4 Then
                _CellConfig("border-top-width") = valueKeyConfig(0)
                _CellConfig("border-right-width") = valueKeyConfig(1)
                _CellConfig("border-bottom-width") = valueKeyConfig(2)
                _CellConfig("border-left-width") = valueKeyConfig(3)
            Else
                Console.WriteLine($"{key} has too many parameters")
            End If
        ElseIf key = "border-style" Then
            If totalParameters = 1 Then
                _CellConfig("border-top-style") = valueKeyConfig(0)
                _CellConfig("border-right-style") = valueKeyConfig(0)
                _CellConfig("border-bottom-style") = valueKeyConfig(0)
                _CellConfig("border-left-style") = valueKeyConfig(0)
            ElseIf totalParameters = 2 Then
                _CellConfig("border-top-style") = valueKeyConfig(0)
                _CellConfig("border-right-style") = valueKeyConfig(1)
                _CellConfig("border-bottom-style") = valueKeyConfig(0)
                _CellConfig("border-left-style") = valueKeyConfig(1)
            ElseIf totalParameters = 3 Then
                _CellConfig("border-top-style") = valueKeyConfig(0)
                _CellConfig("border-right-style") = valueKeyConfig(1)
                _CellConfig("border-bottom-style") = valueKeyConfig(2)
                _CellConfig("border-left-style") = valueKeyConfig(1)
            ElseIf totalParameters = 4 Then
                _CellConfig("border-top-style") = valueKeyConfig(0)
                _CellConfig("border-right-style") = valueKeyConfig(1)
                _CellConfig("border-bottom-style") = valueKeyConfig(2)
                _CellConfig("border-left-style") = valueKeyConfig(3)
            Else
                Console.WriteLine($"{key} has too many parameters")
            End If
        ElseIf key = "border-color" Then
            If totalParameters = 1 Then
                _CellConfig("border-top-color") = GetStringSustitution(valueKeyConfig(0), ConfigStringsDictionary)
                _CellConfig("border-right-color") = GetStringSustitution(valueKeyConfig(0), ConfigStringsDictionary)
                _CellConfig("border-bottom-color") = GetStringSustitution(valueKeyConfig(0), ConfigStringsDictionary)
                _CellConfig("border-left-color") = GetStringSustitution(valueKeyConfig(0), ConfigStringsDictionary)
            ElseIf totalParameters = 2 Then
                _CellConfig("border-top-color") = GetStringSustitution(valueKeyConfig(0), ConfigStringsDictionary)
                _CellConfig("border-right-color") = GetStringSustitution(valueKeyConfig(1), ConfigStringsDictionary)
                _CellConfig("border-bottom-color") = GetStringSustitution(valueKeyConfig(0), ConfigStringsDictionary)
                _CellConfig("border-left-color") = GetStringSustitution(valueKeyConfig(1), ConfigStringsDictionary)
            ElseIf totalParameters = 3 Then
                _CellConfig("border-top-color") = GetStringSustitution(valueKeyConfig(0), ConfigStringsDictionary)
                _CellConfig("border-right-color") = GetStringSustitution(valueKeyConfig(1), ConfigStringsDictionary)
                _CellConfig("border-bottom-color") = GetStringSustitution(valueKeyConfig(2), ConfigStringsDictionary)
                _CellConfig("border-left-color") = GetStringSustitution(valueKeyConfig(1), ConfigStringsDictionary)
            ElseIf totalParameters = 4 Then
                _CellConfig("border-top-color") = GetStringSustitution(valueKeyConfig(0), ConfigStringsDictionary)
                _CellConfig("border-right-color") = GetStringSustitution(valueKeyConfig(1), ConfigStringsDictionary)
                _CellConfig("border-bottom-color") = GetStringSustitution(valueKeyConfig(2), ConfigStringsDictionary)
                _CellConfig("border-left-color") = GetStringSustitution(valueKeyConfig(3), ConfigStringsDictionary)
            Else
                Console.WriteLine($"{key} has too many parameters")
            End If
        ElseIf {"border-top", "border-right", "border-bottom", "border-left"}.Contains(key) Then
            If totalParameters >= 1 And totalParameters <= 3 Then
                Dim repeatWidth, repeatStyle, repeatColor As Boolean

                For i = 0 To valueKeyConfig.Length - 1
                    Dim _valueKeyConfig = GetStringSustitution(valueKeyConfig(i), ConfigStringsDictionary)
                    Dim type As String = GetBorderParameter(_valueKeyConfig)
                    If type = "width" Then
                        If Not repeatWidth Then
                            repeatWidth = True
                            _CellConfig($"{key}-width") = _valueKeyConfig
                        Else
                            Console.WriteLine($"{key} {type} {_valueKeyConfig} is repeated")
                        End If
                    ElseIf type = "style" Then
                        If Not repeatStyle Then
                            repeatStyle = True
                            _CellConfig($"{key}-style") = _valueKeyConfig
                        Else
                            Console.WriteLine($"{key} {type} {_valueKeyConfig} is repeated")
                        End If
                    ElseIf type = "color" Then
                        If Not repeatColor Then
                            repeatColor = True
                            _CellConfig($"{key}-color") = _valueKeyConfig
                        Else
                            Console.WriteLine($"{key} {type} {_valueKeyConfig} is repeated")
                        End If
                    End If
                Next
            Else
                Console.WriteLine($"{key} too much parameters")
            End If
        ElseIf key = "border" Then
            If totalParameters >= 1 And totalParameters <= 3 Then
                For Each _valueKeyConfig In valueKeyConfig
                    _valueKeyConfig = GetStringSustitution(_valueKeyConfig, ConfigStringsDictionary)
                    Dim type As String = GetBorderParameter(_valueKeyConfig)
                    _CellConfig($"{key}-{type}") = _valueKeyConfig
                Next
            Else
                Console.WriteLine($"{key} too much parameters")
            End If
        Else
            LexicalSeparator(key, value, _CellConfig, ConfigStringsDictionary)
        End If

        'Seteamos todas las configuraciones key: value
        If _CellConfig.ContainsKey(key) Then
            value = If(ConfigStringsDictionary.ContainsKey(value), ConfigStringsDictionary(value), value)
            value = Replace(Replace(value, """", ""), "'", "")
            _CellConfig(key) = value
        Else
            Console.WriteLine($"key not found: {key}")
        End If

    End Sub

    Private Shared Function GetStringSustitution(value As String, ConfigStringsDictionary As Dictionary(Of String, String))
        Return If(ConfigStringsDictionary.ContainsKey(value),
                    ConfigStringsDictionary(value),
                    value
            )
    End Function

    Private Shared Sub LexicalSeparator(key As String, value As String, _CellConfig As Dictionary(Of String, String), ConfigStringsDictionary As Dictionary(Of String, String))
        'Primera validación(MULTIPLES OPCIONES) en value (¿hay comas?, si hay entonces ... split(","))
        If value.Contains(",") Then

            '¿hay espacios en blanco? => borrarlos, porque o es multiple opción o es multiple configuración, no puedes servir a dos señores
            value = Replace(value, " ", "")

            Dim valueKeyOptions() As String = value.Split(",")
            Dim _valueKeyOptions() As String = Array.Empty(Of String)()
            Dim iko As Integer = 0      'iterador para keyOption
            'Limpiando las opciones
            For Each keyOption In valueKeyOptions
                keyOption = Trim(keyOption)
                keyOption = If(ConfigStringsDictionary.ContainsKey(keyOption), ConfigStringsDictionary(keyOption), keyOption)
                keyOption = Replace(Replace(keyOption, """", ""), "'", "")            'esta cadena de texto no soporta " (comillas dobles) ni '(comillas sencillas) porque hacemos un replace sobre ella, en caso que se necesite, edítalo

                'Aquí va el procedimiento par las MULTIPLES OPCIONES (hasta ahora no se ha utilizado esta funcionalidad sin embargo dejamos un ejemplo)
                'Ejemplo:
                'font-family: 'Helvetica', 'San serif', 'Arial'
                'Si no encuentra la fuente de letra Helvetica, que le ponga San serif, y si no se encuentra San serif, entonces ponle Arial, ya qué :/

                _valueKeyOptions(iko) = keyOption

                iko += 1
            Next

            'Si son muchas opciones => conseguir la opcion más cercana y disponible
            If _CellConfig.ContainsKey(key) And _valueKeyOptions.Length > 1 Then
                value = GetOptionConfigAvailableValue(key, _valueKeyOptions)
            Else
                Console.WriteLine($"key not found for validate (keyOption): {key}")
            End If
        End If

        'Segunda validación(MULTIPLES CONFIGURACIONES) en value (¿hay espacios en blanco?, si hay entonces ... split(" "))
        If value.Contains(" ") Then
            Dim valueKeyConfig() As String = Regex.Split(value, "\s+").ToArray()

            'Configuraciones individuales
            For Each keyConfig In valueKeyConfig
                keyConfig = Trim(keyConfig)
                keyConfig = If(ConfigStringsDictionary.ContainsKey(keyConfig), ConfigStringsDictionary(keyConfig), keyConfig)
                keyConfig = Replace(Replace(keyConfig, """", ""), "'", "")            'esta cadena de texto no soporta " (comillas dobles) ni '(comillas sencillas) porque hacemos un replace sobre ella, en caso que se necesite, edítalo

                'Hacer aquí el procedimiento de MULTIPLES CONFIGURACIONES (éstas keyConfig mayormente son True, False)
                If _CellConfig.ContainsKey(keyConfig) Then
                    _CellConfig(keyConfig) = GetKeyConfigDefaultValue(keyConfig, key)
                Else
                    Console.WriteLine($"keyConfig not found: {keyConfig}")
                End If
            Next

        End If
    End Sub

    Private Shared Function GetBorderParameter(valueKeyConfig As String) As Object
        Dim parameter As String = ""
        Dim regxARGB As String = "^(a?rgb)?\(?([01]?\d\d?|2[0-4]\d|25[0-5])(\W+)([01]?\d\d?|2[0-4]\d|25[0-5])\W+(([01]?\d\d?|2[0-4]\d|25[0-5])\)?)$"

        If Dictionary.BorderWeightEnumerations.ContainsKey(valueKeyConfig) Then
            parameter = "width"
        ElseIf Dictionary.BorderTypes.ContainsKey(valueKeyConfig) Then
            parameter = "style"
        ElseIf Dictionary.PaletteColorTypes.ContainsKey(valueKeyConfig) Then
            parameter = "color"
        ElseIf valueKeyConfig.IndexOf("#") = 0 Then
            parameter = "color"
        ElseIf Regex.IsMatch(valueKeyConfig, regxARGB) Then
            parameter = "color"
        End If

        Return parameter
    End Function

#End Region
#Region "OTRAS CLASES"
    ''' <summary>
    ''' DataWorksheet - Datos de la Hoja de Trabajo
    ''' </summary>
    Public Class DataWorksheet
        Public PositionRow As Short     'Posición de la fila en uso
        Public SheetName As String      'Nombre de la hoja
    End Class
#End Region
End Class
