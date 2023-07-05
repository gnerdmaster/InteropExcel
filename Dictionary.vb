Imports Microsoft.Office.Interop.Excel.XlColorIndex
'Imports Microsoft.Office.Interop.Excel.XlRgbColor 'pendiente       'rgbNameColor
'Imports Microsoft.Office.Interop.Excel.XlStdColorScale ' pendiente     'scaleColor example. whiteblack blackwhite
Imports Microsoft.Office.Interop.Excel.XlBorderWeight
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.XlUnderlineStyle
Imports Microsoft.Office.Interop.Excel.XlHAlign
Imports Microsoft.Office.Interop.Excel.XlVAlign
Public Class Dictionary
#Region "Constantes"
    'Info extra: https://learn.microsoft.com/en-us/office/vba/api/excel.constants
    'FONTS
    'Const xlUnderlineStyleDouble As Short = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleDouble
    'Const xlUnderlineStyleDoubleAccounting As Short = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleDoubleAccounting
    'Const xlUnderlineStyleNone As Short = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleNone
    'Const xlUnderlineStyleSingle As Short = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle

    'ALIGNMENT
    'Const xlHAlignLeft As Short = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
    'Const xlHAlignRight As Short = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
    'Const xlHAlignCenter As Short = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
    'Const xlHAlignJustify As Short = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignJustify
    'Const xlHAlignFill As Short = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignFill
    'Const xlVAlignTop As Short = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop
    'Const xlVAlignBottom As Short = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom
    'Const xlVAlignCenter As Short = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

    'BORDERS
    'Const xlEdgeLeft As Short = Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft
    'Const xlEdgeRight As Short = Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight
    'Const xlEdgeTop As Short = Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop
    'Const xlEdgeBottom As Short = Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom
    'Const xlInsideHorizontal As Short = Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal
    'Const xlInsideVertical As Short = Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical
    '---
    'Const xlContinuous As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
    'Const xlDash As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlDash
    'Const xlDashDot As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlDashDot
    'Const xlDashDotDot As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlDashDotDot
    'Const xlDot As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot
    'Const xlDouble As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
    'Const xlLineStyleNone As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone
    'Const xlSlantDashDot As Short = Microsoft.Office.Interop.Excel.XlLineStyle.xlSlantDashDot
#End Region

#Region "Diccionarios - CELDA (PRINCIPAL)"
    ''' <summary>
    ''' (DICCIONARIO DE CONFIGURACIONES DE CELDA) AÑADIR AQUÍ LAS NUEVAS CONFIGURACIONES Y SUS VALORES POR DEFECTO)
    ''' Ejemplo llamada desde otra clase: 
    '''     Dim _CellConfig As Dictionary(Of String, String) = clsDictionary.CellConfigurations
    ''' **Crea un nuevo diccionario con las mismas configuración ya predefinidas aquí para cada celda que lo use.
    ''' </summary>
    Public Shared ReadOnly Property CellConfigurations() As Dictionary(Of String, String)
        'Se usa propiedades como la mejor opción de llamado a las configuraciones predeterminadas: https://learn.microsoft.com/en-us/dotnet/visual-basic/language-reference/statements/get-statement
        Get
            CellConfigurations = New Dictionary(Of String, String) From {
                {"border", "none"},
                {"border-top", "none"},        'BORDER - [<width-value> | <style-value> | <color-value>] - https://developer.mozilla.org/es/docs/Web/CSS/border-top
                {"border-right", "none"},      'BORDER - [<width-value> | <style-value> | <color-value>] - https://developer.mozilla.org/es/docs/Web/CSS/border-right
                {"border-bottom", "none"},     'BORDER - [<width-value> | <style-value> | <color-value>] - https://developer.mozilla.org/es/docs/Web/CSS/border-bottom
                {"border-left", "none"},       'BORDER - [<width-value> | <style-value> | <color-value>] - https://developer.mozilla.org/es/docs/Web/CSS/border-left
                {"border-inside-horizontal", "none"},
                {"border-inside-vertical", "none"},
                {"border-width", "none"},               'BORDER WIDTH- Todos los bordes [thin, medium, thick, (length not soported)]
                {"border-top-width", "none"},           'BORDER WIDTH- [thin | medium | thick | (length not soported)]
                {"border-right-width", "none"},         'BORDER WIDTH- [thin | medium | thick | (length not soported)]
                {"border-bottom-width", "none"},        'BORDER WIDTH- [thin | medium | thick | (length not soported)]
                {"border-left-width", "none"},          'BORDER WIDTH- [thin | medium | thick | (length not soported)]
                {"border-style", "none"},               'BORDER STYLE- [continuous(solid) | dash(dashed) | dashdot | dashdotdot | dot(dotted) | double | linestylenone(none) | slantdashdot]
                {"border-top-style", "none"},           'BORDER STYLE- [continuous(solid) | dash(dashed) | dashdot | dashdotdot | dot(dotted) | double | linestylenone(none) | slantdashdot]
                {"border-right-style", "none"},         'BORDER STYLE- [continuous(solid) | dash(dashed) | dashdot | dashdotdot | dot(dotted) | double | linestylenone(none) | slantdashdot]
                {"border-bottom-style", "none"},        'BORDER STYLE- [continuous(solid) | dash(dashed) | dashdot | dashdotdot | dot(dotted) | double | linestylenone(none) | slantdashdot]
                {"border-left-style", "none"},          'BORDER STYLE- [continuous(solid) | dash(dashed) | dashdot | dashdotdot | dot(dotted) | double | linestylenone(none) | slantdashdot]
                {"border-color", "none"},                       'BORDER COLOR- (namecolor | hexdecimal | rgb) - examples [green, white, black, blue, #fff, #ffffff, rgb(255,255,255)]
                {"border-top-color", "none"},                   'BORDER COLOR- (namecolor | hexdecimal | rgb) - examples [green, white, black, blue, #fff, #ffffff, rgb(255,255,255)]
                {"border-right-color", "none"},                 'BORDER COLOR- (namecolor | hexdecimal | rgb) - examples [green, white, black, blue, #fff, #ffffff, rgb(255,255,255)]
                {"border-bottom-color", "none"},                'BORDER COLOR- (namecolor | hexdecimal | rgb) - examples [green, white, black, blue, #fff, #ffffff, rgb(255,255,255)]
                {"border-left-color", "none"},                  'BORDER COLOR- (namecolor | hexdecimal | rgb) - examples [green, white, black, blue, #fff, #ffffff, rgb(255,255,255)]
                {"border-inside-horizontal-color", "none"},     'BORDER COLOR - (namecolor | hexdecimal | rgb) - examples [green, white, black, blue, #fff, #ffffff, rgb(255,255,255)]
                {"border-inside-vertical-color", "none"},       'BORDER COLOR - (namecolor | hexdecimal | rgb) - examples [green, white, black, blue, #fff, #ffffff, rgb(255,255,255)]
                {"font-size", 12},                      'FONT - Tamaño de letra en número
                {"font-family", "Calibri"},             'FONT - [Arial, San Serif, Helvetica, Calibri, ...] (múltiples opciones, con orden de existencia, separados por una coma)
                {"font-style", "normal"},               'FONT - [normal | bold | italic | underline | underline-single | underline-double | underline-doubleaccount, strikethrough (line-through)] (multiples configuraciones separados por un espacio; a diferencia de css, éste maneja otros valores clave)        https://developer.mozilla.org/en-US/docs/Web/CSS/font-style
                {"color", "none"},                      'FONT - (namecolor, hexdecimal format, rgb format) - examples [green, white, black, blue, #fff, #ffffff, rgb(255,255,255)]
                {"italic", False},                      'FONT - [true | false]
                {"bold", False},                        'FONT - [true | false]
                {"underline", "none"},                  'FONT - [none | double | doubleaccount | single(solid) ]
                {"underline-single", False},            'FONT UNDERLINE - [true | false]
                {"underline-double", False},            'FONT UNDERLINE - [true | false]
                {"underline-doubleaccount", False},     'FONT UNDERLINE - [true | false]
                {"strikethrough", False},               'FONT UNDERLINE - [true | false]    'Tachado
                {"text-decoration-style", "none"},      'FONT UNDERLINE - [none | double | doubleaccount | single(solid) ] https://developer.mozilla.org/en-US/docs/Web/CSS/text-decoration-style
                {"text-decoration-line", "none"},       'FONT UNDERLINE - [none | underline | strikethrough(line-through) ]  https://developer.mozilla.org/en-US/docs/Web/CSS/text-decoration-line
                {"text-transform", "none"},             'FONT TRANSFORM - [none | uppercase | lowercase | capitalize]
                {"shrink-to-fit", False},               'ALIGNMENT - [true | false} (Reducir hasta ajustar)
                {"text-align", "start"},                'ALIGNMENT - [none, start, end, center, justify]
                {"vertical-align", "top"},              'ALIGNMENT - [none, top, middle, bottom]
                {"text-wrap", False},                   'ALIGNMENT - [true | false]
                {"number-format", "General"},           'NUMBER FORMAT - [@ = text, General, Estándar, 'for:mat0 personalizad0']    Otros formatos, consultar este link: https://support.microsoft.com/es-es/office/c%C3%B3digos-de-formato-de-n%C3%BAmero-5026bbd6-04bc-48cd-bf33-80f18b4eae68?ui=es-es&rs=es-hn&ad=us
                {"background-color", "none"}            'FILL - (namecolor | hexdecimal | rgb) - examples [green, white, black, blue, #fff, #ffffff, rgb(255,255,255)]
            }
            Exit Property
        End Get
    End Property

#End Region

#Region "Diccionarios - BORDERS"
    ''' <summary>
    ''' Diccionario de datos para los Tipos de Bordes
    ''' Info: https://learn.microsoft.com/es-es/dotnet/api/microsoft.office.interop.excel.border.linestyle?view=excel-pia
    ''' </summary>
    Public Shared BorderStyles As New Dictionary(Of String, Short) From {
            {"continuous", xlContinuous},
            {"solid", xlContinuous},        'css
            {"dash", xlDash},
            {"dashed", xlDash},             'css
            {"dashdot", xlDashDot},
            {"dashdotdot", xlDashDotDot},
            {"dot", xlDot},
            {"dotted", xlDot},              'css
            {"double", xlDouble},
            {"slantdashdot", xlSlantDashDot},
            {"linestylenone", xlLineStyleNone},     'none
            {"none", xlLineStyleNone}               'default
    }

    ''' <summary>
    ''' Diccionario de Nombre y Numeraciones de Tipos de Bordes
    ''' Info: https://learn.microsoft.com/en-us/office/vba/api/excel.xlbordersindex
    ''' </summary>
    Public Shared ReadOnly Property NameAndEnumerationBorders() As Dictionary(Of Short, String)
        Get
            NameAndEnumerationBorders = New Dictionary(Of Short, String) From {
                {xlEdgeTop, "border-top"},
                {xlEdgeRight, "border-right"},
                {xlEdgeBottom, "border-bottom"},
                {xlEdgeLeft, "border-left"},
                {xlInsideHorizontal, "border-inside-horizontal"},
                {xlInsideVertical, "border-inside-vertical"}
            }
            Exit Property
        End Get
    End Property

    ''' <summary>
    ''' Diccionario de enumeración de anchura de borde
    ''' Info: https://learn.microsoft.com/en-us/office/vba/api/excel.xlborderweight
    ''' </summary>
    Public Shared BorderWeightEnumerations As New Dictionary(Of String, Short) From {
            {"hairline", xlHairline},
            {"medium", xlMedium},
            {"thick", xlThick},     '(widest border)
            {"thin", xlThin}
    }
#End Region

#Region "Diccionarios - FONTS"

    ''' <summary>
    ''' Diccionario de datos para los Tipos de Subrayados
    ''' Info: https://learn.microsoft.com/en-us/office/vba/api/excel.xlunderlinestyle
    ''' </summary>
    Public Shared UnderlineStyles As New Dictionary(Of String, Short) From {
            {"double", xlUnderlineStyleDouble},
            {"doubleaccount", xlUnderlineStyleDoubleAccounting},
            {"single", xlUnderlineStyleSingle},
            {"solid", xlUnderlineStyleSingle},      'css
            {"none", xlUnderlineStyleNone}          'default
    }
#End Region

#Region "Diccionarios - ALIGMENTS"
    ''' <summary>
    ''' Diccionario de datos para los Tipos de Alineamiento de Texto
    ''' Info: https://learn.microsoft.com/en-us/office/vba/api/excel.xlhalign
    ''' </summary>
    Public Shared AlignmentTypes_Text As New Dictionary(Of String, Short) From {
        {"start", xlHAlignLeft},
        {"end", xlHAlignRight},
        {"center", xlHAlignCenter},
        {"justify", xlHAlignJustify},
        {"fill", xlHAlignFill},
        {"none", xlHAlignLeft}    'default
    }

    ''' <summary>
    ''' Diccionario de datos para los Tipos de Alineamientos en Vertical
    ''' Info: https://learn.microsoft.com/en-us/office/vba/api/excel.xlvalign
    ''' </summary>
    Public Shared AlignmentTypes_Vertical As New Dictionary(Of String, Short) From {
        {"top", xlVAlignTop},
        {"middle", xlVAlignCenter},
        {"bottom", xlVAlignBottom},
        {"none", xlVAlignBottom}      'default
    }
#End Region

#Region "Diccionarios - FILL"
    ''' <summary>
    ''' Info: https://learn.microsoft.com/en-US/dotnet/api/microsoft.office.interop.excel.xlcolorindex?view=excel-pia
    ''' Info: http://dmcritchie.mvps.org/excel/colors.htm
    ''' Info: https://learn.microsoft.com/en-us/office/vba/api/excel.colorindex
    ''' Info: https://social.msdn.microsoft.com/Forums/es-ES/6e2aabfc-0ca3-4577-8055-8ea81078f2c8/setting-the-fill-color-of-an-excel-cell-using-vb-and-microsoftofficeinterop?forum=vbinterop
    ''' </summary>
    Public Shared PaletteColorTypes As New Dictionary(Of String, Short) From {
        {"aqua", 42},
        {"black", 1},
        {"blue", 5},
        {"bluegray", 47},
        {"brightgreen", 4},
        {"brown", 53},
        {"cream", 19},
        {"darkblue", 11},
        {"darkgreen", 51},
        {"darkpurple", 21},
        {"darkred", 9},
        {"darkteal", 49},
        {"darkyellow", 12},
        {"gold", 44},
        {"gray25", 15},
        {"gray40", 48},
        {"gray50", 16},
        {"gray80", 56},
        {"green", 10},
        {"indigo", 55},
        {"lavender", 39},
        {"lightblue", 41},
        {"lightgreen", 35},
        {"lightlavender", 24},
        {"lightorange", 45},
        {"lightturquoise", 20},
        {"lightyellow", 36},
        {"lime", 43},
        {"navyblue", 23},
        {"olivegreen", 52},
        {"orange", 46},
        {"paleblue", 37},
        {"pink", 7},
        {"plum", 18},
        {"powderblue", 17},
        {"red", 3},
        {"rose", 38},
        {"salmon", 22},
        {"seagreen", 50},
        {"skyblue", 33},
        {"tan", 40},
        {"teal", 14},
        {"turquoise", 8},
        {"violet", 13},
        {"white", 2},
        {"yellow", 6},
        {"automatic", xlColorIndexAutomatic},
        {"none", xlColorIndexNone}
    }
#End Region

#Region "Diccionarios-MIXED"
    ''' <summary>
    ''' Diccionario de equivalencias de palabras clave
    ''' key = key used
    ''' value = equivalence
    ''' </summary>
    Public Shared KeyEquivalences As New Dictionary(Of String, String) From {
        {"underline", "underline-single"},
        {"line-through", "strikethrough"},
        {"horizontal-align", "text-align"},
        {"shrink", "shrink-to-fit"},
        {"shrinktofit", "shrink-to-fit"}
    }
#End Region
End Class
