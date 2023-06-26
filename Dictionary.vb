Imports Microsoft.Office.Interop.Excel.XlColorIndex
Public Class Dictionary
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
                {"border", "linestylenone"},            'BORDER - [continuous, dash, dashdot, dashdotdot, dot, double, linestylenone, slantdashdot]     
                {"border-left", "linestylenone"},       'BORDER - [continuous, dash, dashdot, dashdotdot, dot, double, linestylenone, slantdashdot]
                {"border-right", "linestylenone"},      'BORDER - [continuous, dash, dashdot, dashdotdot, dot, double, linestylenone, slantdashdot]
                {"border-top", "linestylenone"},        'BORDER - [continuous, dash, dashdot, dashdotdot, dot, double, linestylenone, slantdashdot]
                {"border-bottom", "linestylenone"},     'BORDER - [continuous, dash, dashdot, dashdotdot, dot, double, linestylenone, slantdashdot]
                {"border-inside-horizontal", "linestylenone"},      'BORDER - [continuous, dash, dashdot, dashdotdot, dot, double, linestylenone, slantdashdot]
                {"border-inside-vertical", "linestylenone"},        'BORDER - [continuous, dash, dashdot, dashdotdot, dot, double, linestylenone, slantdashdot]
                {"font-size", 8},                       'FONT - Tamaño de letra en número
                {"font-family", "Arial"},               'FONT - [Arial, San Serif, Helvetica, Calibri, ...] (múltiples opciones, con orden de existencia, separados por una coma)
                {"font-style", "normal"},               'FONT - [normal, bold, italic, underline] (multiples configuraciones separados por un espacio)                     
                {"italic", False},                      'FONT - [true, false]
                {"bold", False},                        'FONT - [true, false]
                {"underline", "none"},                  'FONT - [none, double, doubleaccount, single]
                {"underline-single", False},            'FONT - [true, false]
                {"underline-double", False},            'FONT - [true, false]
                {"underline-doubleaccount", False},     'FONT - [true, false]
                {"text-align", "start"},                'ALIGNMENT - [none, start, end, center, justify]
                {"vertical-align", "top"},              'ALIGNMENT - [none, top, middle, bottom]
                {"text-wrap", False},                   'ALIGNMENT - [true, false]
                {"number-format", "General"},           'NUMBER FORMAT - [@ = text, General, Estándar, 'for:mat0 personalizad0']    Otros formatos, consultar este link: https://support.microsoft.com/es-es/office/c%C3%B3digos-de-formato-de-n%C3%BAmero-5026bbd6-04bc-48cd-bf33-80f18b4eae68?ui=es-es&rs=es-hn&ad=us
                {"background-color", "none"}            'FILL - [green, white, black, blue, #fff, #ffffff, rgb(255,255,255)]
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
    Public Shared BorderTypes As New Dictionary(Of String, Short) From {
            {"continuous", xlContinuous},
            {"dash", xlDash},
            {"dashdot", xlDashDot},
            {"dashdotdot", xlDashDotDot},
            {"dot", xlDot},
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
                {xlEdgeLeft, "border-left"},
                {xlEdgeRight, "border-right"},
                {xlEdgeTop, "border-top"},
                {xlEdgeBottom, "border-bottom"},
                {xlInsideHorizontal, "border-inside-horizontal"},
                {xlInsideVertical, "border-inside-vertical"}
            }
            Exit Property
        End Get
    End Property
#End Region

#Region "Diccionarios - FONTS"

    ''' <summary>
    ''' Diccionario de datos para los Tipos de Subrayados
    ''' Info: https://learn.microsoft.com/en-us/office/vba/api/excel.xlunderlinestyle
    ''' </summary>
    Public Shared UnderlineTypes As New Dictionary(Of String, Short) From {
            {"double", xlUnderlineStyleDouble},
            {"doubleaccount", xlUnderlineStyleDoubleAccounting},
            {"single", xlUnderlineStyleSingle},
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
End Class
