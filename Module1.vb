Module Module1

    Sub Main()
        Dim ExcelApp = New Microsoft.Office.Interop.Excel.Application
        Dim Book = ExcelApp.Workbooks.Add
        Dim infoSheet As New Cell.DataWorksheet
        infoSheet.PositionRow = 1
        infoSheet.SheetName = "Hoja1"

        'Libro.Sheets(1).Cells(1, 1).RichText.Substring(startindex, noofchar).SetFontColor(XLColor.Red) ' = "Hola Mundo" 'https://www.codeproject.com/Questions/1130851/How-to-make-different-colors-of-texts-in-same-exce    'Ejemplo básico


        Dim C As New Cell
        C.XlsApp = ExcelApp
        C.infoSheet = infoSheet

        ''------------COMMON FORMATS (PRINCIPALES)---------
        C.Cell("Sin formato", 2, C.N("B"), 2, C.N("C")) 'FORMA 1 de crear una celda

        C.Cell("Texto", "B3", "C3", "number-format:@") 'FORMA 2 de crear una celda
        C.Cell("01/01/2020", "B4", "number-format:'yyyy/mm/dd'") 'Formato fecha     'FORMA 3 de crear una celda
        C.Cell("5123", "B5:C5", "number-format:'??=??'")           'Formato numérico

        C.Cell("MULTIPLES OPCIONES Y KEYCONFIGVALUE POR DEFECTO", "B6:C6", "font-style: bold underline  italic")     'se prueba si hay muchos espacios entre keyConfigs y también la configuración por defecto de un keyConfig que tienen varias configuraciones (underline)

        C.Cell("EJEMPLO DE REDUCCIÓN DE CÓDIGO EN BORDES", "B7:D7", "border-left:dot; border-right:continuous; border-bottom: dashdotdot")

        ''----------ALINEACIONES------------
        C.Cell("TEST ALIGNMENT", "B9:E9", "text-align: center; vertical-align:none")

        ''----------FÓRMULAS----------------
        C.Cell("=CONCAT(""TEST"", "" PARA UNA FUNCIÓN"")", "B11")       'fórmula CONCAT sin formato y en una sola celda
        C.Cell("=CONCAT(""TEST"", "" PARA UNA FUNCIÓN EN CELDA CONVINADA"")", "B12:D14")       'fórmula CONCAT sin formato y en una celda convinada con otras

        Dim wraptext As String = "WrapText - Este es un test de un texto demasiado largo, y tiene que pasar la prueba de Wrap text = Ajustar texto"
        Dim wraptextConfig As String = "text-wrap:true"
        C.Cell(wraptext, "B16:D18", wraptextConfig)

        Dim textjustify As String = "Text Justify - Prueba de justificación de texto, tenemos que asegurarnos por si no jala, entonces tendríamos que volver a programar lo que ya programamos en lo programado en la programación de la redundación y la atareación"
        C.Cell(textjustify, "B20:D22", "text-align:justify")

        ''------------MERGE SPECIAL
        C.Cell("Merge F:F", "F:F")

        ''------------BORDERS---------
        C.Cell("border: continuous", "G2", "border: continuous")    'todos los bordes y una configuración
        C.Cell("border-left: double", "G3", "border-left: double")  'borde izquierdo
        C.Cell("border-top: dot", "G4", "border-top: dot")          'borde superior
        C.Cell("border-right: dash", "G5", "border-right: dash")    'borde derecho
        C.Cell("border-bottom: dashdot", "G6", "border-bottom: dashdot")          'borde inferior
        ''------------COLORES DE FONDO------------------
        C.Cell("background-color: yellow", "I1", "background-color: yellow")
        C.Cell("background-color: #A0F0D0", "I2", "background-color: #A0F0D0")
        C.Cell("background-color: rgb(50,10,200)", "I3", "background-color: rgb(50,10,200)")
        C.Cell("background-color: #5AA456", "I4", "background-color: #5AA456")

        ''-----------UNDERLINE-------------------
        C.Cell("underline: doubleaccount", "I6", "underline: doubleaccount")
        C.Cell("font-style: underline", "I7", "font-style: underline")
        C.Cell("font-style: underline-double", "I8", "font-style: underline-double")
        C.Cell("underline: double", "I9", "underline: double")
        C.Cell("font-style: underline italic", "I10", "font-style: underline italic bold")
        C.Cell("font-style: underline-double bold italic", "I11", "font-style: underline-double bold italic")

        '------------ONLY FORMAT WITHIN VALUE | MERGE OR NOT MERGE RANGE
        '' <> Not Merge within inside borders (no se muestra ningún borde interno aunque tenga configuración personalizada)
        '' <i> Not Merge with inside borders (si no tiene declarado el tipo de borde toma por default los de border-bottom(border-inside-horizontal) y border-right(border-inside-vertical))
        C.Cell(Nothing, "J1:J5", "border-right: continuous; border-left: dot;background-color: pink; font-style: underline-double bold italic")                                                 'formateo con convinación de celdas
        C.Cell(Nothing, "<>J6:K11", "border: dashdotdot; background-color: yellow; font-style: underline italic")                                                                               'formateo sin convinación de celdas
        C.Cell(Nothing, "<>J16:K20", "border-inside-horizontal: continuos; border-inside-vertical: dot;border: double; background-color: blue; font-style: underline-single")                   'formateo sin convinación de celdas, bordes interiores configurados pero no mostrados(ignorados)
        C.Cell(Nothing, "<i>J12:K15", "border: dash; background-color: #F0F0F0; font-style: underline-double bold")                                                                             'formateo sin convinación de celdas y border interiores por default(no declarados en las configuraciones)
        C.Cell(Nothing, "<i>J21:L24", "border-inside-horizontal: continuous; border-inside-vertical: dot;border: double; background-color: red; font-style: underline-doubleaccount")           'formateo sin convinación de celdas y bordes interiores configurados y mostrados

        ''-----------COLORES DE FUENTE DE LETRA-------------
        C.Cell("color: red", "M2", "color: red")
        C.Cell("color: #0E2FC3", "M3", "color: #0E2FC3")
        C.Cell("color: rgb(174, 241, 71)", "M4", "color: rgb(174, 241, 71)")
        ''-----------RICH TEXT (HTML EN LA CELDA)----------
        'C.Cell("<h>Hola mundo</h><u>underline</u> HTML en <b>TODOS lados</b>", "L2:M2") --pendiente
        ''-----------CELL VERSIÓN 4 (CREACIÓN DE UNA TABLA)




        'Siguientes mejoras:
        '1.- Bordes de colores
        '3.- HTML dentro de la celda
        '4.- Controlar el grosor de los bordes
        '5.- Orientación de Texto
        '6.- Sangría de Texto
        '7.- Reducir Hasta ajustar(texto)
        '8.- (Fuente) Efecto tachado 
        '9.- (Fuente) Efecto Superíndice
        '10.- (Fuente) Efecto Subíndice
        '11.- (Color Fondo) Efectos de relleno, Color de Trama, Estilo de Trama
        '12.- Sección Proteger (Bloqueada, Oculta)


        'Siguientes propuestas
        '1.- Dividir texto en varias fracciones de celda (para evitar textos extensos y que no quepan en una celda de excel)
        '2.- Hacer un procedimiento para creación de tablas preconfiguradas(desde títulos de columnas hasta la información que pinte desde una fuente de datos)
        '2.1.- Hacer que esta tabla pintada sea una tipo tabla que se puedan configurar filtros, formatos de pintado, tablas dinámicas, funciones, etc.
        '3.- Dar formato a nivel caracteres(ej. caracteres en diferentes colores y/o fuentes de letras)
        '4.- Programar una macro XD
        '5.- Insertar gráficos
        '6.- Insertar Ilustraciones (Imágenes, Formas, Íconos, Modelos 3D, etc.)
        '7.- Controlar las diposiciones de la página
        '8.- Opciones de Guardado
        '9.- Temas de tablas y celdas



        Book.SaveAs(Filename:="..\Downloads\InteropExcel\Exports\test01.xlsx")  'Cambiar ruta si es necesario


        ExcelApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic
        ExcelApp.ScreenUpdating = True
        ExcelApp.Visible = True
        ExcelApp.DisplayAlerts = True

        ExcelApp.Quit()

        Book = Nothing
        ExcelApp = Nothing
    End Sub

End Module
