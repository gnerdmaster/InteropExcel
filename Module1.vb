Module Module1

    Sub Main()
        Dim ExcelApp = New Microsoft.Office.Interop.Excel.Application
        Dim Book = ExcelApp.Workbooks.Add
        Dim infoSheet As New Cell.DataWorksheet
        infoSheet.XlsApp = ExcelApp     '1.- manera 1 en asignar la Aplicación
        infoSheet.PositionRow = 1
        infoSheet.SheetName = "Hoja1"

        'ExcelApp.Worksheets(1).Cells(1, 1).RichText.Substring(1, 5).SetFontColor(RGB(125, 125, 125)) ' = "Hola Mundo" 'https://www.codeproject.com/Questions/1130851/How-to-make-different-colors-of-texts-in-same-exce    'Ejemplo básico

        'Instanciar la clase Cell para asignar la aplicación (XlsApp) y la información necesario de la hoja(infoSheet)
        Dim C As New Cell
        C.XlsApp = ExcelApp     '2.- manera 2 en asignar la Aplicación
        C.infoSheet = infoSheet

        Const test As Boolean = False        'False = entra en el if, True = Ignora el contenido del if
        If Not test Then


            ''------------COMMON FORMATS (PRINCIPALES)---------
            C.Cell("Sin formato", 2, C.N("B"), 2, C.N("C")) 'FORMA 1 de crear una celda

            C.Cell("Texto", "B3", "C3", "number-format:@") 'FORMA 2 de crear una celda
            C.Cell("01/01/2020", "B4", "number-format:'yyyy/mm/dd'") 'Formato fecha     'FORMA 3 de crear una celda
            C.Cell("5123", "B5:C5", "number-format:'??=??'")           'Formato numérico

            C.Cell("MULTIPLES OPCIONES Y KEYCONFIGVALUE POR DEFECTO", "B6:C6", "font-style: bold underline  italic")     'se prueba si hay muchos espacios entre keyConfigs y también la configuración por defecto de un keyConfig que tienen varias configuraciones (underline)

            C.Cell("EJEMPLO DE REDUCCIÓN DE CÓDIGO EN BORDES", "B7:D7", "border-left:dot; border-right:continuous; border-bottom: dashdotdot")

            ''----------ALINEACIONES------------
            C.Cell("TEST ALIGNMENT", "B9:E9", "text-align: center; vertical-align:none")

            ''REDUCIR TEXTO PARA AJUSTAR (shrink-to-fit)
            C.Cell("shrink-to-fit:true(encoge contenido)", "b10", "shrink-to-fit:true")
            C.Cell("shrink:true(encoge contenido)", "b11", "shrink:true")
            C.Cell("shrinktofit:true", "b12", "shrinktofit:true")

            ''----------FÓRMULAS----------------
            C.Cell("=CONCAT(""TEST"", "" PARA UNA FUNCIÓN"")", "B13")       'fórmula CONCAT sin formato y en una sola celda
            C.Cell("=CONCAT(""TEST"", "" PARA UNA FUNCIÓN EN CELDA CONVINADA"")", "B14:D14")       'fórmula CONCAT sin formato y en una celda convinada con otras

            Dim wraptext As String = "WrapText - Este es un test de un texto demasiado largo, y tiene que pasar la prueba de Wrap text = Ajustar texto"
            Dim wraptextConfig As String = "text-wrap:true"
            C.Cell(wraptext, "B15:D18", wraptextConfig)

            Dim textjustify As String = "Text Justify - Prueba de justificación de texto, tenemos que asegurarnos por si no jala, entonces tendríamos que volver a programar lo que ya programamos en lo programado en la programación de la redundación y la atareación"
            C.Cell(textjustify, "B20:D22", "text-align:justify")

            ''------------MERGE SPECIAL
            C.Cell("Merge F:F", "F:F")

            ''------------BORDERS---------
            Dim borderStyleIndividual As String = "border-top-style: continuous; border-right-style:continuous; border-bottom-style: continuous; border-left-style: double"
            C.Cell(borderStyleIndividual, "G2", borderStyleIndividual)     'border-style individual

            'Estilo de border: continuous (solid), double, dot (dotted), dash (dashed), etc
            C.Cell("border-style: continuous", "G4", "border-style: continuous")
            C.Cell("border-style: continuous double", "G6", "border-style: continuous double")
            C.Cell("border-style: continuous double dot", "G8", "border-style: continuous double dot")
            C.Cell("border-style: continuous double dot dash", "G10", "border-style: continuous double dot dash")

            'Ancho de border: hairline (super delgado), thin (delgado), medium (medio), thick (ancho) :::: al parecer no soporta thick
            C.Cell("hairline, medium, thick, thin", "G12", "border-top-width: hairline")
            C.Cell("hairline, medium, thick, thin", "G12", "border-right-width: medium")
            C.Cell("hairline, medium, thick, thin", "G12", "border-bottom-width: thick")
            C.Cell("hairline, medium, thick, thin", "G12", "border-left-width: thin")

            'border-width: css
            C.Cell("border-width: medium", "G14", "border-width: medium")
            C.Cell("border-width: medium thick", "G16", "border-width: medium thick")
            C.Cell("border-width: medium thick thin", "G18", "border-width: medium thick thin")
            C.Cell("border-width: medium thick thin hairline", "G20", "border-width: medium thick thin hairline")

            'border-color: css
            C.Cell("border-color: yellow", "G22", "border-color: yellow")
            C.Cell("border-color: yellow #FF0000", "G24", "border-color: yellow #FF0000")
            C.Cell("border-color: yellow #FF0000 rgb(103, 49, 71)", "G26", "border-color: yellow #FF0000 rgb(103, 49, 71)")
            C.Cell("border-color: yellow #FF0000 rgb(103, 49, 71) blue", "G28", "border-color: yellow #FF0000 rgb(103, 49, 71) blue")

            'border-top (css)
            'border-right (css)
            'border-bottom (css)
            'border-left (css)
            C.Cell("border-top: rgb(120,120,120) thin dashed; border-right: medium blue dotted; border-bottom: thick #00FF00; border-left: solid", "G30", "border-top: rgb(120,120,120) thin dashed; border-right: medium blue dotted; border-bottom: thick #00FF00; border-left: solid")

            'border: css    
            C.Cell("border: dashed blue medium", "G32", "border: dashed blue medium")
            C.Cell("border: medium dashed rgb(125,125,125)", "G34", "border: medium dashed rgb(125,125,125)")
            C.Cell("border: double #00FF00 thin", "G36", "border: double #00FF00 thin")

            ''-----------COLORES DE BORDERS CON ESTILOS--------------------
            C.Cell("border: continuous; border-color: blue", "G38", "border: continuous; border-color: blue")
            C.Cell("border: continuous; border-color: #F0F0F0", "G40", "border: continuous; border-color: #F0F0F0")
            C.Cell("border: continuous; border-color: rgb(255,100,25)", "G42", "border: continuous; border-color: rgb(255,100,25)")

            C.Cell("border-left: double; border-color:yellow", "G44", "border-left: double; border-color:yellow")
            C.Cell("border-left: double; border-bottom: dashdotdot; border-color: pink", "G46", "border-left: double; border-bottom: dashdotdot; border-color: pink")

            C.Cell("border: double; border-left-color: blue; border-top-color: yellow", "G48", "border: double; border-left-color: blue; border-top-color: yellow")

            C.Cell("border: double; border-color: yellow; border-bottom-color: blue; border-inside-vertical-color: red", "<i>G50:H53", "border: double; border-color: yellow; border-bottom-color: blue; border-inside-vertical-color: red")



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

            '---------STRIKETHROUGH - FUENTE DE LETRA TACHADA---------------
            C.Cell("font-style: strikethrough", "I13", "font-style: strikethrough")
            C.Cell("font-style: line-through", "I14", "font-style: line-through")
            C.Cell("font-style: strikethrough italic", "I15", "font-style: strikethrough italic")
            C.Cell("font-style: line-through bold", "I16", "font-style: line-through bold")
            C.Cell("strikethrough: True", "I17", "strikethrough: True")
            C.Cell("text-decoration-line: strikethrough", "I18", "text-decoration-line: strikethrough")
            C.Cell("text-decoration-line: line-through", "I19", "text-decoration-line: line-through")
            C.Cell("text-decoration-line: strikethrough underline", "I20", "text-decoration-line: strikethrough underline")
            C.Cell("text-decoration-line: line-through underline", "I21", "text-decoration-line: line-through underline")
            C.Cell("text-decoration-line: line-through underline-double", "I22", "text-decoration-line: line-through underline-double")
            C.Cell("text-decoration-style: solid", "I23", "text-decoration-style: solid")

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

            ''---------UPPERCASE & LOWERCASE------
            C.Cell("tExT-tRaNsFoRm: none", "n1", "text-transform: none")
            C.Cell("tExT-tRaNsFoRm: uppercase", "n2", "text-transform: uppercase")
            C.Cell("tExT-tRaNsFoRm: lowercase", "n3", "text-transform: lowercase")
            C.Cell("tExT-tRaNsFoRm: capitalize", "n4", "text-transform: capitalize")

            ''-----------RICH TEXT (HTML EN LA CELDA)----------
            'C.Cell("<h>Hola mundo</h><u>underline</u> HTML en <b>TODOS lados</b>", "L2:M2") --pendiente
            ''-----------CELL VERSIÓN 4 (CREACIÓN DE UNA TABLA)
        End If






        'Siguientes mejoras:
        '3.- HTML dentro de la celda
        '---->4.- Controlar el grosor de los bordes (listo) no logra soportar "dashed thick" guiones supergruesos
        '5.- Orientación de Texto
        '6.- Sangría de Texto
        '---->7.- Reducir Hasta ajustar(texto)  (listo)
        '---->8.- (Fuente) Efecto tachado (listo strikthrough)
        '9.- (Fuente) Efecto Superíndice
        '10.- (Fuente) Efecto Subíndice
        '11.- (Color Fondo) Efectos de relleno, Color de Trama, Estilo de Trama
        '12.- Sección Proteger (Bloqueada, Oculta)
        '---->14.- Ajustar interpretación del key border y sus derivados para apegarse a css (listo border y sus derivados)
        '15.- Estandarizar font-size https://www.freecodecamp.org/espanol/news/tamano-de-fuente-html-como-cambiar-el-tamano-del-texto-usando-el-estilo-css-en-linea/#:~:text=C%C3%B3mo%20cambiar%20el%20tama%C3%B1o%20del%20texto%20usando%20CSS%20en%20l%C3%ADnea,y%20luego%20as%C3%ADgnale%20un%20valor.
        '16.- Uppercase, lowercase


        'Siguientes propuestas
        '1.- (NIVEL CELDA) Dividir texto en varias fracciones de celda (para evitar textos extensos y que no quepan en una celda de excel)
        '2.- (NIVEL TABLA) Hacer un procedimiento para creación de tablas preconfiguradas(desde títulos de columnas hasta la información que pinte desde una fuente de datos)
        '2.1.- (NIVEL TABLA) Hacer que esta tabla pintada sea una tipo tabla que se puedan configurar filtros, formatos de pintado, tablas dinámicas, funciones, etc.
        '3.- (NIVEL CELDA) Dar formato a nivel caracteres(ej. caracteres en diferentes colores y/o fuentes de letras)
        '4.- Programar una macro XD
        '5.- Insertar gráficos
        '6.- Insertar Ilustraciones (Imágenes, Formas, Íconos, Modelos 3D, etc.)
        '7.- Controlar las diposiciones de la página
        '8.- Opciones de Guardado
        '9.- Temas de tablas y celdas


        Book.SaveAs(Filename:="..\source\repos\InteropExcel\Exports\test01.xlsx")  'Cambiar ruta si es necesario

        ExcelApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic
        ExcelApp.ScreenUpdating = True
        ExcelApp.Visible = True
        ExcelApp.DisplayAlerts = True

        ExcelApp.Quit()

        Book = Nothing
        ExcelApp = Nothing
    End Sub

End Module
