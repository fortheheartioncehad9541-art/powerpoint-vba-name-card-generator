Attribute VB_Name = "Módulo1"
Option Explicit

Sub GenerarLetrerosIntegrantes()

    Dim nombres As Variant
    Dim presentacion As Presentation
    Dim diapositivaPlantilla As Slide
    Dim diapositivaNueva As Slide
    Dim rangoDuplicado As SlideRange
    Dim cuadroNombre As Shape
    Dim i As Long

    nombres = Split("Alex Rivera|Sofía Montes|Daniel Lozano|Camila Robles|Mateo Villaseñor|Elena del Bosque|Nicolás Ferrer|Valeria Santillán|Samuel Alcázar|Renata Solís|Leonardo Buenaventura|Marina Castañeda", "|")
    
    Set presentacion = ActivePresentation

    If presentacion.Slides.Count = 0 Then
        MsgBox "La presentación no contiene diapositivas.", _
               vbExclamation, "No se puede continuar"
        Exit Sub
    End If

    If presentacion.Slides.Count > 1 Then
        MsgBox "La presentación debe contener únicamente la diapositiva plantilla." & vbCrLf & vbCrLf & _
               "Elimina las diapositivas adicionales y vuelve a ejecutar la macro.", _
               vbExclamation, "Hay diapositivas adicionales"
        Exit Sub
    End If

    Set diapositivaPlantilla = presentacion.Slides(1)

    On Error Resume Next
    Set cuadroNombre = diapositivaPlantilla.Shapes("NOMBRE_INTEGRANTE")
    On Error GoTo 0

    If cuadroNombre Is Nothing Then
        MsgBox "No se encontró el objeto NOMBRE_INTEGRANTE en la diapositiva 1.", _
               vbCritical, "Objeto no encontrado"
        Exit Sub
    End If

    Application.StartNewUndoEntry

    ' Colocar el primer nombre en la diapositiva original
    ConfigurarNombre diapositivaPlantilla.Shapes("NOMBRE_INTEGRANTE"), _
                     CStr(nombres(0))

    ' Crear las diapositivas restantes
    For i = 1 To UBound(nombres)

        Set rangoDuplicado = diapositivaPlantilla.Duplicate
        Set diapositivaNueva = rangoDuplicado(1)

        diapositivaNueva.MoveTo presentacion.Slides.Count

        ConfigurarNombre diapositivaNueva.Shapes("NOMBRE_INTEGRANTE"), _
                         CStr(nombres(i))

    Next i

    MsgBox "Proceso terminado correctamente." & vbCrLf & vbCrLf & _
           "Se generaron " & CStr(UBound(nombres) + 1) & " letreros.", _
           vbInformation, "Letreros generados"

End Sub


Private Sub ConfigurarNombre( _
    ByVal cuadroNombre As Shape, _
    ByVal nombreIntegrante As String)

    With cuadroNombre.TextFrame2

        .TextRange.Text = nombreIntegrante

        ' Mantener cada nombre en una sola línea
        .WordWrap = msoFalse

        ' Reducir el tamaño cuando el nombre no quepa
        .AutoSize = msoAutoSizeTextToFitShape

        ' Centrar verticalmente
        .VerticalAnchor = msoAnchorMiddle

        ' Márgenes internos
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0

        ' Mantener alineación a la izquierda
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft

    End With

End Sub

