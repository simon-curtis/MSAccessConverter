Module Constants

    Public Enum AcRecord
        acFirst = 2
        acGoTo = 4
        acLast = 3
        acNewRec = 5
        acNext = 1
        acPrevious = 0
    End Enum

    Public ReadOnly Property acDesign As Byte = 1
    Public ReadOnly Property acFormDS As Byte = 3
    Public ReadOnly Property acFormPivotChart As Byte = 5
    Public ReadOnly Property acFormPivotTable As Byte = 4
    Public ReadOnly Property acLayout As Byte = 6
    Public ReadOnly Property acNormal As Byte = 0
    Public ReadOnly Property acPreview As Byte = 2

    Public ReadOnly Property acFormAdd As Integer = 0
    Public ReadOnly Property acFormEdit As Integer = 1
    Public ReadOnly Property acFormPropertySettings As Integer = -1
    Public ReadOnly Property acFormReadOnly As Integer = 2

    Public ReadOnly Property acDialog As Byte = 3
    Public ReadOnly Property acHidden As Byte = 1
    Public ReadOnly Property acIcon As Byte = 2
    Public ReadOnly Property acWindowNormal As Byte = 0

    Public ReadOnly Property acOutputForm As Byte = 2
    Public ReadOnly Property acOutputFunction As Byte = 0
    Public ReadOnly Property acOutputModule As Byte = 5
    Public ReadOnly Property acOutputQuery As Byte = 1
    Public ReadOnly Property acOutputReport As Byte = 3
    Public ReadOnly Property acOutputServerView As Byte = 7
    Public ReadOnly Property acOutputStoredProcedure As Byte = 9
    Public ReadOnly Property acOutputTable As Byte = 0


    Public ReadOnly Property acSendForm As Integer = 2
    Public ReadOnly Property acSendModule As Integer = 5
    Public ReadOnly Property acSendNoObject As Integer = -1
    Public ReadOnly Property acSendQuery As Integer = 1
    Public ReadOnly Property acSendReport As Integer = 3
    Public ReadOnly Property acSendTable As Integer = 0

    Public ReadOnly Property acFormatASP As String = "acFormatASP"
    Public ReadOnly Property acFormatDAP As String = "acFormatDAP"
    Public ReadOnly Property acFormatHTML As String = "acFormatHTML"
    Public ReadOnly Property acFormatIIS As String = "acFormatIIS"
    Public ReadOnly Property acFormatRTF As String = "acFormatRTF"
    Public ReadOnly Property acFormatSNP As String = "acFormatSNP"
    Public ReadOnly Property acFormatTXT As String = "acFormatTXT"
    Public ReadOnly Property acFormatXLS As String = "acFormatXLS"
    Public ReadOnly Property acFormatXLSX As String = "acFormatXLSX"

    Public ReadOnly Property acExportQualityPrint As Byte = 0
    Public ReadOnly Property acExportQualityScreen As Byte = 1


End Module
