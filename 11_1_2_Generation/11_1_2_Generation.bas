Attribute VB_Name = "Module1"
Dim khm(1 To 2, 1 To 24) As Integer                                         ' Почасовки 2-ух генераторов ХАЕС
Dim uzhuk(1 To 3, 1 To 24) As Integer                                       ' Почасовки 3-ех генераторов ЮУАЕС
Dim zap(1 To 6, 1 To 24) As Integer                                         ' Почасовки 6-ти генераторов ЗАЕС
Dim rivn(1 To 6, 1 To 24) As Integer                                        ' Почасовки 6-ти генераторов РАЕС
Dim tash(1 To 24) As Integer                                                ' Почасовки шини ТГАЕС (2-а генератора объединены)

Dim naek(1 To 24) As Integer                                                ' Почасовки атомной генерации (без Ташлыка и Александровки)

' --- Створення файлу XML для завантаженя на TRANSPARENCY PLATFORM ENTSO-E
Sub First()
    Dim shDataAll, shDataFirst, shDataSecond As Worksheet                   ' Листы Эксель
    Dim doc As MSXML2.DOMDocument60                                         ' Документ
    Dim root As MSXML2.IXMLDOMElement, dataNode As MSXML2.IXMLDOMElement    ' Корневой элемент, узел
    Dim strEnergoatom, strmRID, xmlFileName As String                       ' АйДи документа, имя XML-файла
    Dim dDt As Object                                                       ' Объект для времени UTC
    
    Set shDataAll = Worksheets("data")                                      ' Первый лист Эксель в переменную
    Set shDataFirst = Worksheets("1")                                       ' Второй лист Эксель в переменную
    Set shDataSecond = Worksheets("2")                                      ' Третий лист Эксель в переменную
    Set doc = New MSXML2.DOMDocument60                                      ' Создание документа
    Set root = doc.createElement("GL_MarketDocument")                       ' Создание корневого элемента
    doc.appendChild root                                                    ' Добавление в документ корневого элемента
    AddAttributeWithValue root, "xmlns", "urn:iec62325.351:tc57wg16:451-6:generationloaddocument:3:0"   ' Добавление атрибута namespace по умолчанию для корневого элемента
    
    strEnergoatom = "62X205270350215R"                                      ' EIC-код NNEGC
    strmRID = strEnergoatom & "-EA-" & Format(Now, "yyyy-mm-dd")            ' Сборка уникального АйДи документа
  
    ''' ----------------------------------------------- Set header ----------------------------------------------------------------------------------------------------- '''
    
    '''mRID
        Set mRIDNode = doc.createElement("mRID")                            ' Создание элемента "АйДи" -> 1
        mRIDValue = strmRID                                                 ' АйДи документа в переменную
        Set tagText = doc.createTextNode(mRIDValue)                         ' Создание текстового узла со значением АйДи документа
        mRIDNode.appendChild (tagText)                                      ' Добавление текстового узла в новосозданный элемент "АйДи"
        root.appendChild mRIDNode                                           ' Добавление элемента "АйДи" в корневой элемент
    
    '''revisionNumber
        Set dataNode = doc.createElement("revisionNumber")                  ' Элемент "НомерРевизии" -> 2
        revisionNumberValue = shDataAll.Range("B6").Text ' "1"
        Set tagText = doc.createTextNode(revisionNumberValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode
    
    ''type
        Set dataNode = doc.createElement("type")                            ' Элемент "Тип" -> 3
        typeValue = "A73"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''process.processType
        Set dataNode = doc.createElement("process.processType")             ' Элемент "ТипПроцесса" -> 4
        typeValue = "A16"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''sender_MarketParticipant.mRID
        Set dataNode = doc.createElement("sender_MarketParticipant.mRID")   ' Элемент "Отправитель" -> 5
        typeValue = "62X205270350215R"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode
        AddAttributeWithValue dataNode, "codingScheme", "A01"
        
    ''sender_MarketParticipant.marketRole.type
        Set dataNode = doc.createElement("sender_MarketParticipant.marketRole.type")
        typeValue = "A39"   ' --- изменено B1
        Set tagText = doc.createTextNode(typeValue)                         ' Элемент "ОтправительРоль" -> 6
        dataNode.appendChild (tagText)
        root.appendChild dataNode
         
    ''receiver_MarketParticipant.mRID
        Set dataNode = doc.createElement("receiver_MarketParticipant.mRID") ' Элемент "Получатель" -> 7
        typeValue = "10X1001C--00001X"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode
        AddAttributeWithValue dataNode, "codingScheme", "A01"
    
    ''receiver_MarketParticipant.marketRole.type
        Set dataNode = doc.createElement("receiver_MarketParticipant.marketRole.type")
        typeValue = "A32"
        Set tagText = doc.createTextNode(typeValue)                         ' Элемент "ОтправительРоль" -> 8
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''createdDateTime
        Set dataNode = doc.createElement("createdDateTime")                 ' Элемент "ДатаВремяСоздания" -> 9
        Set dDt = CreateObject("WbemScripting.SWbemDateTime")
            dDt.SetVarDate Now
            Now_ = dDt.GetVarDate(False)
        createdTimeValue = Now_
        createdTimeValueFormated = Format(Now_, "yyyy-mm-ddThh:nn:ssZ")
        Set tagText = doc.createTextNode(createdTimeValueFormated)
        dataNode.appendChild (tagText)
        root.appendChild dataNode
    
    ''time_Period.timeInterval
        Set dataNode = doc.createElement("time_Period.timeInterval")        ' Элемент "ВременнойИнтервал"
        dCurrDate = CDate(shDataAll.Range("B5").Text)
        dDt.SetVarDate (dCurrDate)
        dStartTimeLine = dDt.GetVarDate(False)
        dDt.SetVarDate DateAdd("d", 1, dCurrDate)
        dEndTimeLine = dDt.GetVarDate(False)
        
        StartValue = Format(dStartTimeLine, "yyyy-mm-ddThh:mmZ")            ' Элемент "Старт" -> 10
            Set subNode = doc.createElement("start")
            Set tagText = doc.createTextNode(StartValue)
            subNode.appendChild (tagText)
            dataNode.appendChild (subNode)
        
        endValue = Format(dEndTimeLine, "yyyy-mm-ddThh:mmZ")                ' Элемент "Енд" -> 11
            Set subNode = doc.createElement("end")
            Set tagText = doc.createTextNode(endValue)
            subNode.appendChild (tagText)
            dataNode.appendChild (subNode)
            
        root.appendChild dataNode

    '' -------------------------------------------------------------------- set TimeSeries - loop ------------------------------------------------------------------------------------
    
    For i = 1 To 18
        Set dataNodeTimeSeries = doc.createElement("TimeSeries")            ' Элемент "ТаймСерия"

    '' set mRID
        Set DataNodeSeries = doc.createElement("mRID")                      ' Элемент "АйДи" -> 1
        mRIDValue = "1"

        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set businessType
        Set DataNodeSeries = doc.createElement("businessType")              ' Элемент "БизнесТайп" -> 2
        mRIDValue = "A01"
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set objectAggregation
        Set DataNodeSeries = doc.createElement("objectAggregation")         ' Элемент "ОбъектАгрегации" -> 3
        mRIDValue = "A06"
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set area_Domain.mRID
        Set NodeSeries = doc.createElement("inBiddingZone_Domain.mRID")     ' Элемент "БиддингЗона" -> 4
        mRIDValue = "10Y1001C--000182"
        Set tagText = doc.createTextNode(mRIDValue)
        NodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (NodeSeries)
       
        Set att = doc.createAttribute("codingScheme")
        att.Value = "A01"
        NodeSeries.setAttributeNode att

    ''set measure_Unit.name
        Set DataNodeSeries = doc.createElement("quantity_Measure_Unit.name")
        mRIDValue = "MAW"                                                   ' Элемент "ЕдиницаИзмерения" -> 5
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)
         
    ''set curveType
        Set DataNodeSeries = doc.createElement("curveType")                 ' Элемент "ТипКривой" -> 6
        txtCurveType = "A01"
        Set tagText = doc.createTextNode(txtCurveType)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set MktPSRType
        Set DataNodeSeries = doc.createElement("MktPSRType")
        Set psrType = doc.createElement("psrType")                          ' Элемент "ТипГенерации" -> 7
        Set tagText = doc.createTextNode("B14")                             ' Nuclear
        
        If i = 18 Then
            Set tagText = doc.createTextNode("B10")                         ' Hydro Pump Storage
        End If
        
        psrType.appendChild (tagText)
        DataNodeSeries.appendChild (psrType)
        dataNodeTimeSeries.appendChild DataNodeSeries

        Set subNode = doc.createElement("PowerSystemResources")             ' Элемент "РесурсыЭнергетическойСистемы"
        DataNodeSeries.appendChild (subNode)
        dataNodeTimeSeries.appendChild DataNodeSeries
         
        Set mRID = doc.createElement("mRID")                                ' Элемент "АйДи" -> 8
        mRIDValue = shDataFirst.Cells(1 + i, 3)
        Set tagText = doc.createTextNode(mRIDValue)
        mRID.appendChild (tagText)
        subNode.appendChild (mRID)
            Set att = doc.createAttribute("codingScheme")
            att.Value = "A01"
        mRID.setAttributeNode att
        
    '' --------------------------------------------------- set Period -----------------------------------------------------------------------------------------------------------
        
        Set DataNodeSeries = doc.createElement("Period")                    ' Элемент "Период"
    '' timeInterval
         Set dataNode = doc.createElement("timeInterval")                   ' Элемент "ТаймИнтервал"
         StartValue = Format(dStartTimeLine, "yyyy-mm-ddThh:mmZ")  '+++
         
         Set subNode = doc.createElement("start")                           ' Элемент "Старт" -> 1
         Set tagText = doc.createTextNode(StartValue)
         subNode.appendChild (tagText)
         dataNode.appendChild (subNode)
         DataNodeSeries.appendChild dataNode
         
         endValue = Format(dEndTimeLine, "yyyy-mm-ddThh:mmZ")   '+++
         
         Set subNode = doc.createElement("end")                             ' Элемент "Енд" -> 2
         Set tagText = doc.createTextNode(endValue)
         subNode.appendChild (tagText)
         dataNode.appendChild (subNode)
         DataNodeSeries.appendChild dataNode
         
         dataNodeTimeSeries.appendChild DataNodeSeries
    
    '' set  time interval
         Set Resolution = doc.createElement("resolution")                   ' Элемент "Разрешение" -> 3
         Set tagText = doc.createTextNode("PT60M")
         Resolution.appendChild (tagText)
         DataNodeSeries.appendChild Resolution

         For J = 1 To 24
            Set Point = doc.createElement("Point")                          ' Элемент "Точка"
            DataNodeSeries.appendChild Point
              
            Set Position = doc.createElement("position")                    ' Элемент "Позиция" -> 4
            Set tagText = doc.createTextNode(J)
            Position.appendChild (tagText)
            Point.appendChild (Position)
              
            Set Quantity = doc.createElement("quantity")                    ' Элемент "Количество" -> 5
            
            Select Case i
            Case 1 To 2
                shDataFirst.Cells(1 + i, 3 + J).Value = khm(i, J)
                Set tagText = doc.createTextNode(khm(i, J))
            Case 3 To 5
                shDataFirst.Cells(1 + i, 3 + J).Value = uzhuk(i - 2, J)
                Set tagText = doc.createTextNode(uzhuk(i - 2, J))
            Case 6 To 11
                shDataFirst.Cells(1 + i, 3 + J).Value = zap(i - 5, J)
                Set tagText = doc.createTextNode(zap(i - 5, J))
            Case 12 To 17
                shDataFirst.Cells(1 + i, 3 + J).Value = rivn(i - 11, J)
                Set tagText = doc.createTextNode(rivn(i - 11, J))
            
            Case 18
                shDataFirst.Cells(1 + i, 3 + J).Value = tash(J)
                Set tagText = doc.createTextNode(tash(J))
            End Select
            
            Quantity.appendChild (tagText)
            Point.appendChild (Quantity)
         Next J

        root.appendChild dataNodeTimeSeries
    Next i
    
        Set pi = doc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")           ' Создание инструкций обработки
        xmlFileName1 = Left(ThisWorkbook.Path, 69) & "\_NAEK\Rynok_EE\Rynok_XML\18_11.1_NNEGC.xml"      ' Имя файла XML
       
        Smooth_Xml doc                                                                                  ' очистка файла от определенных символов
        doc.InsertBefore pi, doc.ChildNodes.Item(0)                                                     ' вставка инструкций обработки в нулевой элемент документа
        doc.Save xmlFileName1                                                                           ' сохранение созданного файла XML
        
        MsgBox ("Файл XML створений та записаний на диск ...")                                          ' сообщение о успешном создании файла XML
End Sub

Sub AddAttributeWithValue(ByRef el As IXMLDOMElement, attName, attValue)
    Dim att
    Set att = el.OwnerDocument.createAttribute(attName)
    att.Value = attValue
    el.setAttributeNode att
End Sub

Sub Smooth_Xml(inDoc)
 inDoc.LoadXML Replace(inDoc.XML, "><", ">" & vbCrLf & "<")
 inDoc.LoadXML Replace(inDoc.XML, "/>", "/>" & vbCrLf)
 inDoc.LoadXML Replace(inDoc.XML, "xmlns=""""", "")
End Sub

' --- Завантаження даних по ОДИНИЦЯМ ГЕНЕРАЦІЇ з файлу РЕЄСТРА
Sub Upload_Reestr()
    Dim NameFile As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Integer
    Dim shDataAll As Worksheet
    
''strDocDate
    Set shDataAll = Worksheets("data")
    strDocDate = Format(DateAdd("d", -1, Now), "dd.mm.yyyy")
    shDataAll.Range("B5").Value = strDocDate
    
    NameFile = GetFilePath()
    Application.ScreenUpdating = False
    If (NameFile = "") Then Exit Sub
    
    Set wb = Excel.Workbooks.Open(Filename:=NameFile)
    Set ws = wb.Worksheets("Реестр")
    
    With ws
        For i = 1 To 24
           ' Debug.Print
           khm(1, i) = WorksheetFunction.Round(.Cells(67, 10 + i).Value / 1000, 0)
           khm(2, i) = WorksheetFunction.Round(.Cells(68, 10 + i).Value / 1000, 0)
           
           uzhuk(1, i) = WorksheetFunction.Round(.Cells(48, 10 + i).Value / 1000, 0)
           uzhuk(2, i) = WorksheetFunction.Round(.Cells(49, 10 + i).Value / 1000, 0)
           uzhuk(3, i) = WorksheetFunction.Round(.Cells(50, 10 + i).Value / 1000, 0)
           
           zap(1, i) = WorksheetFunction.Round(.Cells(39, 10 + i).Value / 1000, 0)
           zap(2, i) = WorksheetFunction.Round(.Cells(40, 10 + i).Value / 1000, 0)
           zap(3, i) = WorksheetFunction.Round(.Cells(41, 10 + i).Value / 1000, 0)
           zap(4, i) = WorksheetFunction.Round(.Cells(42, 10 + i).Value / 1000, 0)
           zap(5, i) = WorksheetFunction.Round(.Cells(43, 10 + i).Value / 1000, 0)
           zap(6, i) = WorksheetFunction.Round(.Cells(44, 10 + i).Value / 1000, 0)
           
           rivn(1, i) = WorksheetFunction.Round(.Cells(56, 10 + i).Value / 1000, 0)
           rivn(2, i) = WorksheetFunction.Round(.Cells(57, 10 + i).Value / 1000, 0)
           rivn(3, i) = WorksheetFunction.Round(.Cells(59, 10 + i).Value / 1000, 0)
           rivn(4, i) = WorksheetFunction.Round(.Cells(60, 10 + i).Value / 1000, 0)
           rivn(5, i) = WorksheetFunction.Round(.Cells(62, 10 + i).Value / 1000, 0)
           rivn(6, i) = WorksheetFunction.Round(.Cells(63, 10 + i).Value / 1000, 0)
           
           tash(i) = WorksheetFunction.Round(.Cells(52, 10 + i).Value / 1000, 0)
           
           naek(i) = WorksheetFunction.Round((.Cells(35, 10 + i).Value - .Cells(51, 10 + i).Value - .Cells(52, 10 + i).Value) / 1000, 0)
        Next i
    End With
    
    ActiveWorkbook.Close False
    Application.ScreenUpdating = True
    
    MsgBox "Дані зчитані з РЕЄСТРУ ..."
End Sub

' --- Вибір файлу РЕЄСТРУ
Function GetFilePath() As String
       
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False 'флаг вибіра одного, або багато файлів
        .Title = "Віберіть файл" 'Вібір файла для перевірки
        .Filters.Clear 'чистимо встановленні раніше типи файлів
        .Filters.Add "Excel files", "*.xls*;*.xla*" 'втановлюєм типи файлів Excel
        
        .InitialFileName = "J:\ASKUE_(otchet)\Reestr_ORE"
        
        .InitialView = msoFileDialogViewDetails 'вид диалогового вікна
        If .Show <> -1 Then Exit Function
        GetFilePath = .SelectedItems(1)
    End With
End Function


' --- Створення файлу XML для завантаженя на TRANSPARENCY PLATFORM ENTSO-E
Sub Second()
    Dim shDataAll, shDataSecond As Worksheet                                ' Листы Эксель
    Dim doc As MSXML2.DOMDocument60                                         ' Документ
    Dim root As MSXML2.IXMLDOMElement, dataNode As MSXML2.IXMLDOMElement    ' Корневой элемент, узел
    Dim strmRID, strEnergoatom As String                                    ' АйДи документа, имя XML-файла
    Dim dDt As Object                                                       ' Объект для времени UTC

    Set shDataAll = Worksheets("data")                                      ' Первый лист Эксель
    Set shDataDay = Worksheets("2")                                         ' Третий лист Эксель
    Set doc = New MSXML2.DOMDocument60                                      ' Создание документа
    Set root = doc.createElement("GL_MarketDocument")                       ' Создание корневого элемента
    doc.appendChild root                                                    ' Добавление в документ корневого элемента
    AddAttributeWithValue root, "xmlns", "urn:iec62325.351:tc57wg16:451-6:generationloaddocument:3:0"   ' Добавление атрибута namespace по умолчанию для корневого элемента
    
    strEnergoatom = "62X205270350215R"                                      ' EIC-код NNEGC
    strmRID = strEnergoatom & "-EA-" & Format(Now, "yyyy-mm-dd")            ' Сборка уникального АйДи документа
    
    ''' --------------------------------------------------------------- Set header ---------------------------------------------------------------------------------- '''
    
    ''mRID
        Set mRIDNode = doc.createElement("mRID")                            ' Элемент "АйДи" -> 1
        mRIDValue = strmRID ' "1"
        Set tagText = doc.createTextNode(mRIDValue)
        mRIDNode.appendChild (tagText)
        root.appendChild mRIDNode
    
    ''revisionNumber
        Set dataNode = doc.createElement("revisionNumber")                  ' Элемент "НомерРевизии" -> 2
        revisionNumberValue = shDataAll.Range("B6") ' "1"
        Set tagText = doc.createTextNode(revisionNumberValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode
    
    ''type
        Set dataNode = doc.createElement("type")                            ' Элемент "Тип" -> 3
        typeValue = "A75"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''process.processType
        Set dataNode = doc.createElement("process.processType")             ' Элемент "ТипПроцесса" -> 4
        typeValue = "A16"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''sender_MarketParticipant.mRID
        Set dataNode = doc.createElement("sender_MarketParticipant.mRID")   ' Элемент "Отправитель" -> 5
        typeValue = "62X205270350215R"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode
        AddAttributeWithValue dataNode, "codingScheme", "A01"
        
    ''sender_MarketParticipant.marketRole.type
        Set dataNode = doc.createElement("sender_MarketParticipant.marketRole.type")
        typeValue = "A39"                                                   ' Элемент "ОтправительРоль" -> 6
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode
         
    ''receiver_MarketParticipant.mRID
        Set dataNode = doc.createElement("receiver_MarketParticipant.mRID") ' Элемент "Получатель" -> 7
        typeValue = "10X1001C--00001X"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode
        AddAttributeWithValue dataNode, "codingScheme", "A01"
    
    ''receiver_MarketParticipant.marketRole.type
        Set dataNode = doc.createElement("receiver_MarketParticipant.marketRole.type")
        typeValue = "A32"                                                   ' Элемент "ОтправительРоль" -> 8
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''createdDateTime
        Set dataNode = doc.createElement("createdDateTime")                 ' Элемент "ДатаВремяСоздания" -> 9
        Set dDt = CreateObject("WbemScripting.SWbemDateTime")
            dDt.SetVarDate Now
            Now_ = dDt.GetVarDate(False)
        createdTimeValue = Now_
        createdTimeValueFormated = Format(Now_, "yyyy-mm-ddThh:nn:ssZ")
        Set tagText = doc.createTextNode(createdTimeValueFormated)
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''time_Period.timeInterval
        Set dataNode = doc.createElement("time_Period.timeInterval")        ' Элемент "ВременнойИнтервал"
        dCurrDate = CDate(shDataAll.Range("B5").Text)
        
        dDt.SetVarDate (dCurrDate)
        dStartTimeLine = dDt.GetVarDate(False)
                
        dDt.SetVarDate DateAdd("d", 1, dCurrDate)
        dEndTimeLine = dDt.GetVarDate(False)
        
        StartValue = Format(dStartTimeLine, "yyyy-mm-ddThh:mmZ")            ' Элемент "Старт" -> 10
            Set subNode = doc.createElement("start")
            Set tagText = doc.createTextNode(StartValue)
            subNode.appendChild (tagText)
            dataNode.appendChild (subNode)
        
        endValue = Format(dEndTimeLine, "yyyy-mm-ddThh:mmZ")                ' Элемент "Енд" -> 11
            Set subNode = doc.createElement("end")
            Set tagText = doc.createTextNode(endValue)
            subNode.appendChild (tagText)
            dataNode.appendChild (subNode)
        
        root.appendChild dataNode

'' ------------------------------------------------------------------------ set TimeSeries - loop -------------------------------------------------------------------------------
   ' For i = 1 To 18
        Set dataNodeTimeSeries = doc.createElement("TimeSeries")            ' Элемент "ТаймСерия"

    '' set mRID
        Set DataNodeSeries = doc.createElement("mRID")                      ' Элемент "АйДи" -> 1
        mRIDValue = "1"
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set businessType
        Set DataNodeSeries = doc.createElement("businessType")              ' Элемент "БизнесТайп" -> 2
        mRIDValue = "A01"
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set objectAggregation
        Set DataNodeSeries = doc.createElement("objectAggregation")         ' Элемент "ОбъектАгрегации" -> 3
        mRIDValue = "A08"
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set area_Domain.mRID
        Set NodeSeries = doc.createElement("inBiddingZone_Domain.mRID")     ' Элемент "БиддингЗона" -> 4
        mRIDValue = "10Y1001C--000182"
        Set tagText = doc.createTextNode(mRIDValue)
        NodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (NodeSeries)
            Set att = doc.createAttribute("codingScheme")
            att.Value = "A01"
            NodeSeries.setAttributeNode att

    ''set measure_Unit.name
        Set DataNodeSeries = doc.createElement("quantity_Measure_Unit.name")
        mRIDValue = "MAW"                                                   ' Элемент "ЕдиницаИзмерения" -> 5
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)
         
    ''set curveType
        Set DataNodeSeries = doc.createElement("curveType")                 ' Элемент "ТипКривой" -> 6
        txtCurveType = "A01"
        Set tagText = doc.createTextNode(txtCurveType)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set MktPSRType
        Set DataNodeSeries = doc.createElement("MktPSRType")                ' Элемент "ТипГенерации" -> 7
        Set psrType = doc.createElement("psrType")
        Set tagText = doc.createTextNode("B14")                             ' Nuclear
        psrType.appendChild (tagText)
        DataNodeSeries.appendChild (psrType)
        dataNodeTimeSeries.appendChild DataNodeSeries

    '' ---------------------------------------------------------------------- Period ------------------------------------------------------------------------------------------
        Set DataNodeSeries = doc.createElement("Period")                    ' Элемент "Период"
    '' timeInterval
        Set dataNode = doc.createElement("timeInterval")                    ' Элемент "ТаймИнтервал"
        StartValue = Format(dStartTimeLine, "yyyy-mm-ddThh:mmZ")
            Set subNode = doc.createElement("start")                        ' Элемент "Старт" -> 1
            Set tagText = doc.createTextNode(StartValue)
            subNode.appendChild (tagText)
            dataNode.appendChild (subNode)
            DataNodeSeries.appendChild dataNode
         
        endValue = Format(dEndTimeLine, "yyyy-mm-ddThh:mmZ")   '+++
            Set subNode = doc.createElement("end")                          ' Элемент "Енд" -> 2
            Set tagText = doc.createTextNode(endValue)
            subNode.appendChild (tagText)
            dataNode.appendChild (subNode)
            DataNodeSeries.appendChild dataNode
         
         dataNodeTimeSeries.appendChild DataNodeSeries
    
    '' set  time interval
         Set Resolution = doc.createElement("resolution")                   ' Элемент "Разрешение" -> 3
         Set tagText = doc.createTextNode("PT60M")
         Resolution.appendChild (tagText)
         DataNodeSeries.appendChild Resolution

         For J = 1 To 24
            Set Point = doc.createElement("Point")                          ' Элемент "Точка"
            DataNodeSeries.appendChild Point
              
            Set Position = doc.createElement("position")                    ' Элемент "Позиция" -> 4
            Set tagText = doc.createTextNode(J)
            Position.appendChild (tagText)
            Point.appendChild (Position)
              
            Set Quantity = doc.createElement("quantity")                    ' Элемент "Количество" -> 5
            
            shDataDay.Cells(2, 3 + J).Value = naek(J)
            Set tagText = doc.createTextNode(naek(J))
            
            Quantity.appendChild (tagText)
            Point.appendChild (Quantity)
         Next J

        root.appendChild dataNodeTimeSeries
   ' Next i
    
        Set pi = doc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")           ' Создание инструкций обработки
        xmlFileName1 = Left(ThisWorkbook.Path, 69) & "\_NAEK\Rynok_EE\Rynok_XML\18_11.2_NNEGC.xml"      ' Имя файла XML
        
        Smooth_Xml doc                                                                                  ' очистка файла от определенных символов
        doc.InsertBefore pi, doc.ChildNodes.Item(0)                                                     ' вставка инструкций обработки в нулевой элемент документа
        doc.Save xmlFileName1                                                                           ' сохранение созданного файла XML
        
        MsgBox ("Файл XML створений та записаний на диск ...")                                          ' сообщение о успешном создании файла XML
End Sub

