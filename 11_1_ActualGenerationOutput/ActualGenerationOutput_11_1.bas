Attribute VB_Name = "Module1"
Dim khm(1 To 2, 1 To 24) As Integer
Dim uzhuk(1 To 3, 1 To 24) As Integer
Dim zap(1 To 6, 1 To 24) As Integer
Dim rivn(1 To 6, 1 To 24) As Integer
Dim tash(1 To 24) As Integer

' --- Створення файлу XML для завантаженя на TRANSPARENCY PLATFORM ENTSO-E
Sub Main()
    Dim doc As MSXML2.DOMDocument60
    Dim root As MSXML2.IXMLDOMElement, dataNode As MSXML2.IXMLDOMElement
    Dim i As Long
    Dim dDt As Object
    Dim shDataAll, shDataDay As Worksheet
    Dim inQuantity As Integer

    Dim strmRID, strSShNumber, strGeneratorNumber, strCurrentDate, strFromDate As String
    strGeneratorNumber = "62X205270350215R" ' NNEGC
    strFromDate = "31.12.2019"
    strCurrentDate = Date
    'strmRID = strGeneratorNumber & "-" & DateDiff("y", strFromDate, strCurrentDate)
    strmRID = strGeneratorNumber & "-EA-" & Format(Now, "yyyy-mm-dd") ' DateDiff("y", strFromDate, strCurrentDate)

    Set doc = New MSXML2.DOMDocument60
    Set shDataAll = Worksheets("data")
    Set shDataDay = Worksheets("TimeSeries")

    Set root = doc.createElement("GL_MarketDocument")
    doc.appendChild root
   
    AddAttributeWithValue root, "xmlns", "urn:iec62325.351:tc57wg16:451-6:generationloaddocument:3:0"
    Set att = doc.createAttribute("xmlns")
  
    att.Value = "urn:iec62325.351:tc57wg16:451-6:generationloaddocument:3:0"
    
    '''Set header'''
    
    '''mRID
        Set mRIDNode = doc.createElement("mRID")
        mRIDValue = strmRID ' "1"
        Set tagText = doc.createTextNode(mRIDValue)
        mRIDNode.appendChild (tagText)
        root.appendChild mRIDNode
    
    '''revisionNumber
        Set dataNode = doc.createElement("revisionNumber")
        revisionNumberValue = "1"
        Set tagText = doc.createTextNode(revisionNumberValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode
    
    ''type
        Set dataNode = doc.createElement("type")
        typeValue = "A73"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''process.processType
        Set dataNode = doc.createElement("process.processType")
        typeValue = "A16"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''sender_MarketParticipant.mRID
        Set dataNode = doc.createElement("sender_MarketParticipant.mRID")
        typeValue = "62X205270350215R"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode
        AddAttributeWithValue dataNode, "codingScheme", "A01"
        
    ''sender_MarketParticipant.marketRole.type
        Set dataNode = doc.createElement("sender_MarketParticipant.marketRole.type")
        
        ' typeValue = "A04"
        typeValue = "A39"   ' --- изменено B1
        
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode
         
    ''receiver_MarketParticipant.mRID
        Set dataNode = doc.createElement("receiver_MarketParticipant.mRID")
        typeValue = "10X1001C--00001X"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode
        AddAttributeWithValue dataNode, "codingScheme", "A01"
    
    ''receiver_MarketParticipant.marketRole.type
        Set dataNode = doc.createElement("receiver_MarketParticipant.marketRole.type")
        typeValue = "A32"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''createdDateTime
        Set dataNode = doc.createElement("createdDateTime")
        Set dDt = CreateObject("WbemScripting.SWbemDateTime")

            dDt.SetVarDate Now
            Now_ = dDt.GetVarDate(False)

        ' Now_ = Now
        createdTimeValue = Now_
        createdTimeValueFormated = Format(Now_, "yyyy-mm-ddThh:nn:ssZ")
        Set tagText = doc.createTextNode(createdTimeValueFormated)
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''time_Period.timeInterval
        Set dataNode = doc.createElement("time_Period.timeInterval")        ' --- изменено
        
        dCurrDate = CDate(shDataAll.Range("B5").Text)
        
        dDt.SetVarDate (dCurrDate)
        dStartTimeLine = dDt.GetVarDate(False)
                
        dDt.SetVarDate DateAdd("d", 1, dCurrDate)
        dEndTimeLine = dDt.GetVarDate(False)
        
        
        ' StartValue = Format(shDataAll.Range("B2"), "yyyy-mm-ddThh:nn:ssZ")  '+++ изменено B2
        ' StartValue = Format(shDataAll.Range("B2", "yyyy-mm-ddThh:mmZ")  '+++
        StartValue = Format(dStartTimeLine, "yyyy-mm-ddThh:mmZ")  '+++
        
        Set subNode = doc.createElement("start")
        Set tagText = doc.createTextNode(StartValue)
        subNode.appendChild (tagText)
        dataNode.appendChild (subNode)
        
        ' endValue = Format(shDataAll.Range("B3"), "yyyy-mm-ddThh:nn:ssZ")   '+++   изменено B2
        ' endValue = Format(shDataAll.Range("B3"), "yyyy-mm-ddThh:mmZ")   '+++
        endValue = Format(dEndTimeLine, "yyyy-mm-ddThh:mmZ")   '+++
                        
        Set subNode = doc.createElement("end")
        Set tagText = doc.createTextNode(endValue)
        subNode.appendChild (tagText)
        dataNode.appendChild (subNode)
        root.appendChild dataNode

'' set TimeSeries - loop
    For i = 1 To 18
        Set dataNodeTimeSeries = doc.createElement("TimeSeries")

    '' set mRID
        Set DataNodeSeries = doc.createElement("mRID")
        mRIDValue = "1"

        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set businessType
        Set DataNodeSeries = doc.createElement("businessType")
        mRIDValue = "A01"
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set objectAggregation
        Set DataNodeSeries = doc.createElement("objectAggregation")
        mRIDValue = "A06"
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set area_Domain.mRID
        Set NodeSeries = doc.createElement("inBiddingZone_Domain.mRID")
        mRIDValue = "10Y1001C--000182"
        Set tagText = doc.createTextNode(mRIDValue)
        NodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (NodeSeries)
       
        Set att = doc.createAttribute("codingScheme")
        att.Value = "A01"
        NodeSeries.setAttributeNode att

    ''set measure_Unit.name
        Set DataNodeSeries = doc.createElement("quantity_Measure_Unit.name")
        mRIDValue = "MAW"
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)
         
    ''set curveType
        Set DataNodeSeries = doc.createElement("curveType")
        txtCurveType = "A01"
        Set tagText = doc.createTextNode(txtCurveType)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set MktPSRType
        Set DataNodeSeries = doc.createElement("MktPSRType")
        Set psrType = doc.createElement("psrType")
        Set tagText = doc.createTextNode("B14")
        
        If i = 18 Then
            Set tagText = doc.createTextNode("B10")
        End If
        
        psrType.appendChild (tagText)
        DataNodeSeries.appendChild (psrType)
        dataNodeTimeSeries.appendChild DataNodeSeries

        Set subNode = doc.createElement("PowerSystemResources")
        DataNodeSeries.appendChild (subNode)
        dataNodeTimeSeries.appendChild DataNodeSeries
         
        Set mRID = doc.createElement("mRID")
        mRIDValue = shDataDay.Cells(1 + i, 3)
        Set tagText = doc.createTextNode(mRIDValue)
        mRID.appendChild (tagText)
        subNode.appendChild (mRID)
        ' AddAttributeWithValue mRID, "codingScheme", "A01"       ' --- добавлено
        
        Set att = doc.createAttribute("codingScheme")
        att.Value = "A01"
        
        ' NodeSeries.setAttributeNode att
        ' dataNodeTimeSeries.appendChild DataNodeSeries
        mRID.setAttributeNode att
        ' dataNodeTimeSeries.appendChild DataNodeSeries
                 
        Set DataNodeSeries = doc.createElement("Period")
    '' timeInterval
         Set dataNode = doc.createElement("timeInterval")
         ' StartValue = Format(shDataAll.Range("B2"), "yyyy-mm-ddThh:nn:ssZ")  '+++ изменено B3
         ' StartValue = Format(shDataAll.Range("B2"), "yyyy-mm-ddThh:mmZ")  '+++
         StartValue = Format(dStartTimeLine, "yyyy-mm-ddThh:mmZ")  '+++
         
         Set subNode = doc.createElement("start")
         Set tagText = doc.createTextNode(StartValue)
         subNode.appendChild (tagText)
         dataNode.appendChild (subNode)
         DataNodeSeries.appendChild dataNode
         
         ' endValue = Format(shDataAll.Range("B3"), "yyyy-mm-ddThh:nn:ssZ")   '+++ изменено B3
         ' endValue = Format(shDataAll.Range("B3"), "yyyy-mm-ddThh:mmZ")   '+++
         endValue = Format(dEndTimeLine, "yyyy-mm-ddThh:mmZ")   '+++
         
         Set subNode = doc.createElement("end")
         Set tagText = doc.createTextNode(endValue)
         subNode.appendChild (tagText)
         dataNode.appendChild (subNode)
         DataNodeSeries.appendChild dataNode
         
         dataNodeTimeSeries.appendChild DataNodeSeries
    
    '' set  time interval
         Set Resolution = doc.createElement("resolution")
         Set tagText = doc.createTextNode("PT60M")
         Resolution.appendChild (tagText)

         DataNodeSeries.appendChild Resolution

         For j = 1 To 24
            Set Point = doc.createElement("Point")
            DataNodeSeries.appendChild Point
              
            Set Position = doc.createElement("position")
            Set tagText = doc.createTextNode(j)
            Position.appendChild (tagText)
            Point.appendChild (Position)
              
            Set Quantity = doc.createElement("quantity")
            
            ' Set tagText = doc.createTextNode(shDataDay.Cells(27 + i, 3 + j).Text)
            Select Case i
            Case 1 To 2
                shDataDay.Cells(1 + i, 3 + j).Value = khm(i, j)
                Set tagText = doc.createTextNode(khm(i, j))
            Case 3 To 5
                shDataDay.Cells(1 + i, 3 + j).Value = uzhuk(i - 2, j)
                Set tagText = doc.createTextNode(uzhuk(i - 2, j))
            Case 6 To 11
                shDataDay.Cells(1 + i, 3 + j).Value = zap(i - 5, j)
                Set tagText = doc.createTextNode(zap(i - 5, j))
            Case 12 To 17
                shDataDay.Cells(1 + i, 3 + j).Value = rivn(i - 11, j)
                Set tagText = doc.createTextNode(rivn(i - 11, j))
            
            Case 18
                shDataDay.Cells(1 + i, 3 + j).Value = tash(j)
                Set tagText = doc.createTextNode(tash(j))
            End Select
            
            Quantity.appendChild (tagText)
            Point.appendChild (Quantity)
         Next j

        root.appendChild dataNodeTimeSeries
    Next i
    
        sCurrDate = Format(dCurrDate, "yyyy-mm-dd_")
        
        Set pi = doc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
        createdTimeValueFormated = Format(Now, "yyyy-mm-ddThh_nn_ssZ")
        xmlFileName = ThisWorkbook.Path & "\18_11.1_NNEGC.xml"
        Smooth_Xml doc
        doc.InsertBefore pi, doc.ChildNodes.Item(0)
        doc.Save xmlFileName
        
        MsgBox ("Файл XML створений та записаний на диск ...")
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
        .InitialFileName = "C:\Users\fposlushnyi\downloads\_ПРОКОПЕНКО" 'Старовий каталог
        .InitialView = msoFileDialogViewDetails 'вид диалогового вікна
        If .Show <> -1 Then Exit Function
        GetFilePath = .SelectedItems(1)
    End With
End Function

