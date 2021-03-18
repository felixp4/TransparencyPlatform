Attribute VB_Name = "Module1"
Dim naek(1 To 24) As Integer

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
    strmRID = strGeneratorNumber & "-EA-" & Format(Now, "yyyy-mm-dd") ' DateDiff("y", strFromDate, strCurrentDate)

    Set doc = New MSXML2.DOMDocument60
    Set shDataAll = Worksheets("data")
    ' Set shDataDay = Worksheets("TimeSeries")

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
        typeValue = "A71"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''process.processType
        Set dataNode = doc.createElement("process.processType")
        typeValue = "A01"
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
'        sStartDate = Replace(shDataAll.Range("B5").Value, " ", "")
'        If IsDate(sStartDate) Then
'            'strDate = Format(CDate(strDate), "dd.mm.yyyy")
'            dStartDate = Format(CDate(sStartDate), "dd.mm.yyyy hh:mm")
'             '  .Range("B3").Value = CStr(dStartDate)
'        Else
'            MsgBox "Неправильний формат дати в комірці B3."
'           ' GoTo Ends
'        End If
    
        ' dCurrDate = shDataAll.Range("B5")
        ' dDt.SetVarDate (CDate("01.01.2020 12:30:00"))
        dCurrDate = CDate(shDataAll.Range("B5").Text)
        
        dDt.SetVarDate (dCurrDate)
        dStartTimeLine = dDt.GetVarDate(False)
                
        dDt.SetVarDate DateAdd("d", 1, dCurrDate)
        dEndTimeLine = dDt.GetVarDate(False)
    
        Set dataNode = doc.createElement("time_Period.timeInterval")        ' --- изменено
        
        ' StartValue = Format(shDataAll.Range("B2"), "yyyy-mm-ddThh:nn:ssZ")  '+++ изменено B2
        ' StartValue = Format(shDataAll.Range("B2"), "yyyy-mm-ddThh:mmZ")  '+++
        StartValue = Format(dStartTimeLine, "yyyy-mm-ddThh:mmZ")
        
        Set subNode = doc.createElement("start")
        Set tagText = doc.createTextNode(StartValue)
        subNode.appendChild (tagText)
        dataNode.appendChild (subNode)
        
        ' endValue = Format(shDataAll.Range("B3"), "yyyy-mm-ddThh:nn:ssZ")   '+++   изменено B2
        ' endValue = Format(shDataAll.Range("B3"), "yyyy-mm-ddThh:mmZ")   '+++
        endValue = Format(dEndTimeLine, "yyyy-mm-ddThh:mmZ")
                        
        Set subNode = doc.createElement("end")
        Set tagText = doc.createTextNode(endValue)
        subNode.appendChild (tagText)
        dataNode.appendChild (subNode)
        root.appendChild dataNode

'' set TimeSeries - loop
   ' For i = 1 To 18
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
        mRIDValue = "A01"
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
              
    '' set Period
        Set DataNodeSeries = doc.createElement("Period")
        
    '' timeInterval
         Set dataNode = doc.createElement("timeInterval")
         ' StartValue = Format(shDataAll.Range("B2"), "yyyy-mm-ddThh:nn:ssZ")  '+++ изменено B3
         ' StartValue = Format(shDataAll.Range("B2"), "yyyy-mm-ddThh:mmZ")  '+++
         
         StartValue = Format(dStartTimeLine, "yyyy-mm-ddThh:mmZ")
         
         Set subNode = doc.createElement("start")
         Set tagText = doc.createTextNode(StartValue)
         subNode.appendChild (tagText)
         dataNode.appendChild (subNode)
         DataNodeSeries.appendChild dataNode
         
         ' endValue = Format(shDataAll.Range("B3"), "yyyy-mm-ddThh:nn:ssZ")   '+++ изменено B3
         ' endValue = Format(shDataAll.Range("B3"), "yyyy-mm-ddThh:mmZ")   '+++
         endValue = Format(dEndTimeLine, "yyyy-mm-ddThh:mmZ")
         
         Set subNode = doc.createElement("end")
         Set tagText = doc.createTextNode(endValue)
         subNode.appendChild (tagText)
         dataNode.appendChild (subNode)
         DataNodeSeries.appendChild dataNode
         
         dataNodeTimeSeries.appendChild DataNodeSeries
    
    '' set  Resolution
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
            ' shDataDay.Cells(2, 3 + j).Value = naek(j)
            Set tagText = doc.createTextNode(shDataAll.Cells(5, 2 + j).Value)
            
            Quantity.appendChild (tagText)
            Point.appendChild (Quantity)
         Next j

        root.appendChild dataNodeTimeSeries
   ' Next i
    
        sCurrDate = Format(dCurrDate, "yyyy-mm-dd_")
        
        Set Pi = doc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
        createdTimeValueFormated = Format(Now, "yyyy-mm-ddThh_nn_ssZ")
        xmlFileName = ThisWorkbook.Path & "\18_9.3_NNEGC.xml"
        Smooth_Xml doc
        doc.InsertBefore Pi, doc.ChildNodes.Item(0)
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



