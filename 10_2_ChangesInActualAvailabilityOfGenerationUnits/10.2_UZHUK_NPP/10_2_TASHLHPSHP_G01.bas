Attribute VB_Name = "Module1"
' Dim naek(1 To 24) As Integer


' --- Створення файлу XML для завантаженя на TRANSPARENCY PLATFORM ENTSO-E
Sub Main()
    Dim doc As MSXML2.DOMDocument60
    Dim root As MSXML2.IXMLDOMElement, dataNode As MSXML2.IXMLDOMElement
    Dim i As Long
    Dim dDt As Object
    ' Dim shDataAll, shDataDay As Worksheet
    Dim shDataAll As Worksheet
    Dim inQuantity As Integer
    
    Dim strmRID, strSShNumber, strGeneratorNumber, strCurrentDate, strFromDate As String
    strSShNumber = "62W191679871593G"  ' SSh-330
    strGeneratorNumber = "62W487981668344S" ' TASHLUK-1
    strFromDate = "31.12.2019"
    strCurrentDate = Date
    strmRID = strGeneratorNumber & "-" & DateDiff("y", strFromDate, strCurrentDate)
    
    Set doc = New MSXML2.DOMDocument60
    Set shDataAll = Worksheets("data")
    ' Set shDataDay = Worksheets("TimeSeries")

    Set root = doc.createElement("Unavailability_MarketDocument")
    doc.appendChild root
   
    AddAttributeWithValue root, "xmlns", "urn:iec62325.351:tc57wg16:451-6:outagedocument:3:0"
    Set att = doc.createAttribute("xmlns")
  
    att.Value = "urn:iec62325.351:tc57wg16:451-6:outagedocument:3:0"
    
    
    '''Set header'''
    
    '''mRID
        Set mRIDNode = doc.createElement("mRID")
        
        mRIDValue = strmRID  ' shDataAll.Range("H4").Text
        shDataAll.Range("H4").Value = strmRID
        
        Set tagText = doc.createTextNode(mRIDValue)
        mRIDNode.appendChild (tagText)
        root.appendChild mRIDNode
    
    '''revisionNumber
        Set dataNode = doc.createElement("revisionNumber")
        revisionNumberValue = shDataAll.Range("I4").Text
        Set tagText = doc.createTextNode(revisionNumberValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode
    
    ''type
        Set dataNode = doc.createElement("type")
        typeValue = "A80"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''process.processType
        Set dataNode = doc.createElement("process.processType")
        typeValue = "A26"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''createdDateTime
        Set dataNode = doc.createElement("createdDateTime")
        Set dDt = CreateObject("WbemScripting.SWbemDateTime")

            dDt.SetVarDate Now
            Now_ = dDt.GetVarDate(False)

        ' Now_ = Now
        ' DateTime DateNow = DateTime.Now
        ' createdTimeValue = TimeZoneInfo.ConvertTimeToUtc(DateTime.Now)
        
        createdTimeValueFormated = Format(Now_, "yyyy-mm-ddThh:nn:ssZ")
        Set tagText = doc.createTextNode(createdTimeValueFormated)
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

    '' unavailabily_Time_Period.timeInterval
        Set dataNode = doc.createElement("unavailability_Time_Period.timeInterval")        ' --- изменено

       
        StartValue = Format(shDataAll.Range("D4").Text & " " & shDataAll.Range("E4").Text, "yyyy-mm-ddThh:mmZ")  '+++
        
        ' StartValue = Format(shDataAll.Range("D4").Text & " " & shDataAll.Range("E4").Text, "yyyy-mm-dd hh:mm ")  '+++
        
'        Set dDt = CreateObject("WbemScripting.SWbemDateTime")
'            ' dDt.SetVarDate DateValue(StartValue)
'            dDt.SetVarDate shDataAll.Range("D5").Value
'            StartValue = dDt.GetVarDate(False)
        
   '     createdTimeValueFormatted = Format(StartValue, "yyyy-mm-ddThh:mmZ")
        Set subNode = doc.createElement("start")
        Set tagText = doc.createTextNode(StartValue)
        subNode.appendChild (tagText)
        dataNode.appendChild (subNode)

       
        endValue = Format(shDataAll.Range("F4").Text & " " & shDataAll.Range("G4").Text, "yyyy-mm-ddThh:mmZ")   '+++
      '  endValue = Format(shDataAll.Range("F4").Text & " " & shDataAll.Range("G4").Text, "yyyy-mm-dd hh:mm ")   '+++
        
'        Set dDt = CreateObject("WbemScripting.SWbemDateTime")
'            dDt.SetVarDate DateValue(endValue)
'            endValue = dDt.GetVarDate(False)

 '       createdTimeValueFormatted = Format(endValue, "yyyy-mm-ddThh:mmZ")
        Set subNode = doc.createElement("end")
        Set tagText = doc.createTextNode(endValue)
        subNode.appendChild (tagText)
        dataNode.appendChild (subNode)
        root.appendChild dataNode

'' set TimeSeries - loop
        Set dataNodeTimeSeries = doc.createElement("TimeSeries")

    '' set mRID
        Set DataNodeSeries = doc.createElement("mRID")
        mRIDValue = "1"

        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set businessType
        Set DataNodeSeries = doc.createElement("businessType")
        mRIDValue = "A54"
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set area_Domain.mRID
        Set NodeSeries = doc.createElement("biddingZone_Domain.mRID")
        mRIDValue = "10Y1001C--000182"
        Set tagText = doc.createTextNode(mRIDValue)
        NodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (NodeSeries)

        Set att = doc.createAttribute("codingScheme")
        att.Value = "A01"
        NodeSeries.setAttributeNode att
        
    ''set start_DateAndOrTime.date
        Set DataNodeSeries = doc.createElement("start_DateAndOrTime.date")
        mRIDValue = Format(shDataAll.Range("D4").Text, "yyyy-mm-dd")
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)
    ''set start_DateAndOrTime.time
        Set DataNodeSeries = doc.createElement("start_DateAndOrTime.time")
        mRIDValue = Format(shDataAll.Range("E4").Text, "hh:mm:ssZ")
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)
   ''set end_DateAndOrTime.date
        Set DataNodeSeries = doc.createElement("end_DateAndOrTime.date")
        mRIDValue = Format(shDataAll.Range("F4").Text, "yyyy-mm-dd")
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)
    ''set end_DateAndOrTime.time
        Set DataNodeSeries = doc.createElement("end_DateAndOrTime.time")
        mRIDValue = Format(shDataAll.Range("G4").Text, "hh:mm:ssZ")
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set measure_Unit.name
        Set DataNodeSeries = doc.createElement("quantity_Measure_Unit.name")
        mRIDValue = "MAW"
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set curveType
        Set DataNodeSeries = doc.createElement("curveType")
        txtCurveType = "A03"
        Set tagText = doc.createTextNode(txtCurveType)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''production_RegisteredResource.mRID
        Set dataNode = doc.createElement("production_RegisteredResource.mRID")
        typeValue = strSShNumber ' "62W9342984333277"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        dataNodeTimeSeries.appendChild dataNode
        AddAttributeWithValue dataNode, "codingScheme", "A01"

    ''production_RegisteredResource.pSRType.powerSystemResources.mRID
        Set dataNode = doc.createElement("production_RegisteredResource.pSRType.powerSystemResources.mRID")
        typeValue = strGeneratorNumber ' "62W086969517524B"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        dataNodeTimeSeries.appendChild dataNode
        AddAttributeWithValue dataNode, "codingScheme", "A01"
        
        
     '' set Available_Period
        Set DataNodeSeries = doc.createElement("Available_Period")
    '' timeInterval
         Set dataNode = doc.createElement("timeInterval")
         ' StartValue = Format(shDataAll.Range("B2"), "yyyy-mm-ddThh:nn:ssZ")  '+++ изменено B3
         StartValue = Format(shDataAll.Range("D4").Text & " " & shDataAll.Range("E4").Text, "yyyy-mm-ddThh:mmZ")  '+++
         
         Set subNode = doc.createElement("start")
         Set tagText = doc.createTextNode(StartValue)
         subNode.appendChild (tagText)
         dataNode.appendChild (subNode)
         DataNodeSeries.appendChild dataNode
         
         ' endValue = Format(shDataAll.Range("B3"), "yyyy-mm-ddThh:nn:ssZ")   '+++ изменено B3
         endValue = Format(shDataAll.Range("F4").Text & " " & shDataAll.Range("G4").Text, "yyyy-mm-ddThh:mmZ")   '+++
         
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

'         For j = 1 To 24
            Set Point = doc.createElement("Point")
            DataNodeSeries.appendChild Point

            Set Position = doc.createElement("position")
            Set tagText = doc.createTextNode(1)
            Position.appendChild (tagText)
            Point.appendChild (Position)

            Set Quantity = doc.createElement("quantity")
'
'            ' Set tagText = doc.createTextNode(shDataDay.Cells(27 + i, 3 + j).Text)
'            Select Case i
'            Case 1 To 2
'                shDataDay.Cells(1 + i, 3 + j).Value = khm(i, j)
'                Set tagText = doc.createTextNode(khm(i, j))
'            Case 3 To 5
'                shDataDay.Cells(1 + i, 3 + j).Value = uzhuk(i - 2, j)
'                Set tagText = doc.createTextNode(uzhuk(i - 2, j))
'            Case 6 To 11
'                shDataDay.Cells(1 + i, 3 + j).Value = zap(i - 5, j)
'                Set tagText = doc.createTextNode(zap(i - 5, j))
'            Case 12 To 17
'                shDataDay.Cells(1 + i, 3 + j).Value = rivn(i - 11, j)
'                Set tagText = doc.createTextNode(rivn(i - 11, j))
'
'            Case 18
'                shDataDay.Cells(1 + i, 3 + j).Value = oges(j)
                Set tagText = doc.createTextNode(shDataAll.Range("B4").Text)
'            End Select
'
            Quantity.appendChild (tagText)
            Point.appendChild (Quantity)
'         Next j
        
        
     '' set Reason
        Set ReasonObj = doc.createElement("Reason")
        dataNodeTimeSeries.appendChild ReasonObj

        Set Code = doc.createElement("code")
        Set tagText = doc.createTextNode("A95")
        Code.appendChild (tagText)
        ReasonObj.appendChild (Code)

        Set Text = doc.createElement("text")
        Set tagText = doc.createTextNode(shDataAll.Range("C4").Text)
        Text.appendChild (tagText)
        ReasonObj.appendChild (Text)
        
        root.appendChild dataNodeTimeSeries
        
    '' set Reason
        Set ReasonObj1 = doc.createElement("Reason")
        Set Code = doc.createElement("code")
        Set tagText = doc.createTextNode("B18")
        Code.appendChild (tagText)
        ReasonObj1.appendChild (Code)
        root.appendChild ReasonObj1
    
        sCurrDate = Format(dCurrDate, "yyyy-mm-dd_")
        
        Set Pi = doc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
        createdTimeValueFormated = Format(Now, "yyyy-mm-ddThh_nn_ssZ")
        xmlFileName = ThisWorkbook.Path & "\04_10.2_UZHUKNPP-U04.xml"
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


