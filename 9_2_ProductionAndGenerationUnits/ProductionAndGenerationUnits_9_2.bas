Attribute VB_Name = "Module1"
' --- Створення файлу XML для завантаженя на TRANSPARENCY PLATFORM ENTSO-E
Sub Main()
    Dim doc As MSXML2.DOMDocument60
    Dim root As MSXML2.IXMLDOMElement, dataNode As MSXML2.IXMLDOMElement
    Dim i As Long
    Dim dDt As Object
    Dim shDataAll, shDataDay As Worksheet
    Dim inQuantity As Integer
    Dim Pointer As Integer
    
    Dim strmRID, strSShNumber, strGeneratorNumber, strCurrentDate, strFromDate As String
    strGeneratorNumber = "62X205270350215R" ' NNEGC
    strFromDate = "31.12.2019"
    strCurrentDate = Date
    strmRID = strGeneratorNumber & "-EA-" & Format(Now, "yyyy-mm-dd") ' DateDiff("y", strFromDate, strCurrentDate)

    Set doc = New MSXML2.DOMDocument60
    Set shDataAll = Worksheets("data")
   
    Set root = doc.createElement("Configuration_MarketDocument")
    doc.appendChild root
   
    AddAttributeWithValue root, "xmlns", "urn:iec62325.351:tc57wg16:451-6:configurationdocument:3:0"
    Set att = doc.createAttribute("xmlns")
  
    att.Value = "urn:iec62325.351:tc57wg16:451-6:configurationdocument:3:0"
    
    '''Set header'''
    
    '''mRID
        Set mRIDNode = doc.createElement("mRID")
        mRIDValue = strmRID ' "1"
        Set tagText = doc.createTextNode(mRIDValue)
        mRIDNode.appendChild (tagText)
        root.appendChild mRIDNode
    
    ''type
        Set dataNode = doc.createElement("type")
        typeValue = "A95"
        Set tagText = doc.createTextNode(typeValue)
        dataNode.appendChild (tagText)
        root.appendChild dataNode

    ''process.processType
        Set dataNode = doc.createElement("process.processType")
        typeValue = "A36"
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
        createdTimeValue = Now_
        createdTimeValueFormated = Format(Now_, "yyyy-mm-ddThh:nn:ssZ")
        Set tagText = doc.createTextNode(createdTimeValueFormated)
        dataNode.appendChild (tagText)
        root.appendChild dataNode
    
    Pointer = 14
'' set TimeSeries - loop
    For i = 1 To 8
        Set dataNodeTimeSeries = doc.createElement("TimeSeries")

    '' set mRID
        Set DataNodeSeries = doc.createElement("mRID")
        mRIDValue = "1"
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set businessType
        Set DataNodeSeries = doc.createElement("businessType")
        mRIDValue = "B11"
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)

    ''set implementation_DateAndOrTime.date
        Set DataNodeSeries = doc.createElement("implementation_DateAndOrTime.date")
        mRIDValue = Format(shDataAll.Cells(3 + i, 13), "yyyy-dd-mm") ' "2020-08-25"
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
        
     ''set registeredResource.mRID
        Set NodeSeries = doc.createElement("registeredResource.mRID")
        mRIDValue = shDataAll.Cells(3 + i, 4) ' "62W0459599839277"
        Set tagText = doc.createTextNode(mRIDValue)
        NodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (NodeSeries)
       
        Set att = doc.createAttribute("codingScheme")
        att.Value = "A01"
        NodeSeries.setAttributeNode att
        
     ''set registeredResource.name
        Set DataNodeSeries = doc.createElement("registeredResource.name")
        mRIDValue = shDataAll.Cells(3 + i, 3) '  "HAES NPP SSh-750"
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)
        
      ''set registeredResource.location.name
        Set DataNodeSeries = doc.createElement("registeredResource.location.name")
        mRIDValue = shDataAll.Cells(3 + i, 5) ' "Netishin"
        Set tagText = doc.createTextNode(mRIDValue)
        DataNodeSeries.appendChild (tagText)
        dataNodeTimeSeries.appendChild (DataNodeSeries)
        
    ''set ControlArea_Domain
        Set ControlArea = doc.createElement("ControlArea_Domain")
        Set mRID = doc.createElement("mRID")
        Set tagText = doc.createTextNode("10Y1001C--000182")
        mRID.appendChild (tagText)
        
        Set att = doc.createAttribute("codingScheme")
        att.Value = "A01"
        mRID.setAttributeNode att
        
        ControlArea.appendChild (mRID)
        dataNodeTimeSeries.appendChild ControlArea
        
     ''set Provider_MarketParticipant
        Set Provider = doc.createElement("Provider_MarketParticipant")
        Set mRID = doc.createElement("mRID")
        Set tagText = doc.createTextNode("62X205270350215R")
        mRID.appendChild (tagText)
        
        Set att = doc.createAttribute("codingScheme")
        att.Value = "A01"
        mRID.setAttributeNode att
        
        Provider.appendChild (mRID)
        dataNodeTimeSeries.appendChild Provider
        
    ''set MktPSRType
        Set mktPSRType = doc.createElement("MktPSRType")
        
        Set psrType = doc.createElement("psrType")
        Set tagText = doc.createTextNode(shDataAll.Cells(3 + i, 11).Text) ' doc.createTextNode("B14")
        psrType.appendChild (tagText)
        mktPSRType.appendChild (psrType)
        dataNodeTimeSeries.appendChild mktPSRType
        
    ''set Production_PowerSystemResources
        Set Production = doc.createElement("production_PowerSystemResources.highVoltageLimit")
        Set tagText = doc.createTextNode(shDataAll.Cells(3 + i, 10)) ' doc.createTextNode("750")
        Production.appendChild (tagText)
        
        Set att = doc.createAttribute("unit")
        att.Value = "KVT"
        Production.setAttributeNode att
        
        mktPSRType.appendChild (Production)
        dataNodeTimeSeries.appendChild mktPSRType
        
    ''set NominalIP_PowerSystemResources
        Set Nominal = doc.createElement("nominalIP_PowerSystemResources.nominalP")
        Set tagText = doc.createTextNode(shDataAll.Cells(3 + i, 9)) ' doc.createTextNode("1000")
        Nominal.appendChild (tagText)
        
        Set att = doc.createAttribute("unit")
        att.Value = "MAW"
        Nominal.setAttributeNode att
        
        mktPSRType.appendChild (Nominal)
        dataNodeTimeSeries.appendChild mktPSRType
    
    '' --- set GeneratingUnit_PowerSystemResources ---
        For g = 1 To CInt(shDataAll.Cells(3 + i, 17))
            Set GeneratingUnit = doc.createElement("GeneratingUnit_PowerSystemResources")
        
        ''set mRID
            Set mRID = doc.createElement("mRID")
            Set tagText = doc.createTextNode(shDataAll.Cells(Pointer + g, 8)) ' doc.createTextNode("62W642539965223L")
            mRID.appendChild (tagText)
        
            Set att = doc.createAttribute("codingScheme")
            att.Value = "A01"
            mRID.setAttributeNode att
        
            GeneratingUnit.appendChild (mRID)
        
        ''set Name
            Set Name = doc.createElement("name")
            nameValue = shDataAll.Cells(Pointer + g, 7) ' "HAES_NPP Energoblok-2"
            Set tagText = doc.createTextNode(nameValue)
            Name.appendChild (tagText)
            GeneratingUnit.appendChild (Name)
        
        ''set NominalP
            Set NominalP = doc.createElement("nominalP")
            Set tagText = doc.createTextNode(shDataAll.Cells(Pointer + g, 9)) ' doc.createTextNode("1000")
            NominalP.appendChild (tagText)
        
            Set att = doc.createAttribute("unit")
            att.Value = "MAW"
            NominalP.setAttributeNode att
        
            GeneratingUnit.appendChild (NominalP)
       
        ''set psrType
            Set psrType = doc.createElement("generatingUnit_PSRType.psrType")
            psrTypeValue = shDataAll.Cells(Pointer + g, 11).Text ' "B14"
            Set tagText = doc.createTextNode(psrTypeValue)
            psrType.appendChild (tagText)
            GeneratingUnit.appendChild (psrType)
        
        ''set Location
            Set Location = doc.createElement("generatingUnit_Location.name")
            locationValue = shDataAll.Cells(Pointer + g, 5) ' "Netishin"
            Set tagText = doc.createTextNode(locationValue)
            Location.appendChild (tagText)
            GeneratingUnit.appendChild (Location)

            mktPSRType.appendChild (GeneratingUnit)
        Next g
            
        dataNodeTimeSeries.appendChild mktPSRType
        root.appendChild dataNodeTimeSeries
        
        Pointer = Pointer + shDataAll.Cells(3 + i, 17)
     Next i
    
        sCurrDate = Format(dCurrDate, "yyyy-mm-dd_")
        
        Set Pi = doc.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
        createdTimeValueFormated = Format(Now, "yyyy-mm-ddThh_nn_ssZ")
        xmlFileName = ThisWorkbook.Path & "\18_9.2_NNEGC.xml"
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

