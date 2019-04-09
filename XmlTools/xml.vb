Imports System
Imports System.IO
Imports System.Xml
Imports System.Xml.Schema
Imports System.Xml.XPath
Imports Microsoft.VisualBasic

Public Class Xml
    '
    '--------------------
    ' XML utilities class
    '--------------------
    '
    Dim strProgName As String = Nothing
    Dim strArguments As String = Nothing
    Dim strResult As String = Nothing
    Dim strMethod As String = Nothing
    Dim testPl As String = Nothing
    Dim strServerPool As String = Nothing
    Dim strValErrs As String = Nothing
    Dim strPayloadArr() As String
    Dim strTag As String = Nothing
    Dim strReaderTag As String = Nothing
    Dim reader As XmlReader
    Dim strLastElement As String
    Dim validationErrors As String = ""
    Dim LogDetails As String = Nothing
    Dim strDiagsOn As String = Nothing
    Dim strSettingsPath As String = "wslog\xmltools_Diagnostics.txt"
    Dim DiagnosticsOn As Boolean = False
    Dim strLogPath As String = Nothing
    Dim strDrive As String = Nothing
    Public Function CheckQuoteExpiry(ByVal strQuoteRef As String, ByVal ServerPool As String, ByVal lob As String, ByVal strCoverDate As String) As String
        '
        '-----------------------------------------------------
        ' Function to check that the Quotation has not expired
        '-----------------------------------------------------
        '
        strProgName = "check.quote.expiry"
        strMethod = "POST"
        strArguments = "&itemid=" & strQuoteRef
        strArguments = strArguments & "&lob=" & lob
        strArguments = strArguments & "&coverdate=" & strCoverDate
        If lob = "cv" Then
            strArguments = strArguments & "&path=MDS,PF-FLEET,"
        Else
            strArguments = strArguments & "&path=MDS,POLICYFAST,"
        End If
        strServerPool = ServerPool
        strResult = ConnectToD3()
        Return strResult
    End Function
    Public Function CompareRisk(ByVal strQuoteRef As String, ByVal strXmlPayload As String, ByVal ServerPool As String, ByVal lob As String) As String
        '
        '--------------------------------------------------------------------
        ' Function to check that the risk has not changed since the quotation
        '--------------------------------------------------------------------
        '
        strArguments = "&itemid=" & strQuoteRef
        strProgName = "compare.risk"
        strArguments = "&itemid=" & strQuoteRef
        strArguments = strArguments & "&lob=" & lob
        If lob = "cv" Then
            strArguments = strArguments & "&path=MDS,PF-FLEET,"
        Else
            strArguments = strArguments & "&path=MDS,POLICYFAST,"
        End If
        Dim strFCPayloada As String = strXmlPayload.Replace("&", "~")
        Dim strFCPayload As String = strFCPayloada.Replace("20+", "20plus")
        strArguments = strArguments & "&payload=" & strFCPayload
        strMethod = "POST"
        strServerPool = ServerPool
        strResult = ConnectToD3()
        Return strResult
    End Function
    Public Function CreatePolicy(ByVal strQuoteRef As String, ByVal ServerPool As String, ByVal userid As String, ByVal ChosenCo As String, ByVal ChosenPrem As String, ByVal lob As String, ByVal StartDate As String, ByVal DocMode As String, ByVal BrokerFee As String) As String
        '
        '--------------------------------
        ' Function to generate the policy
        '--------------------------------
        '
        '
        Dim strPolRef As String = Nothing
        Dim strDocResult As String = Nothing
        Dim intPolPos As Integer = 0
        Dim strPolArray() As String
        Dim strPolString As String = Nothing
        '
        If Environment.MachineName = "CHRISWARDLAPTOP" Then
            strDrive = "d:\"
        Else
            strDrive = "e:\"
        End If
        Try
            LogDetails = My.Computer.FileSystem.ReadAllText(strDrive & strSettingsPath)
            strDiagsOn = LogDetails.Split(",")(0)
        Catch ex As Exception
            strDiagsOn = "OFF"
        End Try

        If strDiagsOn = "ON" Then
            Dim DatePart As String = Year(Today) & Format(Month(Today), "00") & Format(Day(Today), "00")
            Dim TimePart As String = Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
            Dim rootpath As String = LogDetails.Split(",")(1)
            '    strLogPath = strDrive & rootpath & "_GetQuotes_" & DatePart & TimePart & ".txt"
            strLogPath = strDrive & "wslog\Truck\WriteBusinessProcess_" & strQuoteRef & ".txt"
            DiagnosticsOn = True
        End If
        LogMessage("Opening Log to " & strLogPath)
        '
        '--------------------------------
        ' Write Broker fee to file first
        '--------------------------------
        '
        strProgName = "update.attribute"
        strArguments = "&itemid=" & strQuoteRef
        strArguments &= "&file=cv.quotations&usedict=cv.quotations&attname=cv.brokerfee"
        strArguments &= "&attdata=" & BrokerFee
        If lob = "cv" Then
            strArguments = strArguments & "&path=MDS,PF-FLEET,"
        Else
            strArguments = strArguments & "&path=MDS,POLICYFAST,"
        End If
        strMethod = "POST"
        strServerPool = ServerPool
        strResult = ConnectToD3()
        '
        '-------------------------
        ' Now create the Policy
        '-------------------------
        '
        strProgName = "writebusiness"
        strArguments = "&itemid=" & strQuoteRef
        strArguments = strArguments & "&quotetype=" & lob
        strArguments = strArguments & "&userid=" & userid
        strArguments = strArguments & "&chosenco=" & ChosenCo
        strArguments = strArguments & "&chosenprem=" & ChosenPrem
        strArguments = strArguments & "&formatreqd=xml"
        strArguments = strArguments & "&startdate=" & StartDate
        strArguments = strArguments & "&bfee=" & BrokerFee
        If lob = "cv" Then
            strArguments = strArguments & "&path=MDS,PF-FLEET,"
        Else
            strArguments = strArguments & "&path=MDS,POLICYFAST,"
        End If
        strMethod = "POST"
        strServerPool = ServerPool
        strResult = ConnectToD3()
        '
        '-------------------------------------
        ' Now generate docs if policy returned
        '-------------------------------------
        '
        intPolPos = strResult.IndexOf("<Policy_Reference")
        strPolString = Mid(strResult, intPolPos, 99)
        strPolArray = strPolString.Split(Chr(34))
        strPolRef = strPolArray(1)
        If DocMode = "WRITE" And strPolRef <> Nothing Then
            If lob = "cv" Then
                LogMessage("XmlTools_CreatePolicy Starting Documentation")
                strProgName = "prepareprint.pffleet"
                strArguments = "&path=MDS,PF-FLEET,"
                strArguments = strArguments & "&userid=" & userid
                strArguments = strArguments & "&quotetype=" & lob
                strArguments = strArguments & "&reference=" & strPolRef
                strArguments = strArguments & "&mode=" & DocMode
                strDocResult = ConnectToD3()
            End If
        End If
        Return strResult
    End Function
    Public Function GetSection(ByVal Payload As String, ByVal lob As String, ByVal rtn_fmt As String, ByVal section As String, ByVal ServerPool As String) As String
        '
        '----------------------------------------------------------
        ' Function to extract a specified section of an XML message
        '----------------------------------------------------------
        '
        strProgName = "extract.xml.data"
        strMethod = "POST"
        Dim strFCPayload As String = Payload.Replace("&", "~")
        strArguments = "&payload=" & strFCPayload
        strArguments = strArguments & "&lob=" & lob
        strArguments = strArguments & "&rtn_fmt=" & rtn_fmt
        strArguments = strArguments & "&section=" & section
        strServerPool = ServerPool
        strResult = ConnectToD3()
        Return strResult
    End Function
    Public Function WriteQuote(ByVal Payload As String, ByVal lob As String, ByVal Section As String, ByVal ServerPool As String) As String
        '
        '--------------------------------------------------------------------
        ' Function to write the risk data and create a quotation record on D3
        '--------------------------------------------------------------------
        '
        strProgName = "write.xml.data"
        Dim strFCPayloada As String = Payload.Replace("&", "~")
        Dim strFCPayload As String = strFCPayloada.Replace("20+", "20plus")
        strArguments = "&payload=" & strFCPayload
        strArguments = strArguments & "&lob=" & lob
        strMethod = "POST"
        strServerPool = ServerPool
        strResult = ConnectToD3()
        Return strResult
    End Function
    Public Function GetQuotes(ByVal itemid As String, ByVal lob As String, ByVal userid As String, ByVal BreakdownReqd As String, ByVal ServerPool As String) As String
        '
        '------------------------------------------------------
        ' Function to return quotations for specified risk data
        '------------------------------------------------------
        '
        ' -----------------------------
        ' Set up logging
        ' -----------------------------
        '
        If Environment.MachineName = "CHRISWARDLAPTOP" Then
            strDrive = "d:\"
        Else
            strDrive = "e:\"
        End If
        Try
            LogDetails = My.Computer.FileSystem.ReadAllText(strDrive & strSettingsPath)
            strDiagsOn = LogDetails.Split(",")(0)
        Catch ex As Exception
            strDiagsOn = "OFF"
        End Try

        If strDiagsOn = "ON" Then
            Dim DatePart As String = Year(Today) & Format(Month(Today), "00") & Format(Day(Today), "00")
            Dim TimePart As String = Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
            Dim rootpath As String = LogDetails.Split(",")(1)
            strLogPath = strDrive & rootpath & "_GetQuotes_" & DatePart & TimePart & ".txt"
            DiagnosticsOn = True
        End If
        LogMessage("Opening Log to " & strLogPath)

        strProgName = "pfast.displayquote"
        strArguments = "&itemid=" & itemid
        strArguments = strArguments & "&quotetype=" & lob & "*"
        strArguments = strArguments & "&subtype=" & "truck"
        strArguments = strArguments & "&userid=" & userid
        strArguments = strArguments & "&addbreakdown=" & BreakdownReqd
        strArguments = strArguments & "&apptype=ws"
        If lob = "cv" Then
            strArguments = strArguments & "&path=MDS,PF-FLEET,"
        Else
            strArguments = strArguments & "&path=MDS,POLICYFAST,"
        End If
        strArguments = strArguments & "&formatreqd=xml"
        strMethod = "POST"
        strServerPool = ServerPool
        strResult = ConnectToD3()
        Return strResult
    End Function
    Private Function ConnectToD3() As String
        '
        '----------------------------------------------------------------------------
        ' Utility Function to connect to D3 and run specified FlashConnect subroutine
        '----------------------------------------------------------------------------
        '
        Dim Http As New MSXML.XMLHTTPRequest
        Dim RandomNumber As Random = New Random
        Dim strRandomString As String = "&random=" & RandomNumber.Next.ToString
        Dim strDomain As String = Nothing
        Select Case strServerPool
            Case "doris"
                strDomain = "http://www.datamatters.info"
            Case "wendy"
                strDomain = "http://www.coversure.co.uk"
        End Select

        Dim ConnectionString As String = strDomain & "/cgi-bin/fccgi.exe?w3exec=" + strProgName
        Dim strArg As String = Nothing
        Dim strResult As String = Nothing

        If strServerPool <> Nothing Then
            strArg = "&w3serverpool=" & strServerPool & strArguments & strRandomString
        Else
            strArg = strArguments & strRandomString
        End If

        If strMethod = Nothing Then
            strMethod = "GET"
        End If
        Try
            Http.open(strMethod, ConnectionString, False)
            Http.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
            Http.setRequestHeader("Timeout", "600000")
            Http.send(strArg)
        Catch ex As Exception
            strResult = strProgName & " - " & ex.Message
            GoTo ExitPoint
        End Try
        strResult = Http.responseText
ExitPoint:
        Return strResult
    End Function
    Public Function ProcessMTA(ByVal strPayload As String, ByVal strPolRef As String, ByVal ServerPool As String, ByVal lob As String, ByVal strUserId As String, ByVal strAdjType As String, ByVal strAdjNotes As String) As String
        '
        '------------------------------
        ' Function to Process MTA on D3
        '------------------------------
        '
        Dim strFCPayload As String = strPayload.Replace("&", "~")
        strProgName = "ws.process.mta"
        strMethod = "POST"
        Dim strHeader As String = "internet"
        Dim strOption As String = "adj"
        Dim strDocResult As String = Nothing
        Dim strAdjustNow As String = Nothing
        Dim strDocsReqd As String = Nothing

        strArguments = "&itemid=" & strPolRef
        strArguments = strArguments & "&pol.ref=" & strPolRef
        strArguments = strArguments & "&lob=" & lob
        strArguments = strArguments & "&option=" & strOption
        strArguments = strArguments & "&header=" & strHeader
        strArguments = strArguments & "&userid=" & strUserId
        strArguments = strArguments & "&adjnotes=" & strAdjNotes
        strArguments = strArguments & "&payload=" & strFCPayload
        strArguments = strArguments & "&adjtype=" & strAdjType
        strArguments = strArguments & "&formatreqd=xml"

        If lob = "cv" Then
            strArguments = strArguments & "&path=MDS,PF-FLEET,"
        Else
            strArguments = strArguments & "&path=MDS,POLICYFAST,"
        End If

        strServerPool = ServerPool
        strResult = ConnectToD3()
        'if adjustnow = "Y" and docs.reqd = "Y" then
        strAdjustNow = ExtractFromXml("Header_AdjustNow", strFCPayload, "Val")
        strDocsReqd = ExtractFromXml("Header_DocsReqdInd", strFCPayload, "Val")

        If strAdjustNow = "Y" And strDocsReqd = "Y" And strPolRef <> Nothing Then
            If lob = "cv" Then
                strProgName = "prepareprint.pffleet"
                strArguments = "&path=MDS,PF-FLEET,"
                strArguments = strArguments & "&userid=" & strUserId
                strArguments = strArguments & "&quotetype=" & lob
                strArguments = strArguments & "&reference=" & strPolRef
                strArguments = strArguments & "&mode=" & strOption
                strDocResult = ConnectToD3()
            End If
        End If

        Return strResult
    End Function

    Public Function CancelPolicy(ByVal strPayload As String, ByVal ServerPool As String, ByVal lob As String) As String
        '
        '--------------------------------
        ' Function to cancel Policy on D3
        '--------------------------------
        '
        strProgName = "CancelPolicy"
        strMethod = "POST"
        Dim strFCPayload As String = strPayload.Replace("&", "~")
        strArguments = "&payload=" & strFCPayload
        If lob = "cv" Then
            strArguments = strArguments & "&path=MDS,PF-FLEET,"
        Else
            strArguments = strArguments & "&path=MDS,POLICYFAST,"
        End If

        strArguments &= "&formatreqd=xml"
        strServerPool = ServerPool
        strResult = ConnectToD3()
        Return strResult
    End Function
    Public Function ValidateXML(ByVal strPayload As String, ByVal strSchemaPath As String) As String
        '
        '----------------------------------------------------
        ' Function to validate XML message against XSD Schema
        '----------------------------------------------------
        '
        Dim strValidationResults As String = Nothing
        Dim settings As XmlReaderSettings = New XmlReaderSettings()
        '  Dim aStringReader As TextReader

        settings.Schemas.Add("", strSchemaPath)
        settings.ValidationType = ValidationType.Schema
        settings.ValidationFlags = settings.ValidationFlags Or XmlSchemaValidationFlags.ReportValidationWarnings
        AddHandler settings.ValidationEventHandler, AddressOf ValidationCallBack
        '   strPayloadArr = strPayload.Split("<")
        checkxml(strPayload, strSchemaPath)
        'aStringReader = New StringReader(strPayload)

        'reader = XmlReader.Create(aStringReader, settings)
        '' Dim lineInfo As IXmlLineInfo = CType(reader, IXmlLineInfo)
        'If reader.Name <> "Val" And reader.Name <> Nothing Then
        '    strReaderTag = reader.Name
        'End If
        'While (reader.Read())
        '    If reader.Name <> "Val" And reader.Name <> Nothing Then
        '        strTag = reader.Name
        '    End If
        '    If reader.NodeType = XmlNodeType.Element Then
        '        If reader.Name <> "Val" And reader.Name <> Nothing Then
        '            strLastElement = reader.Name
        '        End If
        '    End If
        'End While
        strValidationResults = strValErrs
        If strValidationResults = Nothing Then
            strValidationResults = "valid"
        End If

        Return strValidationResults
    End Function
    Public Sub ValidationEventHandler(ByVal sender As Object, ByVal e As ValidationEventArgs)
        Dim zz As System.Xml.Schema.XmlSchemaValidationException = New System.Xml.Schema.XmlSchemaValidationException
        zz = e.Exception
        Dim ff As System.Type = zz.SourceObject.GetType
        Dim ww As String = Nothing
        If ff.Name = "XmlAttribute" Then
            Dim yy As System.Xml.XmlAttribute = zz.SourceObject
            Dim xx As System.Xml.XmlElement = yy.OwnerElement
            ww = xx.Name
        End If
        If ff.Name = "XmlElement" Then
            Dim tt As System.Xml.XmlElement = zz.SourceObject
            ww = tt.Name
        End If
        strValErrs &= "~Element " & ww & ": " & e.Message
    End Sub

    Public Sub ValidationCallBack(ByVal sender As Object, ByVal e As ValidationEventArgs)
        ' Dim strPayloadSegment As String = strPayloadArr(strTag)

        'Dim strPayloadTag() As String = strPayloadSegment.Split(" ")

        'If e.Message.IndexOf("invalid child element", 0) > -1 Then
        '    strValErrs &= "~Element " & strTag & ": " & e.Message
        'Else
        '    strValErrs &= "~Element " & strLastElement & ": " & e.Message
        'End If
        strValErrs &= "~Element " & strLastElement & "/" & strTag & ": " & e.Message
    End Sub
    Function checkxml(ByVal strPayload As String, ByVal strSchemaPath As String)
        Try
            Dim objSchemasColl As New System.Xml.Schema.XmlSchemaSet
            objSchemasColl.Add("", strSchemaPath)
            Dim xmlDocument As New XmlDocument
            xmlDocument.LoadXml(strPayload)
            xmlDocument.Schemas.Add(objSchemasColl)
            xmlDocument.Validate(AddressOf ValidationEventHandler)
            If validationErrors = "" Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    Public Function GetLogDetails(ByVal strEnv As String, ByVal strSelCriteria As String) As String
        Dim strResults As String = Nothing
        Select Case strEnv
            Case "T"
                strServerPool = "doris"
            Case "L"
                strServerPool = "wendy"
        End Select

        strArguments = "&selcriteria=" & strSelCriteria
        strProgName = "SelectLogRecords"
        strMethod = "POST"
        strResults = ConnectToD3()
        Return strResults
    End Function
    Public Function BuildResponseHeader(ByVal ErrorMsg As String, ByVal strXmlRequest As String) As String
        Dim strRespDate As String = Format(Today, "dd/MM/yyyy")
        Dim strRespTime As String = Format(Now, "HH:mm:ss")

        '
        ' -----------------------------
        ' Set up logging
        ' -----------------------------
        '
        Try
            LogDetails = My.Computer.FileSystem.ReadAllText(strSettingsPath)
            strDiagsOn = LogDetails.Split(",")(0)
        Catch ex As Exception
            strDiagsOn = "OFF"
        End Try

        If strDiagsOn = "ON" Then
            Dim DatePart As String = Year(Today) & Format(Month(Today), "00") & Format(Day(Today), "00")
            Dim TimePart As String = Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
            Dim rootpath As String = LogDetails.Split(",")(1)
            strLogPath = rootpath & "_BuildResponseHeader_" & DatePart & TimePart & ".txt"
            DiagnosticsOn = True
        End If
        LogMessage("Opening Log to " & strLogPath)
        Dim strEnvRef As String = Nothing
        Dim strUserName As String = ExtractFromXml("Header_UserName", strXmlRequest, "Val")
        If strUserName = "insurecom" Or Mid(strUserName, 1, 4) = "ICL-" Then
            strEnvRef = ExtractFromXml("Header_Environment", strXmlRequest, "Val")
        Else
            Dim strEnvPath As String = Environment.CurrentDirectory & "\env.txt"
            Try
                strEnvPath = XmlTools.My.Application.Info.DirectoryPath
            Catch ex As Exception
            End Try
            LogMessage("env dir = " & strEnvPath)
            If strEnvPath.IndexOf("secure_chris_ct_csure_truck", 1) > -1 Then
                strEnvRef = "T"
            End If
            If strEnvPath.IndexOf("secure_ws_ct_csure_truck", 1) > -1 Then
                strEnvRef = "L"
            End If
        End If

        LogMessage("strEnvRef= " & strEnvRef)
        '
        Dim strTransRef As String = "9999"
        Dim strRespHeader As String = "<QuoteResults>"
        strRespHeader &= "<Header>"
        strRespHeader &= "<Header_Environment Val=""" & strEnvRef & """/>"
        strRespHeader &= "<Header_TransRef Val= """ & strTransRef & """/>"
        strRespHeader &= "<Header_TransDate Val=""" & strRespDate & """/>"
        strRespHeader &= "<Header_TransTime Val=""" & strRespTime & """/>"
        strRespHeader &= "</Header>"
        strRespHeader &= ErrorMsg
        strRespHeader &= "</QuoteResults>"
        Return strRespHeader
    End Function
    Public Function ExtractFromXml(ByVal strElement As String, ByVal strXmlRequest As String, ByVal strType As String) As String
        '
        '-------------------------------------------
        ' Utility to extract fields from XML Payload
        '-------------------------------------------
        '
        '--------------------------------
        ' Can also extract from attribute
        '--------------------------------
        '
        Dim strXmlField As String = Nothing
        Dim IntTagStart As Integer = 0
        Dim IntTagEnd As Integer = 0
        Dim strTemp As String = Nothing
        Dim strTempArray As String()
        Dim intExtractLen As Integer = 0
        If strType = "Val" Then
            IntTagStart = strXmlRequest.IndexOf(strElement)
            If IntTagStart < 0 Then
                strXmlField = "Error:[S1.11]Tag " & strElement & " not found"
            Else
                intExtractLen = strXmlRequest.Length - IntTagStart
                strTemp = Mid(strXmlRequest, IntTagStart, intExtractLen)
                strTempArray = strTemp.Split(Chr(34))
                strXmlField = strTempArray(1)
            End If
        Else
            IntTagStart = strXmlRequest.IndexOf("<" & strElement, 0)
            IntTagEnd = strXmlRequest.IndexOf("</" & strElement, 1)
            If intTagStart > -1 Then intTagStart += 1
            strTemp = Mid(strXmlRequest, IntTagStart, IntTagEnd)
            strTempArray = strTemp.Split(">")
            strTemp = strTempArray(1)
            strTempArray = strTemp.Split("<")
            strXmlField = strTempArray(0)
        End If
        Return strXmlField
    End Function
    Public Sub FormatXml(ByVal strRawXml As String, ByRef strFormattedXml As String, ByRef strError As String)
        '
        '-----------------------------
        ' Utility to format XML string
        '-----------------------------
        '
        Dim xd As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        Dim strResult As String = Nothing
        Try
            xd.LoadXml(strRawXml)
        Catch ex As Exception
            strError = "Error:[S1.7]Formatxml " & ex.Message
        End Try
        If strError = Nothing Then
            Dim sb As System.Text.StringBuilder = New System.Text.StringBuilder
            Dim sw As System.IO.StringWriter = New System.IO.StringWriter(sb)
            Dim xtw As System.Xml.XmlTextWriter = Nothing
            xtw = New System.Xml.XmlTextWriter(sw)
            xtw.Formatting = System.Xml.Formatting.Indented
            xd.WriteTo(xtw)
            strFormattedXml = sb.ToString
        End If

    End Sub
    Public Sub stripchars(ByVal strXmlFile As String, ByRef newXml As String)
        '
        '-------------------------------------
        ' Utility to strip characters from XML
        '-------------------------------------
        '
        Dim intCharCount As Integer = strXmlFile.Length
        Dim intCount As Integer = 0
        Dim strTestChar As String = Nothing
        Dim blnFormatArea As Boolean = False

        For intCount = 1 To intCharCount
            strTestChar = Mid(strXmlFile, intCount, 1)
            Select Case strTestChar
                Case "<"
                    blnFormatArea = False
                Case ">"
                    newXml = newXml & strTestChar
                    blnFormatArea = True
                Case Else
            End Select
            If blnFormatArea = False Then
                newXml = newXml & strTestChar
            End If
        Next
    End Sub
    Protected Sub ReadData(ByVal strPath As String, ByRef strRecord As String)
        '
        '****************************************************
        ' Read data from file, path selected, record returned
        '****************************************************
        '
        Dim ln As String
        Dim sr As System.IO.StreamReader = New System.IO.StreamReader(strPath)
        Do
            ln = sr.ReadLine()
            strRecord = strRecord & ln
        Loop Until ln Is Nothing
        sr.Close()
        sr.Dispose()
    End Sub
    Private Sub LogMessage(ByVal ErrorMsg As String)
        If DiagnosticsOn = True Then
            Dim LogDate As Date = Today
            Dim LogTime As Date = Now
            FileOpen(1, strLogPath, OpenMode.Append)
            Print(1, LogTime.ToString & " " & ErrorMsg & vbCrLf)
            FileClose(1)
        End If
    End Sub
End Class
