<?xml version="1.0" encoding="utf-8" ?>
<%
Dim LoginGuid
Dim SessionKey
Dim PageGuid
Dim DefaultExtension
Dim PageId
Dim CreateDate
Dim CreateBy
Dim ModifiedDate
Dim ModifiedBy

Response.ContentType = "text/xml"

LoginGuid = Session("loginguid")
SessionKey = Session("sessionkey")
PageGuid = request.querystring("pageguid")
DefaultExtension = request.querystring("defaultExtension")
PageId = ""
CreateDate = ""
CreateBy = ""
ModifiedDate = ""
ModifiedBy = ""

Function GetPageStatus()
    dim objEList
    dim xmlData
    dim n
    dim ElementGUID 
    dim sError
    dim pageStatus
    dim lockedguid 
    pageStatus = 0
    set XMLDom = Server.CreateObject("RDCMSAspObj.RDObject")
    set RQLObject = Server.CreateObject("RDCMSServer.XmlServer")     
    set objXMLDOM=Server.CreateObject("Microsoft.XMLDOM")
    xmlData = "<IODATA loginguid=""" & LoginGuid & """ sessionkey=""" & SessionKey & """><PAGE action=""load"" guid=""" & PageGuid & """ /></IODATA>" 
    xmlData = RQLObject.Execute(xmlData, sError) 
    objXMLDOM.LoadXml( xmlData ) 
    set objEList=objXMLDOM.getElementsByTagName("PAGE")
    For n = 0 to (objEList.Length-1)
        pageStatus = CLng(objEList.Item(n).getAttribute("flags"))   
		PageId = objEList.Item(n).getAttribute("id")
		CreateDate = objEList.Item(n).getAttribute("createdate")
		CreateBy = objEList.Item(n).getAttribute("createuserguid")
		ModifiedDate = objEList.Item(n).getAttribute("changedate")
		ModifiedBy = objEList.Item(n).getAttribute("changeuserguid")
        lockedguid = objEList.Item(n).getAttribute("lockuserguid")
        exit for
    Next
    bitvalues = array(268435456, 134217728, 8388608, 2097152, 524288, 262144, 131072, 8192, 1024, 64, 4)
    bitdescriptions = array("Locked", _
                            "Waiting for release","", _
                            "", _
                            "Released", _
                            "Draft", _
                            "Waiting for correction", _
                            "", _
                            "", _
                            "Waiting for release", _
                            "")   
    bitclasses = array("locked", _
                            "waiting-release","", _
                            "", _
                            "released", _
                            "draft", _
                            "waiting-correction", _
                            "", _
                            "", _
                            "waiting-release", _
                            "")    
    foundItems = ""
    i = 0
    For i = 0 To UBound( bitvalues )
        if Clng(bitvalues(i)) <= pageStatus then
            foundItems = bitdescriptions(i)
            pageStatus = pageStatus - bitvalues(i)
            if foundItems = "Locked" then
                foundItems = foundItems & " - " & GetUserFullName(lockedguid)         
            end if
            if foundItems <> "" then
                exit for
            end if
        end if
    Next
	
    GetPageStatus = "<span class=""status-icon "& bitclasses(i) &""">" & foundItems & "</span>" 
End Function

Function GetUserFullName(UserGuid)
    dim objEList
    dim xmlData
    dim n
    dim ElementGUID 
    dim sError
    dim username
    dim emailaddress
    set XMLDom = Server.CreateObject("RDCMSAspObj.RDObject")
    set RQLObject = Server.CreateObject("RDCMSServer.XmlServer")     
    set objXMLDOM=Server.CreateObject("Microsoft.XMLDOM")
    xmlData = "<IODATA loginguid=""" & LoginGuid & """><ADMINISTRATION><USER action=""load"" guid=""" & UserGuid & """ /></ADMINISTRATION></IODATA>" 
    xmlData = RQLObject.Execute(xmlData, sError) 
    objXMLDOM.LoadXml( xmlData ) 
    set objEList=objXMLDOM.getElementsByTagName("USER")
    For n = 0 to (objEList.Length-1)
        username = objEList.Item(n).getAttribute("fullname")    
        emailaddress = objEList.Item(n).getAttribute("email") 
        exit for
    Next
    GetUserFullName = "<a href=""mailto:" & emailaddress & """ title=""Email " & username & """ class=""email-link"">" & username & "</a>" 
end function 

Function GetPageFilename()
    set XMLDom = Server.CreateObject("RDCMSAspObj.RDObject")
    set RQLObject = Server.CreateObject("RDCMSServer.XmlServer")     
    set objXMLDOM=Server.CreateObject("Microsoft.XMLDOM")
    
    xmlData = "<IODATA loginguid=""" & LoginGuid & """ sessionkey=""" & SessionKey & """><PAGE action=""load"" guid=""" & PageGuid & """ /></IODATA>" 
    xmlData = RQLObject.Execute(xmlData, sError) 

    objXMLDOM.LoadXml( xmlData ) 
    set objEList=objXMLDOM.getElementsByTagName("PAGE")
    For n = 0 to (objEList.Length-1)
        pagename = objEList.Item(n).getAttribute("name")   
        pageid = objEList.Item(n).getAttribute("id")
        exit for
    Next
    if pagename = "" then
        pagename = pageid
    end if 

    GetPageFilename = CheckFilenameExtension(pagename) 
End Function

Function CheckFilenameExtension(filename)
    thisFilename = filename
    if InStr(filename,".") = 0 then
        thisFilename = filename & DefaultExtension
    end if
    CheckFilenameExtension = thisFilename
End Function

Function FormatCreatedBy()
	Dim fullname
	Dim email
	Dim formatcreatedate

	set XMLDom = Server.CreateObject("RDCMSAspObj.RDObject")
    set RQLObject = Server.CreateObject("RDCMSServer.XmlServer")     
    set objXMLDOM=Server.CreateObject("Microsoft.XMLDOM")
    
    xmlData = "<IODATA loginguid=""" & LoginGuid & """><ADMINISTRATION><USER action=""load"" guid=""" & CreateBy & """ /></ADMINISTRATION></IODATA>" 
    xmlData = RQLObject.Execute(xmlData, sError)
	objXMLDOM.LoadXml( xmlData ) 
    set objEList=objXMLDOM.getElementsByTagName("USER")
    For n = 0 to (objEList.Length-1)
        fullname = objEList.Item(n).getAttribute("fullname")   
        email = objEList.Item(n).getAttribute("email")
        exit for
    Next

	formatcreatedate = CDate(DecodeDate(CreateDate))	

	FormatCreatedBy = "<a href=""mailto:" & email& """ title=""Email "&fullname&""" class=""email-link"">" & fullname & "</a> - " & Day(formatcreatedate) & " " & GetShortMonthName(Month(formatcreatedate)) & " " & Year(formatcreatedate)
End Function

Function FormatModifiedBy()
	Dim fullname
	Dim email
	Dim formatmodifydate

	set XMLDom = Server.CreateObject("RDCMSAspObj.RDObject")
    set RQLObject = Server.CreateObject("RDCMSServer.XmlServer")     
    set objXMLDOM=Server.CreateObject("Microsoft.XMLDOM")
    
    xmlData = "<IODATA loginguid=""" & LoginGuid & """><ADMINISTRATION><USER action=""load"" guid=""" & ModifiedBy & """ /></ADMINISTRATION></IODATA>" 
    xmlData = RQLObject.Execute(xmlData, sError)
	objXMLDOM.LoadXml( xmlData ) 
    set objEList=objXMLDOM.getElementsByTagName("USER")
    For n = 0 to (objEList.Length-1)
        fullname = objEList.Item(n).getAttribute("fullname")   
        email = objEList.Item(n).getAttribute("email")
        exit for
    Next

	formatmodifydate = CDate(DecodeDate(ModifiedDate))	

	FormatModifiedBy = "<a href=""mailto:" & email& """ title=""Email "&fullname&""" class=""email-link"">" & fullname & "</a> - " & Day(formatmodifydate) & " " & GetShortMonthName(Month(formatmodifydate)) & " " & Year(formatmodifydate)
End Function

Function GetShortMonthName(eventMonth)
    Dim monthName
    Select Case eventMonth 
        Case "01","1" 
            monthName = "Jan" 
        Case "02","2"  
            monthName = "Feb"
        Case "03","3" 
            monthName = "Mar"
        Case "04","4" 
            monthName = "Apr"  
        Case "05","5" 
            monthName = "May" 
        Case "06","6" 
            monthName = "Jun"
        Case "07","7" 
            monthName = "Jul" 
        Case "08","8" 
            monthName = "Aug"      
        Case "09","9" 
            monthName = "Sep" 
        Case "10" 
            monthName = "Oct"
        Case "11" 
            monthName = "Nov" 
        Case "12" 
            monthName = "Dec"   
    End Select
    GetShortMonthName = monthName
End function

Function DecodeDate(OldDate)
   Dim objIO 
   Set objIO = CreateObject("RDCMSAsp.RDPageData") 
   DecodeDate = objIO.decodedate(OldDate) 
End Function
%>
<root>
<status><![CDATA[<%=GetPageStatus()%>]]></status>
<filename><![CDATA[<%=GetPageFilename()%>]]></filename>
<createdby><![CDATA[<%=FormatCreatedBy()%>]]></createdby>
<modifiedby><![CDATA[<%=FormatModifiedBy()%>]]></modifiedby>
<pageid><![CDATA[<%=PageId%>]]></pageid>
</root>