<%
' ------------------------------------------------------------------------------------------
'
' NAME: Copy Element
' DESCRIPTION:  
' This plug-in will make a copy of the selected elemement in SmartTree mode
'
' AUTHOR:  Kim Dezen (kim@kimdezen.com)
' VERSION: 1.0
' DATE: May 4 2011
' COMPATIBLE CMS VERSIONS:  10.0+
'
' -----------------------------DISCLAIMER-----------------------------
' This script is not an official component of the RedDot Content 
' Management Server, and is not supported or guaranteed by RedDot 
' Solutions.  All claims of functionality are made by the 
' scripts author, and are not guaranteed by RedDot (regardless of any 
' affiliation the author might have with RedDot Solutions). 
'
' This software is licensed under a Creative Commons
' Attribution-Share Alike 3.0 License. Some rights reserved.
' http://creativecommons.org/licenses/by-sa/3.0/
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
' EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
' OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
' NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
' HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
' OTHER DEALINGS IN THE SOFTWARE.
' -----------------------------------------------------------------------

Server.ScriptTimeout = 500 
Response.Buffer = true
On Error Resume Next

Dim errorFound 
Dim errorMessage
Dim sessionKey
Dim loginGUID
Dim sourceElementGUID
Dim sourceContentClassGUID
Dim xmlResult
Dim xmlString
Dim elementInfo
Dim elementAttributes
Dim elementProperties
Dim elementName
Dim objElementsList
Dim objItemsList
Dim elementsDict
Dim i
Dim n
Dim isfound
Dim newName
Dim elementType
Dim elementTypeName
Dim sampleText
Dim defaultText
Dim defaultTextGUID
Dim defaultOptionList
Dim optionListItems
Dim optionListCount

' intialise variables
errorFound = false
errorMessage = ""
isfound = true
newName = ""
sessionKey = Session("SessionKey")
loginGUID = Session("LoginGuid")
sourceElementGUID = Session("TreeGuid")
sourceContentClassGUID = ""
elementType = ""
sampleText = ""
defaultText = ""
defaultTextGUID = ""
sampleTextGUID = ""
defaultOptionList = ""
optionListItems = ""
optionListCount = 0
elementTypeName = ""

set elementsDict = CreateObject("Scripting.Dictionary")

if Session("LoginGuid") <> "" <> "" and Session("SessionKey") <> ""  then
	set objXMLDOMAssigned=Server.CreateObject("Microsoft.XMLDOM")
	
	' obtain details of selected element
	xmlString = "<IODATA loginguid=""" & loginGUID & """><PROJECT sessionkey="""& sessionKey &"""><TEMPLATE><ELEMENT action=""load"" guid="""& sourceElementGUID &""" /></TEMPLATE></PROJECT></IODATA>"	
	xmlString = sendXML(xmlString)
	objXMLDOMAssigned.LoadXml(xmlString) 
	
	Set elementInfo = objXMLDOMAssigned.documentElement.selectSingleNode("TEMPLATE/ELEMENT") 
	For Each elementAttributes in elementInfo.attributes
		' build up a string containing all element properties (ignore: GUID, name, action and type)
		if elementAttributes.name = "eltname" then
			elementName = elementAttributes.value
		elseif elementAttributes.name = "templateguid" then
			sourceContentClassGUID = elementAttributes.value
		elseif elementAttributes.name <> "guid" and elementAttributes.name <> "action" and elementAttributes.name <> "eltrdexampleguid" and elementAttributes.name <> "eltdefaulttextguid" then
			' determine what type of element to process
			if elementAttributes.name = "elttype" then	
				elementType = elementAttributes.value
			end if
			' dont include dafaultvalue when processing option lists
			if elementType = "8" and elementAttributes.name = "eltdefaultvalue" then
				defaultOptionList = elementAttributes.value
			else
				elementProperties = elementProperties & elementAttributes.name & "=""" & cleanXMLValue(elementAttributes.value) & """ "
			end if
		end if
	Next
	
	' determine if there already is an element of the same name
	xmlString = "<IODATA loginguid=""" & loginGUID & """ sessionkey="""& sessionKey &"""><TEMPLATE guid=""" & sourceContentClassGUID & """><ELEMENTS action=""list"" /></TEMPLATE></IODATA>"	
	xmlResult = sendXML(xmlString)
	objXMLDOMAssigned.LoadXml(xmlResult) 
		
	' store all element names in a dictionary, so we can check for duplication element name
	set objElementsList=objXMLDOMAssigned.getElementsByTagName("ELEMENT")
	For i = 0 to (objElementsList.Length-1)    
		elementsDict.Add objElementsList.Item(i).getAttribute("name"), objElementsList.Item(i).getAttribute("guid")
	Next
	Do until isFound = false
		newName = newName & "Copy_of_" & elementName
		If not elementsDict.Exists(newName) Then 
			isFound = false
		End if
	Loop
	
	' assign default and sample text if we are dealing with a text elements
	if elementType = "31" or elementType = "32" then
		'create new default text 
		xmlString = "<IODATA loginguid=""" & loginGUID & """ sessionkey="""& sessionKey &"""  format=""1""><PROJECT><TEXT action=""load"" guid=""" & sourceElementGUID & """ texttype=""3"" /></PROJECT></IODATA>"	
		defaultText = sendXML(xmlString)
		
		xmlString = "<IODATA loginguid=""" & loginGUID & """ sessionkey="""& sessionKey &"""  format=""1""><PROJECT><TEXT action=""save"" texttype=""3"" guid="""">" & cleanXMLValue(defaultText) & "</TEXT></PROJECT></IODATA>"
		xmlString = sendXML(xmlString)
		objXMLDOMAssigned.LoadXml(xmlString) 
		defaultTextGUID = Left(Trim(objXMLDOMAssigned.documentElement.firstChild.nodeValue), 32)
		
		elementProperties = elementProperties & "eltdefaulttextguid=""" & defaultTextGUID & """ "

		'create new sample text
		xmlString = "<IODATA loginguid=""" & loginGUID & """ sessionkey="""& sessionKey &"""  format=""1""><PROJECT><TEXT action=""load"" guid=""" & sourceElementGUID & """ texttype=""10"" /></PROJECT></IODATA>"	
		sampleText = sendXML(xmlString)
		
		xmlString = "<IODATA loginguid=""" & loginGUID & """ sessionkey="""& sessionKey &"""  format=""1""><PROJECT><TEXT action=""save"" texttype=""3"" guid="""">" & cleanXMLValue(sampleText) & "</TEXT></PROJECT></IODATA>"
		xmlString = sendXML(xmlString)
		objXMLDOMAssigned.LoadXml(xmlString) 
		sampleTextGUID = Left(Trim(objXMLDOMAssigned.documentElement.firstChild.nodeValue), 32)
		
		elementProperties = elementProperties & "eltrdexampleguid=""" & sampleTextGUID & """ "
	end if 
	
	if elementType = "8" then
		' obtain list of all option list values
		xmlString = "<IODATA loginguid=""" & loginGUID & """><PROJECT sessionkey="""& sessionKey &""" ><TEMPLATE><ELEMENT action=""load"" guid=""" & sourceElementGUID & """><SELECTIONS action=""load"" guid=""" & sourceElementGUID & """/></ELEMENT></TEMPLATE></PROJECT></IODATA>"	
		xmlString = sendXML(xmlString)
		objXMLDOMAssigned.LoadXml(xmlString) 
	
		set objElementsList=objXMLDOMAssigned.getElementsByTagName("SELECTION")
		For i = 0 to (objElementsList.Length-1) 
			if objElementsList.Item(i).getAttribute("guid") = defaultOptionList then
				optionListItems = optionListItems & " &lt;SELECTION identifier=&quot;NaN&quot; "
				
				set objItemsList= objElementsList(i).getElementsByTagName("ITEM")
				For n = 0 to (objItemsList.Length-1) 
					if n = 0 then
						optionListItems = optionListItems & "description=&quot;"&objItemsList.Item(n).getAttribute("name")&"&quot; value=&quot;"&objItemsList.Item(n).childNodes(0).nodeValue&"&quot;&gt;"
					end if
				
					optionListItems = optionListItems & " &lt;ITEM languageid=&quot;"&objItemsList.Item(n).getAttribute("languageid")&"&quot; name=&quot;" & objItemsList.Item(n).getAttribute("name") & "&quot;&gt;" & objItemsList.Item(n).childNodes(0).nodeValue & "&lt;/ITEM&gt;"
				Next
				
				optionListItems = optionListItems & "&lt;/SELECTION&gt;"
				
			else
				optionListCount = optionListCount + 1
				optionListItems = optionListItems & " &lt;SELECTION identifier=&quot;"&optionListCount&"&quot; "
				
				set objItemsList=objElementsList(i).getElementsByTagName("ITEM")
				For n = 0 to (objItemsList.Length-1) 
					if n = 0 then
						optionListItems = optionListItems & "description=&quot;"&objItemsList.Item(n).getAttribute("name")&"&quot; value=&quot;"&objItemsList.Item(n).childNodes(0).nodeValue&"&quot;&gt;"
					end if
				
				
					optionListItems = optionListItems & " &lt;ITEM languageid=&quot;"&objItemsList.Item(n).getAttribute("languageid")&"&quot; name=&quot;" & objItemsList.Item(n).getAttribute("name") & "&quot;&gt;" & objItemsList.Item(n).childNodes(0).nodeValue & "&lt;/ITEM&gt;"
				Next
				
				optionListItems = optionListItems & "&lt;/SELECTION&gt;"
				
			end if
		Next
		
		optionListItems = "&lt;SELECTIONS&gt;" & optionListItems & "&lt;/SELECTIONS&gt;"
		elementProperties = elementProperties & " eltoptionlistdata=""" & optionListItems & """ "
		
		if defaultOptionList <> "" then
			elementProperties = elementProperties & " eltdefaultvalue=""NaN"" "
		else
			elementProperties = elementProperties & " eltdefaultvalue=""#"&sessionKey&""" "
		end if 
	end if
	
	' create new element now that we have all the details
	xmlString = "<IODATA loginguid=""" & loginGUID & """><PROJECT sessionkey="""& sessionKey &"""><TEMPLATE guid=""" & sourceContentClassGUID & """><ELEMENT action=""save"" eltname="""&newName&""" " & elementProperties & " /></TEMPLATE></PROJECT></IODATA>"	
	xmlResult = sendXML(xmlString)
	
	' determine element type name
	Select Case elementType
	Case "1"
		elementTypeName = "Standard Field (Text)"
	Case "2"
		elementTypeName = "Image"
	Case "3"
		elementTypeName = "Frame"
	Case "5"
		elementTypeName = "Standard Field (Date)"
	Case "8"
		elementTypeName = "Option List"
	Case "12"
		elementTypeName = "Headline"
	Case "14"
		elementTypeName = "Database Content"
	Case "15"
		elementTypeName = "Area"
	Case "19"
		elementTypeName = "Background"
	Case "24"
		elementTypeName = "Hit List"
	Case "25"
		elementTypeName = "List Entry"
	Case "26"
		elementTypeName = "Anchor as text"
	Case "27"
		elementTypeName = "Anchor as image"
	Case "28"
		elementTypeName = "Container"
	Case "31"
		elementTypeName = "Text (ASCI)"
	Case "32"
		elementTypeName = "Text (HTML)"
	Case "38"
		elementTypeName = "Media"
	Case "39"
		elementTypeName = "Standard Field (Time)"
	Case "48"
		elementTypeName = "Standard Field (Numeric)"
	Case "60"
		elementTypeName = "Transfer"
	Case "1002"
		elementTypeName = "Info"
	Case "2627"
		elementTypeName = "Anchor not yet defined as text or image"
	Case "1000"
		elementTypeName = "Standard Field"
	End Select

else
	errorFound = true
	errorMessage = "Please logon to CMS before accessing this plugin."
end if

function sendXML(xmlString) 
	Dim objData
	Dim sErrors
	Dim xmlResult
	Set objData = server.CreateObject("RDCMSASP.RdPageData") 
	objData.XMLServerClassname="RDCMSServer.XmlServer" 
	xmlResult = objData.ServerExecuteXML(xmlString, sErrors)
	if sErrors <> "" then
		errorFound = true
		errorMessage = sErrors
	end if
	sendXML = xmlResult
end function 

function cleanXMLValue(xmlValue)
	cleanXMLValue = Server.HTMLEncode(xmlValue)
end function
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title><%=Session("CmsWindowTitle")%> - Copy Element</title>
	
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="imagetoolbar" content="no" />
	
	<!-- stylesheets references -->
	<link rel="Stylesheet" type="text/css" href="rdUICopyElement.css" />
	
	<!-- javascript references -->
	<script src="jquery-1.2.6.pack.js" type="text/javascript"></script>
</head>
<body>

<div class="page">
		<div class="page-inner">
		<div class="main">
			<div class="title"><h1>Copy Element</h1></div>

			<% if errorFound then %>
			<div class="error">
				<p><strong>Sorry, this page could not be displayed properly.</strong></p>
				<p><%=errorMessage%></p>
			</div>
			<% else %>
			<div class="content">
				<div class="success">
					<p>'<%=elementName%>' successfully copied:<br /></p><p><br /><strong><%=newName%></strong> - <%=elementTypeName%></p>
				</div>
			</div>
			<% end if %>
		</div>
		</div>
		<div class="buttons clearfix">
			<div class="plugin-by"><a href="http://www.kimdezen.com" target="_blank"><img src="http://www.kimdezen.com/wp-content/uploads/2011/05/plugin-by-kim-dezen.gif" width="58" height="23" alt="Plugin By: Kim Dezen" border="0" /></a></div>
			<div class="commandButton okButton"><img class="commandButtonIcon" id="okIcon" src="/cms/Icons/CommandButtons/Ok.gif" alt="OK" />&nbsp;OK</div>&nbsp;
		</div>
</div>
<script type="text/javascript">
//<![CDATA[
	$(document).ready(function(){
		// hover styles for buttons
		$(".okButton").hover(
		function () {
			$(this).addClass("activeCommandButton");
		}, 
		function () {
			$(this).removeClass("activeCommandButton");
		});
	
		// button actions
		$(".okButton").click(function () {
			window.close();
		});
	});  
//]]>  
</script>
</body>
</html>