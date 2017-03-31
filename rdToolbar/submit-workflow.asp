<%
' ------------------------------------------------------------------------------------------
'
' NAME: Submit page to workflow
' DESCRIPTION:  
' This plug-in enables a page to be easily submitted to workflow without having
' to open the page in SmartEdit mode.
'
' AUTHOR:  Kim Dezen (kim@kimdezen.com)
' VERSION: 1.0
' DATE: Feb 20 2012
' COMPATIBLE CMS VERSIONS:  10
'
' PLEASE NOTE:
' This script will not copy the option list entries for Option List
' elements and the default & sample text values for Text elements.
' These items will need to be managed manually.
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
' -------------------------------------------------------------------

Option Explicit
Server.ScriptTimeout = 500 
Response.Buffer = true
On Error Resume Next

Dim pageGuid
Dim workflowAction
Dim rqlString
Dim rqlResult

pageGuid = trim(request("pageGuid"))
workflowAction = 32768

if Session("LoginGuid") <> "" and Session("SessionKey") <> "" then
	' submit page to workflow
	rqlString = "<IODATA loginguid="""&Session("LoginGuid")&""" sessionkey=""" & Session("SessionKey") & """>" &_
				"<PAGE guid=""" & pageGuid & """ action=""save"" actionflag=""" & workflowAction & """/>"&_
	   	        "</IODATA>"
	rqlResult = SendXML(rqlString)
end if

function SendXML(xmlString) 
	dim objData
	dim sErrors
	dim xmlResult
	set objData = server.CreateObject("RDCMSASP.RdPageData") 
	objData.XMLServerClassname="RDCMSServer.XmlServer" 
	xmlResult = objData.ServerExecuteXML(xmlString, sErrors)
	if sErrors <> "" then
		errorFound = true
		errorMessage = sErrors
	end if
	sendXML = xmlResult
end function 

%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>Processing elements</title>

	<meta http-equiv="Content-Type" content="text/html; charset=iso-8" />
	<meta http-equiv="imagetoolbar" content="no" />

	<link rel="Stylesheet" type="text/css" HREF="/cms/stylesheets/ioStyleSheet.css">

</head>
<body class="bodyBGColorIoDialog" background="/cms/icons/back5.gif">

<table class="tdgrey" border="0" width="600" align="center" cellspacing="0" cellpadding="3">
<tr>
	<td width="100%">
		<table class="tdgreylight" border="0" width="100%" cellspacing="0" cellpadding="1">
		<tr>
			<td width="100%" vAlign="top" height="50">
			<table border="0" width="100%"><!-- Cell.Title -->
			<tr>
				<td class="titlebar" width="100%">Page submitted to Workflow</td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100%" vAlign="top">
				<blockquote>The page has been submitted into the workflow.</blockquote>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
</body>
</html>
