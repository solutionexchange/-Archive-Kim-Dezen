<%
' ------------------------------------------------------------------------------------------
'
' NAME: Delete Unlinked Pages
' DESCRIPTION:  
' This plug-in will automatically delete pages under Unlinked Pages.
' ONLY those pages that are currently shown in the tree will be deleted,
' therefore any display filter settings that have been selected will 
' affect the amount of pages that are deleted.
'
' AUTHOR:  Kim Dezen (kim@kimdezen.com)
' VERSION: 1.0
' DATE: July 19 2009
' COMPATIBLE CMS VERSIONS:  7.5
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
' Attribution-Noncommercial-Share Alike 3.0 Unported
' http://creativecommons.org/licenses/by-nc-sa/3.0/
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

Dim i
Dim rqlString
Dim objXMLDOMElementData
Dim unlinkedPages
Dim rqlResult

' get a list of all pages currently listed under unlinked pages within SmartTree
rqlString = "<IODATA loginguid="""& Session("LoginGuid")&""" sessionkey="""& Session("SessionKey") &""">" &_
			"<TREESEGMENT type=""app.1805"" action=""load"" guid="""& Session("TreeGuid") &""" descent=""app"" parentguid="""& Session("TreeParentGuid") &"""/></IODATA>"
rqlResult = SendXML(rqlString)

' loop through all pages and build up RQL query to delete these pages
' and insert them into the recycle bin.
set objXMLDOMElementData=Server.CreateObject("Microsoft.XMLDOM")
objXMLDOMElementData.LoadXml(rqlResult)
set unlinkedPages = objXMLDOMElementData.getElementsByTagName("SEGMENT")
rqlString = "<IODATA><PROJECT sessionkey="""& Session("SessionKey") &"""><PAGES action=""deletefreepages"">"
For i = 0 to (unlinkedPages.Length-1)    
	rqlString = rqlString & "<ENTRY guid="""& unlinkedPages.Item(i).getAttribute("guid") &""" type=""page"" descent=""unknown"" />"
Next
rqlString = rqlString & "<LANGUAGEVARIANTS><LANGUAGEVARIANT language="""& Session("LanguageId") &"""/></LANGUAGEVARIANTS></PAGES></PROJECT></IODATA>"
response.write rqlString
rqlResult = SendXML(rqlString)

function SendXML(xmlString) 
	dim objData
	dim sErrors
	dim xmlResult
	dim errorFound
	dim errorMessage
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
	<title>Delete Unlinked Pages</title>

	<meta http-equiv="Content-Type" content="text/html; charset=iso-8" />
	<meta http-equiv="imagetoolbar" content="no" />
	
	<link rel="Stylesheet" type="text/css" href="rdUIDeleteUnlinkedPages.css" />
	<script src="jquery-1.2.6.pack.js" type="text/javascript"></script>

</head>
<body onunload="top.opener.parent.frames.ioTree.ReloadTreeSegment();">

<div class="page">
	<div class="main">
		<h1>Delete Unlinked Pages</h1>
		<div class="content">
			<div class="busy"><img src="/cms/icons/PageFlag0.gif" width="16" height="16" alt="success" /><span class="busy_text">All shown unlinked pages have been deleted<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;and have been moved to the recycle bin.</span></div>
		</div>
	</div>
	<div class="buttons clearfix">
		<div class="logo"><a href="http://www.reddotcmsblog.com/author/kimdezen" target="_blank"><img src="img_reddotcmsblog_kimdezen.gif" height="31" width="130" alt="reddotcmsblog.com - Plugin By: Kim Dezen" border="0" /></a></div>
		<div class="commandButton okButton"><img class="commandButtonIcon" id="okIcon" src="/cms/Icons/CommandButtons/Ok.gif" alt="Close Page" />&nbsp;Close Page</div>
	</div>
</div>

<script type="text/javascript">
//<![CDATA[
	$(document).ready(function(){
		// hover styles for buttons
		$(".okButton, .cancelButton").hover(
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