<%
' ------------------------------------------------------------------------------------------
'
' NAME: Connect To Mulitple Elements in Clipboard
' DESCRIPTION:  
' This plug-in enables users to duplicate multiple elements from one
' content class to another based on those that are currently selected 
' in the clipboard.
'
' AUTHOR:  Kim Dezen (kim@kimdezen.com)
' VERSION: 1.0
' DATE: July 4 2009
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

dim selectedTemplateGUID
dim clipboardItemsToProcess
dim clipboardSelectionElements
dim clipboardSelectionId
dim clipboardItem
dim clipboardSelectionType
dim currentElementDict
dim currentElement
dim optionElements
dim optionElementNode
dim rqlString
dim objXMLDOMElementData
dim anchorDict
dim areaDict
dim attributeDict
dim pageTypesDict
dim browseDict
dim backgroundDict
dim containerDict
dim databaseDict
dim frameDict
dim headlineDict
dim hitlistDict
dim imageDict
dim infoDict
dim ivwDict
dim listentryDict
dim listDict
dim mediaDict
dim optionlistDict
dim projectcontentDict
dim sitemapDict
dim standardfieldDict
dim textDict
dim transferDict
dim xcmsDict
dim liveserverconstraintDict
dim rqlResult
dim elementName
dim i
dim att
dim typeFound

if Session("LoginGuid") <> "" and Session("SessionKey") <> "" and Request.Form("actionFlag") = "1" then
	call PopulateDictionarys
	set objXMLDOMElementData=Server.CreateObject("Microsoft.XMLDOM")

	' contains all of the clipboard data to process
	clipboardItemsToProcess = Request.Form("clipboardItemsToProcess")
	
	' the content class/template we want to copy the element into
	selectedTemplateGUID = Session("TreeParentGuid")

	'convert clipboard items into array for processing
	clipboardItemsToProcess = Split(clipboardItemsToProcess, "*")

	for each clipboardItem in clipboardItemsToProcess
		typeFound = true
		clipboardSelectionElements = Split(clipboardItem, "_")

		'store id and type of current element
		clipboardSelectionType = clipboardSelectionElements(0)
		clipboardSelectionId = clipboardSelectionElements(1)
		
		' determine what type of element we need to process
		select case pageTypesDict.Item(clipboardSelectionType)
		case "anchor"
			set currentElementDict = anchorDict
		case "area"
			set currentElementDict = areaDict
		case "attribute"
			set currentElementDict = attributeDict
		case "background"
			set currentElementDict = backgroundDict
		case "browse"
			set currentElementDict = browseDict
		case "container"
			set currentElementDict = containerDict
		case "database"
			set currentElementDict = databaseDict
		case "frame"
			set currentElementDict = frameDict
		case "headline"
			set currentElementDict = headlineDict
		case "hitlist"
			set currentElementDict = hitlistDict
		case "image"
			set currentElementDict = imageDict
		case "info"
			set currentElementDict = infoDict
		case "ivw"
			set currentElementDict = ivwDict
		case "listentry"
			set currentElementDict = listentryDict
		case "list"
			set currentElementDict = listDict
		case "liveserverconstraint"
			set currentElementDict = liveserverconstraintDict
		case "media"
			set currentElementDict = mediaDict
		case "optionlist"
			set currentElementDict = optionlistDict
		case "projectcontent"
			set currentElementDict = projectcontentDict
		case "sitemap"
			set currentElementDict = sitemapDict
		case "standardfield"
			set currentElementDict = standardfieldDict
		case "text"
			set currentElementDict = textDict
		case "transfer"
			set currentElementDict = transferDict
		case "xcms"
			set currentElementDict = xcmsDict
		case else
			typeFound = false
		end select 

		if typeFound then
			' obtain details of the current element
			rqlString = "<IODATA loginguid="""&Session("LoginGuid")&""">" &_
						"<PROJECT sessionkey="""&Session("SessionKey")&""">" &_
						"<TEMPLATE><ELEMENT action=""load"" guid="""&clipboardSelectionId&""">" &_
						"<SELECTIONS action=""load""/></ELEMENT></TEMPLATE></PROJECT></IODATA>"
			rqlResult = SendXML(rqlString)
			objXMLDOMElementData.LoadXml(rqlResult)
			set currentElement=objXMLDOMElementData.getElementsByTagName("ELEMENT")
			
			' obtain the keys(or elements) in the dictionary for the current type
			' and build the rql to copy the element to new content class
			rqlString = "<IODATA loginguid="""&Session("LoginGuid")&""">" &_
						"<PROJECT sessionkey="""&Session("SessionKey")&""">" &_
						"<TEMPLATE guid="""&selectedTemplateGUID&""">" &_
						"<ELEMENT action=""save"" " 

			' exclude elements within the dictionary, but include the rest
			for each att in currentElement.Item(0).attributes
				if not currentElementDict.Exists(att.name) then
					rqlString = rqlString & att.name & "=""" & Server.HTMLEncode(att.text) &""" "
				end if
			next
			
			' close the rql query and execute
			rqlString = rqlString & "></ELEMENT></TEMPLATE></PROJECT></IODATA>"
			rqlResult = SendXML(rqlString)
		end if
	next
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

sub PopulateDictionarys
	' create new dictionary instances
	set anchorDict = CreateObject("Scripting.Dictionary") 
	set pageTypesDict = CreateObject("Scripting.Dictionary") 
	set areaDict = CreateObject("Scripting.Dictionary") 
	set attributeDict = CreateObject("Scripting.Dictionary") 
	set backgroundDict = CreateObject("Scripting.Dictionary") 
	set browseDict = CreateObject("Scripting.Dictionary") 
	set containerDict = CreateObject("Scripting.Dictionary") 
	set databaseDict = CreateObject("Scripting.Dictionary") 
	set frameDict = CreateObject("Scripting.Dictionary") 
	set headlineDict = CreateObject("Scripting.Dictionary") 
	set hitlistDict = CreateObject("Scripting.Dictionary")
	set imageDict = CreateObject("Scripting.Dictionary")
	set infoDict = CreateObject("Scripting.Dictionary")
	set ivwDict = CreateObject("Scripting.Dictionary")
	set listentryDict = CreateObject("Scripting.Dictionary")
	set listDict = CreateObject("Scripting.Dictionary")
	set mediaDict = CreateObject("Scripting.Dictionary")
	set optionlistDict = CreateObject("Scripting.Dictionary")
	set projectcontentDict = CreateObject("Scripting.Dictionary")
	set sitemapDict = CreateObject("Scripting.Dictionary")
	set standardfieldDict = CreateObject("Scripting.Dictionary") 
	set textDict = CreateObject("Scripting.Dictionary") 
	set transferDict = CreateObject("Scripting.Dictionary") 
	set liveserverconstraintDict = CreateObject("Scripting.Dictionary") 
	set xcmsDict = CreateObject("Scripting.Dictionary") 
	
	' page types
	pageTypesDict.add "project.4141", "anchor"
	pageTypesDict.add "project.4142", "area"
	pageTypesDict.add "project.4165", "attribute"
	pageTypesDict.add "project.4143", "background"
	pageTypesDict.add "project.4144", "browse"
	pageTypesDict.add "project.4145", "container"
	pageTypesDict.add "project.4146", "database"
	pageTypesDict.add "project.4147", "frame"
	pageTypesDict.add "project.4159", "headline"
	pageTypesDict.add "project.4158", "hitlist"
	pageTypesDict.add "project.4148", "image"
	pageTypesDict.add "project.4149", "info"
	pageTypesDict.add "project.4150", "ivw"
	pageTypesDict.add "project.4152", "listentry"
	pageTypesDict.add "project.4151", "list"
	pageTypesDict.add "project.TE24", "liveserverconstraint"
	pageTypesDict.add "project.4154", "media"
	pageTypesDict.add "project.4155", "optionlist"
	pageTypesDict.add "project.4162", "projectcontent"
	pageTypesDict.add "project.4168", "sitemap"
	pageTypesDict.add "project.4156", "standardfield"
	pageTypesDict.add "project.4157", "text"
	pageTypesDict.add "project.4160", "transfer"
	pageTypesDict.add "project.4169", "xcms"

	'*******************************************************
	' you can customise the following dictionaries to exclude
	' the element properties that are automatically applied  
	' when copying elements from one content class to another
	'*******************************************************
	
	' anchor elements
	anchorDict.add "action", "" 
	anchorDict.add "parentguid", "" 
	anchorDict.add "templateguid", ""
	anchorDict.add "languagevariantid", ""
	anchorDict.add "dialoglanguageid", ""
	anchorDict.add "guid", ""
	
	' area elements
	areaDict.add "action", "" 
	areaDict.add "parentguid", "" 
	areaDict.add "templateguid", ""
	areaDict.add "languagevariantid", ""
	areaDict.add "dialoglanguageid", ""
	areaDict.add "guid", ""
	
	' attribute elements
	attributeDict.add "action", "" 
	attributeDict.add "parentguid", "" 
	attributeDict.add "templateguid", ""
	attributeDict.add "languagevariantid", ""
	attributeDict.add "dialoglanguageid", ""
	attributeDict.add "guid", ""
	
	' background elements
	backgroundDict.add "action", "" 
	backgroundDict.add "parentguid", "" 
	backgroundDict.add "templateguid", ""
	backgroundDict.add "languagevariantid", ""
	backgroundDict.add "dialoglanguageid", ""
	backgroundDict.add "guid", ""
	
	' browse elements
	browseDict.add "action", "" 
	browseDict.add "parentguid", "" 
	browseDict.add "templateguid", ""
	browseDict.add "languagevariantid", ""
	browseDict.add "dialoglanguageid", ""
	browseDict.add "guid", ""
	
	' container elements
	containerDict.add "action", "" 
	containerDict.add "parentguid", "" 
	containerDict.add "templateguid", ""
	containerDict.add "languagevariantid", ""
	containerDict.add "dialoglanguageid", ""
	containerDict.add "guid", ""
	
	' database elements
	databaseDict.add "action", "" 
	databaseDict.add "parentguid", "" 
	databaseDict.add "templateguid", ""
	databaseDict.add "languagevariantid", ""
	databaseDict.add "dialoglanguageid", ""
	databaseDict.add "guid", ""
	
	' frame elements
	frameDict.add "action", "" 
	frameDict.add "parentguid", "" 
	frameDict.add "templateguid", ""
	frameDict.add "languagevariantid", ""
	frameDict.add "dialoglanguageid", ""
	frameDict.add "guid", ""
	
	' headline elements
	headlineDict.add "action", "" 
	headlineDict.add "parentguid", "" 
	headlineDict.add "templateguid", ""
	headlineDict.add "languagevariantid", ""
	headlineDict.add "dialoglanguageid", ""
	headlineDict.add "guid", ""
	
	' hitlist elements
	hitlistDict.add "action", "" 
	hitlistDict.add "parentguid", "" 
	hitlistDict.add "templateguid", ""
	hitlistDict.add "languagevariantid", ""
	hitlistDict.add "dialoglanguageid", ""
	hitlistDict.add "guid", ""
	
	' image elements
	imageDict.add "action", "" 
	imageDict.add "parentguid", "" 
	imageDict.add "templateguid", ""
	imageDict.add "languagevariantid", ""
	imageDict.add "dialoglanguageid", ""
	imageDict.add "guid", ""
	
	' info elements
	infoDict.add "action", "" 
	infoDict.add "parentguid", "" 
	infoDict.add "templateguid", ""
	infoDict.add "languagevariantid", ""
	infoDict.add "dialoglanguageid", ""
	infoDict.add "guid", ""
	
	' ivw elements
	ivwDict.add "action", "" 
	ivwDict.add "parentguid", "" 
	ivwDict.add "templateguid", ""
	ivwDict.add "languagevariantid", ""
	ivwDict.add "dialoglanguageid", ""
	ivwDict.add "guid", ""
	
	' list entry elements
	listentryDict.add "action", "" 
	listentryDict.add "parentguid", "" 
	listentryDict.add "templateguid", ""
	listentryDict.add "languagevariantid", ""
	listentryDict.add "dialoglanguageid", ""
	listentryDict.add "guid", ""
	
	' list elements
	listDict.add "action", "" 
	listDict.add "parentguid", "" 
	listDict.add "templateguid", ""
	listDict.add "languagevariantid", ""
	listDict.add "dialoglanguageid", ""
	listDict.add "guid", ""
	
	' liveserver constraint elements
	liveserverconstraintDict.add "action", "" 
	liveserverconstraintDict.add "parentguid", "" 
	liveserverconstraintDict.add "templateguid", ""
	liveserverconstraintDict.add "languagevariantid", ""
	liveserverconstraintDict.add "dialoglanguageid", ""
	liveserverconstraintDict.add "guid", ""
	
	' media elements
	mediaDict.add "action", "" 
	mediaDict.add "parentguid", "" 
	mediaDict.add "templateguid", ""
	mediaDict.add "languagevariantid", ""
	mediaDict.add "dialoglanguageid", ""
	mediaDict.add "guid", ""
	
	' option list elements
	optionlistDict.add "action", "" 
	optionlistDict.add "parentguid", "" 
	optionlistDict.add "templateguid", ""
	optionlistDict.add "languagevariantid", ""
	optionlistDict.add "dialoglanguageid", ""
	optionlistDict.add "guid", ""
	
	' project content elements
	projectcontentDict.add "action", "" 
	projectcontentDict.add "parentguid", "" 
	projectcontentDict.add "templateguid", ""
	projectcontentDict.add "languagevariantid", ""
	projectcontentDict.add "dialoglanguageid", ""
	projectcontentDict.add "guid", ""
	
	' sitemap elements
	sitemapDict.add "action", "" 
	sitemapDict.add "parentguid", "" 
	sitemapDict.add "templateguid", ""
	sitemapDict.add "languagevariantid", ""
	sitemapDict.add "dialoglanguageid", ""
	sitemapDict.add "guid", ""
	
	' standard field elements
	standardfieldDict.add "action", "" 
	standardfieldDict.add "parentguid", "" 
	standardfieldDict.add "templateguid", ""
	standardfieldDict.add "languagevariantid", ""
	standardfieldDict.add "dialoglanguageid", ""
	standardfieldDict.add "guid", ""
	
	' text elements
	textDict.add "action", "" 
	textDict.add "parentguid", "" 
	textDict.add "templateguid", ""
	textDict.add "languagevariantid", ""
	textDict.add "dialoglanguageid", ""
	textDict.add "guid", ""
	
	' transfer elements
	transferDict.add "action", "" 
	transferDict.add "parentguid", "" 
	transferDict.add "templateguid", ""
	transferDict.add "languagevariantid", ""
	transferDict.add "dialoglanguageid", ""
	transferDict.add "guid", ""
	
	' xcms elements
	xcmsDict.add "action", "" 
	xcmsDict.add "parentguid", "" 
	xcmsDict.add "templateguid", ""
	xcmsDict.add "languagevariantid", ""
	xcmsDict.add "dialoglanguageid", ""
	xcmsDict.add "guid", ""
	
end sub

if Request.Form("actionFlag") = "1" then 
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>Connect to Multiple Elements in Clipboard</title>

	<meta http-equiv="Content-Type" content="text/html; charset=iso-8" />
	<meta http-equiv="imagetoolbar" content="no" />
	
	<link rel="Stylesheet" type="text/css" href="rdUIConnectMultipleElements.css" />
	<script src="jquery-1.2.6.pack.js" type="text/javascript"></script>

	<script  type="text/javascript" language="javascript">
	function ClosePage()
	{
		window.close();
	}
	</script>
</head>
<body>

<div class="page">
	<div class="main">
		<h1>Connect to Multiple Elements in Clipboard</h1>
		<div class="content">
			<div class="busy"><img src="/cms/icons/PageFlag0.gif" width="16" height="16" alt="success" /><span class="busy_text">All elements copied</span></div>
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
	
	function closeWindowSubmit() {
		window.close();
	}
//]]>  
</script>
</body>
</html>

<% else %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>Connect to Multiple Elements in Clipboard</title>

	<meta http-equiv="Content-Type" content="text/html; charset=iso-8" />
	<meta http-equiv="imagetoolbar" content="no" />

	<link rel="Stylesheet" type="text/css" href="rdUIConnectMultipleElements.css" />
	<script src="jquery-1.2.6.pack.js" type="text/javascript"></script>
	
	<script  type="text/javascript" language="javascript">
	function ProcessClipboard()
	{

		// obtain all of the items in the clipboard
		var clipboardItems = top.opener.parent.frames.ioClipboard.document.all.tags("INPUT");
		var clipboardItemsToProcess = "";

		// process all items
		for (var i = 1; i < clipboardItems.length; i++) 
		{
			if (clipboardItems(i).type == "checkbox") 
			{
				sName = clipboardItems(i).name;
				objParent = clipboardItems(i).parentElement.parentElement;

				if(clipboardItems(i).checked)
				{
					if(clipboardItemsToProcess != "")
					{
						clipboardItemsToProcess =  clipboardItemsToProcess + "*"
					}
					clipboardItemsToProcess =  clipboardItemsToProcess + objParent.elttype + "_" + objParent.id;
				}
				
        		}
		}
		document.multipleElememts.clipboardItemsToProcess.value = clipboardItemsToProcess;
		document.multipleElememts.submit();
	}
	</script>
</head>
<body onload="ProcessClipboard();">

<form name="multipleElememts" action="rdUIConnectMultipleElements.asp" method="POST">
	<input type="hidden" name="clipboardItemsToProcess" value="">
	<input type="hidden" name="actionFlag" value="1">
</form>

<div class="page">
	<div class="main">
		<h1>Connect to Multiple Elements in Clipboard</h1>
		<div class="content">
			<div class="busy"><img src="img_bgbusy.gif" height="15" width="15" alt="busy icon" /><span class="busy_text">Processing elements...</span></div>
		</div>
	</div>
	<div class="buttons clearfix">
		<div class="logo"><a href="http://www.reddotcmsblog.com/author/kimdezen" target="_blank"><img src="img_reddotcmsblog_kimdezen.gif" height="31" width="130" alt="reddotcmsblog.com - Plugin By: Kim Dezen" border="0" /></a></div>
		<div class="commandButton cancelButton"><img class="commandButtonIcon" id="cancelIcon" src="/cms/Icons/CommandButtons/Cancel.gif" alt="Cancel" />&nbsp;Cancel</div>
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
		$(".cancelButton").click(function () {
			window.close();
		});
	});  
	
	function closeWindowSubmit() {
		window.close();
	}
//]]>  
</script>
</body>
</html>

<% end if %>