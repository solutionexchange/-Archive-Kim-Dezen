<%
' ------------------------------------------------------------------------------------------
'
' NAME: Checkbox Assign Keywords
' DESCRIPTION:  
' This plug-in enables users to assign keywords to a page or link by 
' selecting the corresponding checkboxes. 
'
' AUTHOR:  Kim Dezen (kim@kimdezen.com)
' VERSION: 1.1
' DATE: March 11 2011
' COMPATIBLE CMS VERSIONS:  10.0+
'
'
' INSTRUCTIONS:
' By default, users can select keywords within any category - however if
' required you can limit which category keywords can be selected by passing
' in the name of the categories you want to display via the "shownCategories"
' querystring variable. Separate multiple categories using the "~" delimiter
' e.g. ?shownCategories=Category 1~Category 2
'
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

Dim errorFound 
Dim errorMessage
Dim objXMLDOMAssigned
Dim xmlString
Dim postBack
Dim xmlResult
Dim objPageKeywordsList
Dim objXMLDOMCategories
Dim objXMLDOMKeywords
Dim i
Dim n
Dim allKeywords
Dim previousKeywords
Dim previousKeywordsDict
Dim objCategoryList
Dim objKeywordList
Dim categoryName
Dim categoryGUID
Dim keywordName
Dim keywordGUID
Dim isChecked
Dim keywordsDict
Dim rowCount
Dim shownCategories
Dim showAllCategories
Dim shownCategoriesDict
Dim hideCategory

set keywordsDict = CreateObject("Scripting.Dictionary")
set shownCategoriesDict = CreateObject("Scripting.Dictionary")
set previousKeywordsDict = CreateObject("Scripting.Dictionary")

' intialise variables
errorFound = false
showAllCategories = true
hideCategory = false
errorMessage = ""
postBack = trim(request("postBack"))
shownCategories = trim(request("shownCategories"))
allKeywords = trim(request("allKeywords"))
previousKeywords = trim(request("previousKeywords"))

if Session("LoginGuid") <> "" <> "" and Session("SessionKey") <> ""  then
	set objXMLDOMAssigned=Server.CreateObject("Microsoft.XMLDOM")

	if postBack <> "1" then
		' page is displayed for the first time, so we need to determine which keywords (if any) are assigned to the current page/link
		' the PageID session variable is not empty if we are currently applying keywords to a page
		if Session("PageId") <> "" then
			' page
			xmlString = "<IODATA loginguid=""" & Session("LoginGuid") & """ sessionkey=""" & Session("SessionKey") & """><PROJECT sessionkey="""& Session("SessionKey") &"""><PAGE guid="""& Session("PageGuid") &"""><KEYWORDS action=""load""/></PAGE></PROJECT></IODATA>"
		else
			' link
			xmlString = "<IODATA loginguid=""" & Session("LoginGuid") & """ sessionkey=""" & Session("SessionKey") & """><PROJECT sessionkey="""& Session("SessionKey") &"""><LINK guid="""&Session("TreeGuid")&"""><KEYWORDS action=""load""/></LINK></PROJECT></IODATA>"
		end if
		xmlResult = sendXML(xmlString)
		objXMLDOMAssigned.LoadXml(xmlResult) 
		
		' store these keywords in a dictionary so they can be accessed later on when we display the checkboxes on the page
		set objPageKeywordsList=objXMLDOMAssigned.getElementsByTagName("KEYWORD")
		For i = 0 to (objPageKeywordsList.Length-1)    
			keywordsDict.Add objPageKeywordsList.Item(i).getAttribute("guid"), objPageKeywordsList.Item(i).getAttribute("guid")
		Next
		
		' obtain a list of omitted categories and store these in a dictionary for access later
		if shownCategories <> "" then
			'categories to display are delimited using the  ~ character
			shownCategories = Split(shownCategories, "~")
			For i = 0 To UBound(shownCategories)
				if shownCategories(i) <> "" then
					shownCategoriesDict.Add shownCategories(i), shownCategories(i) 
				end if
			Next
		end if
	else
		'
		' form has been submitted, so we need to save the keyword selection for the page/link
		'
		'get all previously selected keywords and store the guids of the keywords in a dictionary
		if previousKeywords <> "" then
			previousKeywords = Split(previousKeywords, ",")
			For i = 0 To UBound(previousKeywords)-1
				previousKeywordsDict.Add previousKeywords(i), previousKeywords(i) 
			Next
		end if
	
		if Session("PageId") <> "" then
			' page
			xmlString = "<IODATA loginguid=""" & Session("LoginGuid") & """ sessionkey=""" & Session("SessionKey") & """><PROJECT sessionkey=""" & Session("SessionKey") & """><PAGE guid="""&Session("pageguid")&""" action=""assign""><KEYWORDS>"
		else
			' link
			xmlString = "<IODATA loginguid=""" & Session("LoginGuid") & """ sessionkey=""" & Session("SessionKey") & """><PROJECT sessionkey=""" & Session("SessionKey") & """><LINK guid="""&Session("TreeGuid")&""" action=""assign"" allkeywords=""0""><KEYWORDS>"
		end if
		
		' get GUIDs of all keywords in project (passed in from hidden variable)
		allKeywords = Split(allKeywords, ",")
		For i = 0 To UBound(allKeywords)-1
			 if request.form(allKeywords(i)) <> "" then
				'
				' A keyword has been selected
				'
				'
				' build up string containing all keyword settings for the page
				if previousKeywordsDict.Item(allKeywords(i)) <> "" then
					' keyword previously selected - no change
					xmlString = xmlString & "<KEYWORD guid="""&allKeywords(i)&""" changed=""0"" />"
				else
					' keyword was not previously selected - add it to the page
					xmlString = xmlString & "<KEYWORD guid="""&allKeywords(i)&""" changed=""1"" />"
				end if
			 else
			 	'
				' A keyword has not been selected
				'
				'
				if previousKeywordsDict.Item(allKeywords(i)) <> "" then
					' keyword previously select - remove it from the page
					xmlString = xmlString & "<KEYWORD guid="""&allKeywords(i)&""" changed=""1"" delete=""1"" />"
				end if
			 end if
		Next

		if Session("PageId") <> "" then
			' page
			xmlString = xmlString & "</KEYWORDS></PAGE></PROJECT></IODATA>"
		else
			' link
			xmlString = xmlString & "</KEYWORDS></LINK></PROJECT></IODATA>"
		end if
		xmlResult = sendXML(xmlString)

	end if
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
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title><%=Session("CmsWindowTitle")%> - Assign Keywords</title>
	
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta http-equiv="imagetoolbar" content="no" />
	
	<!-- stylesheets references -->
	<link rel="Stylesheet" type="text/css" href="rdUICheckboxAssignKeywords.css" />
	
	<!-- javascript references -->
	<script src="jquery-1.2.6.pack.js" type="text/javascript"></script>
</head>
<% if postBack <> "1" then %>
<body>

<form action="<%=Request.ServerVariables("SCRIPT_NAME")%>" name="assignKeywords" method="post">
<div class="page">
		<div class="page-inner">
		<div class="main">
			<div class="title"><h1>Assign Keywords</h1></div>
			
			<% if errorFound then %>
			<div class="error">
				<p><strong>Sorry, this page could not be displayed properly.</strong></p>
				<p><%=errorMessage%></p>
			</div>
			<% else %>
			
			<div class="content">
				<% 
				' intialise total lists of all keywords in project and those already assigned to the page
				allKeywords = ""
				previousKeywords = ""

				set objXMLDOMCategories=Server.CreateObject("Microsoft.XMLDOM")
				set objXMLDOMKeywords=Server.CreateObject("Microsoft.XMLDOM")
				
				'obtain list of categories
				xmlString = "<IODATA loginguid=""" & Session("LoginGuid") & """ sessionkey=""" & Session("SessionKey") & """><PROJECT><CATEGORIES action=""list""/></PROJECT></IODATA>"
				xmlResult = sendXML(xmlString)
				objXMLDOMCategories.LoadXml(xmlResult)
				set objCategoryList=objXMLDOMCategories.getElementsByTagName("CATEGORY")
				
				For i = 0 to (objCategoryList.Length-1)
					' loop through all categories found in current CMS project
					categoryName = objCategoryList.Item(i).getAttribute("value")  
					categoryGUID = objCategoryList.Item(i).getAttribute("guid")
					rowCount = 0 
					
					' we want display selected categories only - hide the rest
					if shownCategoriesDict.Count > 0 then
						if shownCategoriesDict.Item(categoryName) = categoryName then 
							response.write "<div class=""category"">" & vbCrLf
						else
							response.write "<div class=""category"" style=""display:none;"">" & vbCrLf
						end if 
					else
						response.write "<div class=""category"">" & vbCrLf
					end if

					response.write "<h2 class=""heading expanded"">"& categoryName &"</h2><div class=""category-content clearfix"">"

					' obtain keywords for the current category
					xmlString = "<IODATA loginguid=""" & Session("LoginGuid") & """ sessionkey=""" & Session("SessionKey") & """><PROJECT><CATEGORY guid="""& categoryGUID&"""><KEYWORDS action=""load""/></CATEGORY></PROJECT></IODATA>"
					xmlResult = sendXML(xmlString)
					objXMLDOMKeywords.LoadXml(xmlResult)
					
					set objKeywordList=objXMLDOMKeywords.getElementsByTagName("KEYWORD")
					For n = 0 to (objKeywordList.Length-1)
						keywordName = objKeywordList.Item(n).getAttribute("value")
						keywordGUID = objKeywordList.Item(n).getAttribute("guid")
						isChecked = ""

						' skip the first keyword that was returned as it is just the category guid
						if categoryGUID <> keywordGUID then
							' check to see if the current keyword is already assigned to the page
							if keywordsDict.Item(keywordGUID) = keywordGUID then 
								' keep a record of the previously selected keywords
								previousKeywords = previousKeywords & keywordGUID &  ","
								isChecked = " checked=""checked"" "
							end if
							
							'keep a count of how many checkboxes that have been displayed
							' as we only want to display 3 checkboxes per line
							rowCount = rowCount + 1
							if rowCount = 1 then
								response.write "<div class=""cb-row clearfix"">"  & vbCrLf
							elseif rowCount = 4 then
								response.write "</div><div class=""cb-row clearfix"">"  & vbCrLf
								rowCount = 1
							end if

							' display checkbox for this keyword on the page
							response.write "<div class=""cb""><input id="""&keywordGUID&""" type=""checkbox"" name="""&keywordGUID&""" value=""1"" "&isChecked&" /><label for="""&keywordGUID&""">"&keywordName&"</label></div>" & vbCrLf
							
							' keep a record of all keywords in the system
							allKeywords = allKeywords & keywordGUID & ","
						end if
						
					Next
					
					if allKeywords <> "" then
						response.write "</div></div></div>"  & vbCrLf
					end if
				
				Next

			end if %>
			</div>
		</div>
		</div>
		<div class="buttons clearfix">
			<div class="commandButton cancelButton"><img class="commandButtonIcon" id="cancelIcon" src="/cms/Icons/CommandButtons/Cancel.gif" alt="Cancel" />&nbsp;Cancel</div>
			<% if not errorFound then %><div class="commandButton okButton"><img class="commandButtonIcon" id="okIcon" src="/cms/Icons/CommandButtons/Ok.gif" alt="OK" />&nbsp;OK</div><% end if %>
		</div>
</div>

<input type="hidden" name="postBack" value="1" />
<input type="hidden" name="allKeywords" value="<%=allKeywords%>" />
<input type="hidden" name="previousKeywords" value="<%=previousKeywords%>" />

</form>
<% else %>
<body onLoad="closeWindowSubmit();">

<% end if %>
<script type="text/javascript">
//<![CDATA[
	$(document).ready(function(){
		// toggle categories
		$("h2.heading").click(function () {
			var checkboxes = $(this).next();
			if (checkboxes.is(':visible')) {
				$(this).addClass("closed");
			} else {
				$(this).removeClass("closed");
			}
			$(this).next().toggle();
		}); 
	
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
		$(".okButton").click(function () {
			document.assignKeywords.submit();
		});
	});  
	
	function closeWindowSubmit() {
		<% if Session("LastActiveModule") = "smartedit" then %>
		window.opener.ReloadEditedPage();
		<% end if %>
		window.close();
	}
//]]>  
</script>
</body>
</html>