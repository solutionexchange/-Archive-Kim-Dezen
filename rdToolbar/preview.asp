<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html>
<head>
<title>Page Preview</title>
</head>
<frameset rows="10%, 90%">
  <frame src="preview-return.asp" frameborder="0" scrolling="no" marginwidth="0" marginheight="0">
  <frame src="/CMS/ioRD.asp?Action=Preview&Pageguid=<%= request.querystring("pageguid")%>" frameborder="0" scrolling="yes" marginwidth="0" marginheight="0">
</frameset>
</html>
