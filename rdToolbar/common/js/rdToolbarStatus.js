var rdToolbarStatus = (function() {

	var vars = {
			// page and script properties
			pageguid: 'setinTPL',
			pagemode: '0',
			defaultextension: 'setinTPL',
			pluginLocation: '/cms/plugins/rdtoolbar/',
			processingPage: 'rdToolbarStatus.asp'
	};
		
	// run script
	var loadSettings = function() {		
		// insert toolbar html into page
		jQuery("#reddot-menubar").html(buildToolbar());
		
		// add handlers for each button in the toolbar here
		jQuery("#assign-keywords").click(function() {
            window.open("/cms/plugins/rdUICheckboxAssignKeywordsv10/rdUICheckboxAssignKeywords.asp?pageguid="+vars.pageguid, "AssignKeywords", "top=20, left=20, width=746, height=500, scrollbars=yes, resizable=no");
        });
		
		jQuery("#submit-worflow").click(function() {
            window.open(vars.pluginLocation + "submit-workflow.asp?pageguid="+vars.pageguid, "SubmitWorkflow", "top=20, left=20, width=746, height=300, scrollbars=yes, resizable=no");
        });
	
		// assign filename to page for first time
		jQuery.ajax({
			async: true,
			cache: false,
			type: "GET",
			timeout: 5000,
			dataType: 'xml',
			url: buildProcessingURL(),
			success: function(data) {
				 jQuery(data).find("status").each(function() {  
					jQuery("#reddot-menu-bar-pageinfo #status-holder").html((jQuery(this).text()));
					
				 });  
 
				 jQuery(data).find("filename").each(function() {  
					jQuery("#reddot-menu-bar-pageinfo #filename-holder").html((jQuery(this).text()));
				 }); 
				 
				 jQuery(data).find("createdby").each(function() {  
					jQuery("#reddot-menu-bar-pageinfo #created-holder").html((jQuery(this).text()));
				 }); 
				 
				 jQuery(data).find("modifiedby").each(function() {  
					jQuery("#reddot-menu-bar-pageinfo #modified-holder").html((jQuery(this).text()));
				 }); 
				 
				 jQuery(data).find("pageid").each(function() {  
					jQuery("#reddot-menu-bar-pageinfo #pageid-holder").html((jQuery(this).text()));
				 }); 
				 
				 jQuery("#reddot-menu-bar-pageinfo #pageguid-holder").html(vars.pageguid);
			},
			error:function (xhr, ajaxOptions, thrownError)
			{
				// handle any errors (404, 500 etc)
				jQuery("#reddot-menu-bar-pageinfo #error-messages").html('The following error occured:' + xhr.status + ' ' + thrownError);
			}
		});
		
		return true;
		
	};
	
	var buildToolbar = function()
	{	
		var toolbarContent = "" +
		"<div id=\"reddot-menu-bar\" class=\"clearfix\" style=\"background-color:#fff;z-index:1000000;\">" +
			"<div id=\"reddot-menu-bar-buttons\">";
				if (vars.pagemode == "0")
				{
					toolbarContent += "<div class=\"open-page reddot-menu-bar-button\">" +
						"<a href=\"./PreviewHandler.ashx?Action=RedDot&Mode=2&PageGUID=" + vars.pageguid + "&EditPageGUID=" + vars.pageguid + "&Opener=1#CloseRedDot\" class=\"image-replace\" title=\"Open Page\">Open Page<span></span></a>" +
					"</div>";
				}
				else
				{
					toolbarContent += "<div class=\"close-page reddot-menu-bar-button\">" +
						"<a href=\"./PreviewHandler.ashx?Action=RedDot&Mode=4&PageGUID=" + vars.pageguid + "&EditPageGUID=" + vars.pageguid + "&Opener=0&EditPageLocked=1\" class=\"image-replace\">Close Page<span></span></a>" +
					"</div>";
				}
				toolbarContent += "" +
	
				// start custom buttons
				// DEVELOPERS: add the required custom buttons to the toolbar here 

				"<div class=\"assign-keywords reddot-menu-bar-button\">" +
					"<a href=\"#\" class=\"image-replace\" title=\"Assign Keywords\" id=\"assign-keywords\">Assign Keywords<span></span></a>" +
				"</div>" +
				
				"<div class=\"submit-worflow reddot-menu-bar-button\">" +
					"<a href=\"#\" class=\"image-replace\" title=\"Submit to Workflow\" id=\"submit-worflow\">Submit to Workflow<span></span></a>" +
				"</div>" +

				// end custom buttons
				//
				
				"<div class=\"rdtoolbar\"><a href=\"http://www.kimdezen.com\" target=\"_blank\"><img src=\""+ vars.pluginLocation + "common/images/buttons/but_rdtoolbar.gif\" width=\"91\" height=\"32\" border=\"0\" alt=\"RDToolbar : Developed by Kim Dezen (www.kimdezen.com)\" style=\"padding-right: 35px;\"></a></div>" +
			"</div>" +
			"<div id=\"reddot-menu-bar-pageinfo\" class=\"clearfix\">" +
				"<div id=\"reddot-menu-bar-pageinfo-inner\" class=\"clearfix\">" +
					"<div class=\"column-1\"></div>" +
						"<div class=\"column-2\">" +
							"<div class=\"column-2-inner clearfix\">" +
								"<div class=\"column\">" +
									"<p><strong>Status:</strong> <span id=\"status-holder\"><img src=\""+ vars.pluginLocation + "common/images/icons/loading.gif\" height=\"12\" width=\"12\" alt=\"\" /></span></p>" +
									"<p><strong>Filename:</strong> <span id=\"filename-holder\"><img src=\""+ vars.pluginLocation + "common/images/icons/loading.gif\" height=\"12\" width=\"12\" alt=\"\" /></span></p>" +
								"</div>" +
								"<div class=\"column\">" +
									"<p><strong>Created:</strong> <span id=\"created-holder\"><img src=\""+ vars.pluginLocation + "common/images/icons/loading.gif\" height=\"12\" width=\"12\" alt=\"\" /></span></p>" +
									"<p><strong>Modified:</strong> <span id=\"modified-holder\"><img src=\""+ vars.pluginLocation + "common/images/icons/loading.gif\" height=\"12\" width=\"12\" alt=\"\" /></span></p>" +
								"</div>" +
							"<div class=\"column\">" +
								"<p><strong>PageId:</strong> <span id=\"pageid-holder\"><img src=\""+ vars.pluginLocation + "common/images/icons/loading.gif\" height=\"12\" width=\"12\" alt=\"\" /></span></p>" +
								"<p><strong>PageGUID:</strong> <span id=\"pageguid-holder\"><img src=\""+ vars.pluginLocation + "common/images/icons/loading.gif\" height=\"12\" width=\"12\" alt=\"\" /></span></p>" +
							"</div>" +
						"</div>" +
					"</div>" +
				"</div>" +
			"</div>" +
			"<span id=\"error-messages\"></span>" +
		"</div>";
		return toolbarContent;
	}
	
	var buildProcessingURL = function()
	{
		// build up querystring to pass to processing page
		var processingURL = vars.pluginLocation + vars.processingPage + "?pageguid=" + vars.pageguid + "&defaultextension=" + vars.defaultextension;
		return processingURL;
	};
	
	return  { 
		vars: vars,
		init: function() {
			loadSettings();
		}
	};
	
})(); 


