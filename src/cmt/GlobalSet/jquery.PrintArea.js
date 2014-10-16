﻿/**
 *  Version 2.1
 *      -Contributors: "mindinquiring" : filter to exclude any stylesheet other than print.
 *  Tested ONLY in IE 8 and FF 3.6. No official support for other browsers, but will
 *      TRY to accomodate challenges in other browsers.
 *  Example:
 *      Print Button: <div id="print_button">Print</div>
 *      Print Area  : <div class="PrintArea"> ... html ... </div>
 *      Javascript  : <script>
 *                       $("div#print_button").click(function(){
 *                           $("div.PrintArea").printArea( [OPTIONS] );
 *                       });
 *                     </script>
 *  options are passed as json (json example: {mode: "popup", popClose: false})
 *
 *  {OPTIONS} | [type]    | (default), values      | Explanation
 *  --------- | --------- | ---------------------- | -----------
 *  @mode     | [string]  | ("iframe"),"popup"     | printable window is either iframe or browser popup
 *  @popHt    | [number]  | (500)                  | popup window height
 *  @popWd    | [number]  | (400)                  | popup window width
 *  @popX     | [number]  | (500)                  | popup window screen X position
 *  @popY     | [number]  | (500)                  | popup window screen Y position
 *  @popTitle | [string]  | ('')                   | popup window title element
 *  @popClose | [boolean] | (false),true           | popup window close after printing
 *  @strict   | [boolean] | (undefined),true,false | strict or loose(Transitional) html 4.01 document standard or undefined to not include at all (only for popup option)
 */
(function ($) {
	var counter = 0;
	var modes = { iframe: "iframe", popup: "popup" };
	var defaults = {
	    mode: modes.popup,
	    popHt: 600,
	    popWd: 400,
	    popX: 200,
	    popY: 200,
	    popTitle: '.',
	    popClose: true,
	    formData: { title: "", vesselName: "", vesselList: "", vesselDate: "" }
	};

	var settings = {};//global settings

	$.fn.printArea = function (options) {
		$.extend(settings, defaults, options);

		counter++;
		var idPrefix = "printArea_";
		$("[id^=" + idPrefix + "]").remove();
		var ele = getFormData($(this));

		settings.id = idPrefix + counter;

		var writeDoc;
		var printWindow;

		switch (settings.mode) {
			case modes.iframe:
				var f = new Iframe();
				writeDoc = f.doc;
				printWindow = f.contentWindow || f;
				break;
			case modes.popup:
				printWindow = new Popup();
				writeDoc = printWindow.doc;
		}

		writeDoc.open();
		writeDoc.write(docType() + "<html>" + getHead() + getBody(ele) + "</html>");
		writeDoc.close();

		printWindow.focus();
		printWindow.print();
        
		//if (settings.mode == modes.popup && settings.popClose)
		//	printWindow.close();
	}

	function docType() {
		if (settings.mode == modes.iframe || !settings.strict) return "";

		var standard = settings.strict == false ? " Trasitional" : "";
		var dtd = settings.strict == false ? "loose" : "strict";

		return '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01' + standard + '//EN" "http://www.w3.org/TR/html4/' + dtd + '.dtd">';
	}

	function getHead() {
		var head = "<head><title>" + settings.popTitle + "</title>";
		$(document).find("link")
            .filter(function () {
            	return $(this).attr("rel").toLowerCase() == "stylesheet";
            })
            .filter(function () { // this filter contributed by "mindinquiring"
            	var media = $(this).attr("media") + "";

            	return (media.toLowerCase() == "" || media.toLowerCase() == "print")
            })
            .each(function () {
            	head += '<link type="text/css" rel="stylesheet" href="' + $(this).attr("href") + '" >';
            });
		head += "<link href='/GlobalSet/sg.css' rel='stylesheet' type='text/css' />";
		head += "<style>.logo{width: 100%;height:100px;font: bold 11px   Arial, Helvetica, sans-serif;} .table{border: 1px solid #000000;border-spacing: 0;font:  11px   Arial, Helvetica, sans-serif;} table{border: 1px solid #000000;border-spacing: 0;font:  11px   Arial, Helvetica, sans-serif;}th {font: bold 12px   Arial, Helvetica, sans-serif;border: 1px solid #CCCCCC;color: #000000;text-align: left;padding: 6px 6px 6px 12px;background-color:#ccc ;}td {border: 1px solid #CCCCCC;background: #fff;padding: 6px 6px 6px 12px;text-align: left;color: #000000;}</style>";
		head += '</head>';
		//head += '<link type="text/css" rel="stylesheet" href="http://localhost/sisgef/css/table.css" media="print">';

		return head;
	}

	function getBody(printElement) {
	    //return '<body><div class="logo"><img src="../sisgef/images/logo.png" ></div><h3>' + settings.popTitle + '</h3>' + $(printElement).html() + '</body>';
	    var body
            = '<body style="background-image: url(../image/Logo.gif); background-position: center; background-repeat: no-repeat; background-position-y: top;">'
            + '<div style="margin: 0px auto; margin-top: 100px; text-align: center;">'
            + '<span style="font-size: large;">上 林 公 證 有 限 公 司<br />KINGSWOOD SURVEY & MEASURER LTD</span>'
            + '<br />'
            + '<br />'
            + '<span style="font-size: small;">版權所有 &copy 2013. All Rights Reserved.</span>'
            + '</div>';
	    if (settings.formData.title.length > 0) {
	        body +=
                "<table cellspacing=\"1\" style=\"background-color: #c9e0f8; border-width: 0px; width: 95%; margin: 0 auto;\">" +
                "<tr><td colspan=\"3\" style=\"text-align: right; font-family: Arial; font-size: 14.0pt;\">結關日期：" + settings.formData.vesselDate + "</td></tr>" +
                "<tr><td colspan=\"3\" style=\"text-align: center; font-size: 18.0pt\">" + settings.formData.title + "</td></tr>" +
				"<tr><td colspan=\"3\"><hr /></td></tr>" +
				"<tr valign=\"middle\">" +
					"<td style=\"width: 10%;\"></td>" +
					"<td style=\"width: 40%; font-family: Arial; font-size: 12.0pt\">航次：" + settings.formData.vesselList + "</td>" +
					"<td style=\"width: 50%; font-family: Arial; font-size: 12.0pt\">船名：" + settings.formData.vesselName + "</td>" +
                "</tr>" +
				//"<tr valign=\"middle\">" +
				//	"<td style=\"width: 20%;\"></td>" +
				//	"<td style=\"width: 20%; font-family: Arial; font-size: 14.0pt\"></td>" +
				//	"<td style=\"width: 60%;\"></td>" +
                //"</tr>" +
                "<tr><td colspan=\"3\"><hr /></td></tr>" +
                "</table>";
	    }
	    var tempBody = $(printElement).clone();
	    tempBody.find("td.print-remove").remove();
	    body += tempBody.html();// $(printElement).remove(".print-remove").html();
	    body += '</body>';
	    return body;
	}

	function getFormData(ele) {
		$("input,select,textarea", ele).each(function () {
			// In cases where radio, checkboxes and select elements are selected and deselected, and the print
			// button is pressed between select/deselect, the print screen shows incorrectly selected elements.
			// To ensure that the correct inputs are selected, when eventually printed, we must inspect each dom element
			var type = $(this).attr("type");
			if (type == "radio" || type == "checkbox") {
				if ($(this).is(":not(:checked)")) this.removeAttribute("checked");
				else this.setAttribute("checked", true);
			}
			else if (type == "text")
				this.setAttribute("value", $(this).val());
			else if (type == "select-multiple" || type == "select-one")
				$(this).find("option").each(function () {
					if ($(this).is(":not(:selected)")) this.removeAttribute("selected");
					else this.setAttribute("selected", true);
				});
			else if (type == "textarea") {
				var v = $(this).attr("value");
				if ($.browser.mozilla) {
					if (this.firstChild) this.firstChild.textContent = v;
					else this.textContent = v;
				}
				else this.innerHTML = v;
			}
		});
		return ele;
	}

	function Iframe() {
		var frameId = settings.id;
		var iframeStyle = 'border:0;position:absolute;width:0px;height:0px;left:0px;top:0px;';
		var iframe;

		try {
			iframe = document.createElement('iframe');
			document.body.appendChild(iframe);
			$(iframe).attr({ style: iframeStyle, id: frameId, src: "" });
			iframe.doc = null;
			iframe.doc = iframe.contentDocument ? iframe.contentDocument : (iframe.contentWindow ? iframe.contentWindow.document : iframe.document);
		}
		catch (e) { throw e + ". iframes may not be supported in this browser."; }

		if (iframe.doc == null) throw "Cannot find document.";

		return iframe;
	}

	function Popup() {
		var windowAttr = "location=yes,statusbar=no,directories=no,menubar=no,titlebar=no,toolbar=no,dependent=no";
		windowAttr += ",width=" + settings.popWd + ",height=" + settings.popHt;
		windowAttr += ",resizable=yes,screenX=" + settings.popX + ",screenY=" + settings.popY + ",personalbar=no,scrollbars=no";

		var newWin = window.open("", "_blank", windowAttr);

		newWin.doc = newWin.document;

		return newWin;
	}
})(jQuery);