//---------------------------------------------------------------------
// <copyright file="xslt.js">
//    This code is licensed under the MIT License.
//    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF 
//    ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED 
//    TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
//    PARTICULAR PURPOSE AND NONINFRINGEMENT.
// </copyright>
// <summary>A utility to transform XML using XSLT .</summary>
//---------------------------------------------------------------------



function TransformXml(xslText, xml) {
    var xslDocument, resultDocument, xsltProcessor, xslt, xslDoc, xslProc, transformedHtml;
  //  Diag.logVerbose("[HtmlDocumentGenerator._displayResult] Parse the xslt string");
    xslDocument = jQuery.parseXML(xslText); // Utils_Core.parseXml(xslText);

    if (window.ActiveXObject || "ActiveXObject" in window) {
        if (typeof (xml.transformNode) != "undefined") {
            transformedHtml = xml.transformNode(xslDocument);
            resultDocument = transformedHtml;
        }
        else {
          
            if (window.ActiveXObject || "ActiveXObject" in window) {
                xslt = new ActiveXObject("Msxml2.XSLTemplate");
                xslDoc = new ActiveXObject("Msxml2.FreeThreadedDOMDocument");
                xslDoc.loadXML(xslText);
                xslt.stylesheet = xslDoc;
                xslProc = xslt.createProcessor();
                xslProc.input = xml;
                xslProc.transform();
                transformedHtml = xslProc.output;
                resultDocument = transformedHtml;
            }
          
        }
    }
    else if (document.implementation && document.implementation.createDocument) {
        xsltProcessor = new window.XSLTProcessor();
        xsltProcessor.importStylesheet(xslDocument);
        transformedHtml = xsltProcessor.transformToDocument(xml);
        resultDocument = transformedHtml.documentElement.outerHTML;
    }
    return resultDocument;
 };

 function xmlEscape(s) {
     if ($.type(s) === "string") {
         return (s
                  .replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
            .replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\t/g, '&#x9;').replace(/\n/g, '&#xA;').replace(/\r/g, '&#xD;')
            .replace(/#/g, '').replace(/-/g, '&#45;').replace(/ü/g, '&#252;').replace(/&xA;/g, '&#10;').replace(/&xD;/g, '&#13;')
            .replace(/&x9;/g, '&#9;')
                 );
     }
     else {
         return s;
     }
 }
