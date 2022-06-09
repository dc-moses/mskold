//---------------------------------------------------------------------
// <copyright file="common.js">
//    This code is licensed under the MIT License.
//    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF 
//    ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED 
//    TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
//    PARTICULAR PURPOSE AND NONINFRINGEMENT.
// </copyright>
// <summary>A set of shared common functions.</summary>
//---------------------------------------------------------------------


function iframeDataURITest(src) {
    var support,
        iframe = document.createElement('iframe');

    iframe.style.display = 'none';
    iframe.setAttribute('src', src);

    document.body.appendChild(iframe);

    try
    {
        support = !!iframe.contentDocument;
    }
    catch (e) {
        support = false;
    }

    document.body.removeChild(iframe);
    return support;
}

function selectText(element) {
    var doc = document;
    var range, selection;
    
    if (doc.body.createTextRange) {
        range = document.body.createTextRange();
        range.moveToElementText(element);
        range.select();
    } else if (window.getSelection) {
        selection = window.getSelection();
        range = document.createRange();
        range.selectNodeContents(element);
        selection.removeAllRanges();
        selection.addRange(range);
    }
}

Array.prototype.unique = function () {
    var n = {}, r = [];
    for (var i = 0; i < this.length; i++) {
        if (!n[this[i]]) {
            n[this[i]] = true;
            r.push(this[i]);
        }
    }
    return r;
}

wrapInDiv = function (content) {
    return "<div>" + content + "</div>";
}


function getHubUrl(contributionId) {
    var context = VSS.getWebContext();
    var extCont = VSS.getExtensionContext();

    var collectionWithTrailingSlash = context.collection.uri;
    if (collectionWithTrailingSlash.charAt(collectionWithTrailingSlash.length-1) != '/') {
        collectionWithTrailingSlash = collectionWithTrailingSlash + "/";
    }
    
    return collectionWithTrailingSlash + context.project.name + "/_apps/hub/" + extCont.publisherId + "." + extCont.extensionId + "." + contributionId.replace('.', '-');
}

function SaveFile(data, fileName, type)
{
    if (window.navigator.msSaveOrOpenBlob != null) {
        SaveFileMsBlob(data, fileName, type);
    }
    else {
        SaveFileDataUri(
                         btoa(unescape(encodeURIComponent(data))),
                         fileName,
                         type);
    }
}

function SaveFileDataUri(data, fileName, type)
{
    var a = document.body.appendChild(
        document.createElement("a")
    );
   
    a.download = fileName;
    a.href = "data:text/"+ type +";base64," + data;
    a.innerHTML = "download";
    a.click();
    delete a;
}

function SaveFileMsBlob(data, fileName, type) {
    var blob = new Blob([data], { type: "text/" + type });
    window.navigator.msSaveOrOpenBlob(blob, fileName);
}

function ConvertImage2Inline(html, $http) {
    //$(html).find('img').each(function (img) {
    //    var url = $(this).attr('src');
    //    $http.get(url).then(function (data) {
    //        $(this).attr('src', 'data:base64,' + btoa(data));
            
    //    });
    //    $(this).attr('src', 'file');

    //});

    return html;
}

function secureHTMLisXml(html) {

    var htElem = $.parseHTML("<html xmlns='http://www.w3.org/1999/xhtml' /><!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>" + html);
    var s = "<div>";
    htElem.forEach(function (e) {
        if (e.outerHTML != null) {
            s = s + e.outerHTML
        }
        else if (e.innerHTML != null) {
            s = s + e.innerHTML
        }
        else if (e.nodeValue != null) {
            s = s + xmlEscape(e.nodeValue);
        }
    });

    s = s + "</div>";
    s = s.replace(/&nbsp;/g, "&#160;");
    s = hmlNormalizer.sanitize(s);
    s = hmlNormalizer.normalize(s);
    try {

        jQuery.parseXML(s);
    }
    catch (ex) {

        s = "<div>ERROR parsign formated text: " + ex.message + "</div>";
    }
    return s;

}