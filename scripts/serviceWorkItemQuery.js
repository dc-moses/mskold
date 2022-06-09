//---------------------------------------------------------------------
// <copyright file="serviceWorkItemQuert.js">
//    This code is licensed under the MIT License.
//    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF 
//    ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED 
//    TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
//    PARTICULAR PURPOSE AND NONINFRINGEMENT.
// </copyright>
// <summary>A service so get the resutl of work item queries as xml .</summary>
//---------------------------------------------------------------------

try { angular.module("ExportQuery.services") } catch (err) { angular.module('ExportQuery.services', []); }
angular.module('ExportQuery.services').service('WorkItemQueryService', function () {


   

             this.ExportName = "";
             this.ServiceName = "WorkItemQuery";

             this.LoadFieldTypes = function () {
                 var deferred = $.Deferred();
                 VSS.require([ "TFS/WorkItemTracking/RestClient", "TFS/WorkItemTracking/Contracts"],
                    function (RestClient, _TFS_Wit_Contract) {
                        var fieldTypes = {};
                        var fieldTypesConstants = _TFS_Wit_Contract.FieldType;

                        var wiClient = RestClient.getClient();
                        wiClient.getFields().then(function (data) {
                            data.forEach(function (f) {
                                fieldTypes[fieldTypesConstants[f.type]] += ( f.referenceName+"|");
                            })
                            deferred.resolve(fieldTypes);
                        });
                    });
                 return deferred.promise();
             };

             this.ConvertResultList2Xml = function (ret, fieldTypes) {
                 var s = "<result collectionUrl='" + VSS.getWebContext().collection.uri + "'>'";

                 s = s + "<columns>"
                 ret.Columns.forEach(function (c) {
                     var width = 75;
                     if (c.referenceName == "System.Title") {
                         width = 150;
                     }
                     if (c.referenceName.indexOf("To") > 0) {
                         width = 100;
                     }
                     s = s + "<" + xmlEscape(c.referenceName) + " name='" + xmlEscape(c.name) + "' width='" + width + "'/>";
                 });
                 s = s + "</columns>";


                 ret.QueryResults.forEach(function (item) {
                     s += BuildXMLforNode(item, fieldTypes);
                 });

                 s += '</result>';

                 return jQuery.parseXML(s);
             };

             this.InitDecorator = function (data) {
                 var deferred = $.Deferred()
                 this.ProgressMessage("Fetching extra data");

                 VSS.require(["VSS/Controls", "VSS/Service", "TFS/WorkItemTracking/RestClient",
                 "TFS/WorkItemTracking/Contracts", "VSS/Utils/Html"],
                 function (Controls, VSS_Service, TFS_Wit_WebApi, TFS_Wit_Contract) {

                     var witClient = VSS_Service.getCollectionClient(TFS_Wit_WebApi.WorkItemTrackingHttpClient);
                     this.ProgressMessage("Fetching work item types");
                     witClient.queryById(queryId).then(function (wits) {
                         this._workItemTypes = wits;
                         deferred.resolve(data);
                     });
                 });

                 return deferred.promise();
             }

             this.Decorate = function (xml) {
                 this.ProgressMessage("Adding Extra data to query result");
                 var wiNodes = xml.find('//workiem');
                 wiNodes.forEach(function (n) {
                     n.attrib('color') = this._workItemTypes.filter(function (i) { return i.type == i.type; }).color;
                 })
             }


             this.execute = function (progressCallback) {
                 var deffered = $.Deferred();

                 var srv = this;
                 srv.progressCallback = progressCallback;
                 
                 var queryFetcher = QueryFetcher();
                 
                 var queryPromise = queryFetcher.executeQuery(this.queryId, progressCallback);
                 var fieldTypesPromise = srv.LoadFieldTypes();
                 
                 Q.all([queryPromise, fieldTypesPromise]).then(
                      function (data) {
                          //   var defferedInitDecorator= InitDecorator(ret);
                          srv.progressMessage("Converting query result to xml");

                          srv.QueryResultXml = srv.ConvertResultList2Xml(data[0], data[1]);

                          deffered.resolve(srv.QueryResultXml)
                      },
                      function (err) {
                          deffered.reject(err);
                      }
                  );

                 return deffered.promise();
             };

             this.init = function (configutation) {
                 this.ExportName = configutation.properties.query.name;
                 this.queryId = configutation.properties.query.id;

             };

             this.progressMessage = function (s) {
                 if (this.progessCallback != null) {
                     this.progressCallback(s);
                 }
             }

         });



    function getFieldType(field, fieldTypes){
        for (var type in fieldTypes) {
            if (fieldTypes.hasOwnProperty(type)) {
                if(fieldTypes[type].indexOf(field+"|")!=-1){
                    return type;
                }
            }

        }
        return "";
    }


    function renderLinksExml(links) {
        var linksXml = "";
        if (links != null) {
            var linksXml = "<links>";
            links.forEach( function(l) {
                linksXml += "<link "
                if(l.title!=null){
                    linksXml +="title ='" + l.title + "' ";
                }
                if (l.url != null) {
                    linksXml += "url='" + l.url + "' ";
                }
                if (l.rel != null) {
                    linksXml += "rel='" + l.rel + "' ";
                }

                for (var a in l.attributes) {
                    if (l.attributes.hasOwnProperty(a)) {
                        linksXml += a + "='" + l.attributes[a] + "' ";
                    }
                };
                linksXml += " />";
            });
            linksXml += "</links>";
        }
        return linksXml;
    }


    function BuildXMLforNode(wi, fieldTypes) {
        var s = "";
        try {
            s = '<workitem id=\'' + wi.id + '\' type=\'' + wi.fields['System.WorkItemType'] + '\'  state=\'' + wi.fields['System.State'] + '\' >'

            for (var property in wi.fields) {
                if (wi.fields.hasOwnProperty(property)) {
                    var val = "";
                    var attrib = "";
                    switch(property){
                        case "System.AcceptedBy":
                        case "System.ActivatedBy":
                        case "System.AssignedTo":
                        case "System.CalledBy":
                        case "System.ChangedBy ":
                        case "System.ClosedBy":
                        case "System.CreatedBy":
                        case "System.ResolvedBy":
                        case "System.ReviewedBy":
                        case "System.AssignedTo":
                        case "System.AssignedTo":
                            var assignedto = wi.fields[property];
                            var name = assignedto.displayName;
                            var email = null;
                            if (name.indexOf("<") > 0) {
                                name = name.split("<")[0];
                                email = assignedto.split("<")[1].split(">")[0].replace(">", "");
                            }

                            attrib = "name='" + name +"'"+  (email != null ? "' email='" + email + "' md5='" + $.md5(email) + "'" : "");
                            val = xmlEscape(wi.fields[property].displayName);
                            break;
                        case "System.Tags":
                            var tags = "";
                            wi.fields[property].split(";").forEach(function (t) {
                                tags += "<tag>" + xmlEscape(t) + "</tag>"
                            });

                            val = tags + xmlEscape(wi.fields[property]);
                            break;

                        default:
                            switch (getFieldType(property, fieldTypes) ) {
                                case "Html":
                                    var desc = wi.fields[property];

                                    desc = desc.replace(/(\r\n|\n|\r)/gm, " ");
                                    val = secureHTMLisXml(desc);
                                    break;
                                
                                default:
                                    val = xmlEscape(wi.fields[property]);
                            }   
                    }

                }
                s += '<' + xmlEscape(property) + (attrib != '' ? ' ' + attrib : '') + '>' + val + '</' + xmlEscape(property) + '>';

            }

            s = s.replace(/&nbsp;/g, "&#160;");

            s += renderLinksExml(wi.relations);

            if (wi.childs != null) {
                wi.childs.forEach(function (item) {
                    s += BuildXMLforNode(item, fieldTypes);
                });
            }

            s += '</workitem>'

            //Vallidate 
            jQuery.parseXML(s);
        }
        catch (ex) {
            console.log("Error parsing wi " + wi.id + " " + ex.message);
            console.log(s);
            s = "";
        }

        return s;
    }


