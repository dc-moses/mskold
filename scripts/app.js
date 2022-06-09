//---------------------------------------------------------------------
// <copyright file="app.js">
//    This code is licensed under the MIT License.
//    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF 
//    ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED 
//    TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
//    PARTICULAR PURPOSE AND NONINFRINGEMENT.
// </copyright>
// <summary>the code responsilbe for menu extensionloading/Starting .</summary>
//---------------------------------------------------------------------

var showQueryExportMenuProvider = (function () {
    "use strict";
    return {
        ExportWorkItemQueriesDialog: function (properties) {
            try {
                window.appInsights.trackEvent("EnhancedExport.ExportWorkItemQuery");
            }
            catch (ex) {

            }

            VSS.getService("ms.vss-web.dialog-service").then(function (dialogSvc) {

                var extInfo = VSS.getExtensionContext();

            

                var contributionConfig = {
                    properties: properties,
                    serviceName: 'WorkItemQueryService'
                };
                
                if (contributionConfig.properties.query == null) {
                    contributionConfig.properties.query ={ name: "Product backlog FAKE", path: "FAKE PATH", id:587};
                }

                var dialogOptions = {
                    title: "Export Query - " + contributionConfig.properties.query.name + '  [ ' + contributionConfig.properties.query.path + ' ]',
                    width: 1000,
                    height: 700,
                    buttons: null
                };
                dialogSvc.openDialog(extInfo.publisherId + "." + extInfo.extensionId + "." + "exportForm", dialogOptions, contributionConfig);
            });
        },
        execute: function (actionContext) {
            this.ExportWorkItemQueriesDialog(actionContext);
        }
    };
}());


var showTestPlanExportMenuProvider = (function () {
    "use strict";
    return {
        ExportTestPlanDialog: function (properties) {
            try{
                window.appInsights.trackEvent("EnhancedExport.ExportTestPlan");
            }
            catch (ex) {

            }

            VSS.getService("ms.vss-web.dialog-service").then(function (dialogSvc) {

                var extInfo = VSS.getExtensionContext();


                var contributionConfig = {
                    properties: properties,
                    serviceName : 'TestPlanService'
                };
                contributionConfig.properties.testPlanId = properties.plan.id;
                contributionConfig.properties.teamProject = VSS.getWebContext().project.name;
                contributionConfig.properties.name = properties.plan.name;

                var dialogOptions = {
                    title: "Export Testplan - " + contributionConfig.properties.name,
                    width: 1000,
                    height: 700,
                    buttons: null
                };

                dialogSvc.openDialog(extInfo.publisherId + "." + extInfo.extensionId + "." + "exportForm", dialogOptions, contributionConfig);

              
            });
        },
        execute: function (actionContext) {
            this.ExportTestPlanDialog(actionContext); //  actionContext[0]["System.Title"]);
        }
    };
}());


var showTestSuiteExportMenuProvider = (function () {
    "use strict";
    return {
        ExportTestSuiteDialog: function (properties) {
            try {
                window.appInsights.trackEvent("EnhancedExport.ExportTestPlan");
            }
            catch (ex) {

            }

            VSS.getService("ms.vss-web.dialog-service").then(function (dialogSvc) {

                var extInfo = VSS.getExtensionContext();


                var contributionConfig = {
                    properties: properties,
                    serviceName: 'TestPlanService'
                };
                contributionConfig.properties.testPlanId = properties.plan.id;
                contributionConfig.properties.startSuiteId = properties.suite.id;
                contributionConfig.properties.teamProject = VSS.getWebContext().project.name;
                contributionConfig.properties.name = properties.plan.name;

                var dialogOptions = {
                    title: "Export TestSuite - " + contributionConfig.properties.suite.title,
                    width: 1000,
                    height: 700,
                    buttons: null
                };

                dialogSvc.openDialog(extInfo.publisherId + "." + extInfo.extensionId + "." + "exportForm", dialogOptions, contributionConfig);


            });
        },
        execute: function (actionContext) {
            this.ExportTestSuiteDialog(actionContext); //  actionContext[0]["System.Title"]);
        }
    };
}());


