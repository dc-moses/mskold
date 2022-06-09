//---------------------------------------------------------------------
// <copyright file="serviceTestPlan.js">
//    This code is licensed under the MIT License.
//    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF 
//    ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED 
//    TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
//    PARTICULAR PURPOSE AND NONINFRINGEMENT.
// </copyright>
// <summary>A service to fetch test plan data.</summary>
//---------------------------------------------------------------------
var FEATURE_TestResultAttatchments= false;

var Utils_Date = null;
var Utils_String = null;

var DownloadArtifactUrl = "";


try { angular.module("ExportQuery.services") } catch (err) { angular.module('ExportQuery.services', []); }
var TFS_Wit_Contract;
var TestContracts;
var testPointColumns = [
    { name: "Outcome", property: "outcome" },
    { name: "Tester", property: "assignedTo.displayName" },
    { name: "Configuration", property: "configuration.name" },
    { name: "Run by", property: "lastResultDetails.runBy.displayName" },
    { name: "Date completed", property: "lastResultDetails.dateCompleted" },
    { name: "Duration in seconds", property: "lastResultDetails.duration" },
    { name: "Build number", property: "lastRunBuildNumber" }
    ];



angular.module('ExportQuery.services').service('TestPlanService', function () {

    this.ServiceName = "TestPlan";
    this.testPlanId = null;
    this.startSuiteId = null;
    this.teamProject = null;
    this.client = null;

    this.execute = function (progressCallback) {
        var deffered  = $.Deferred();
        var srv = this;
        srv.progressCallback = progressCallback;
        var xml="";
        var root = {};

        var rootXml = $.parseXML("<planAndSuites></planAndSuites>")

        DownloadArtifactUrl = VSS.getWebContext().collection.uri + VSS.getWebContext().project.name + "/_api/_testresult/DownloadAttachment?attachmentId={id}&__v=5";

        srv.xml = rootXml;

        VSS.require(["VSS/WebApi/Constants", "VSS/Service", "TFS/TestManagement/RestClient", "VSS/Utils/Date", "VSS/Utils/String", "TFS/WorkItemTracking/RestClient", "TFS/WorkItemTracking/Contracts", "TFS/TestManagement/Contracts"],
            function (WebApi_Constants, Service, RestClient, UtilsDate, UtilsString, TFS_Wit_WebApi, _TFS_Wit_Contract, _TestContracts)
            {

                TFS_Wit_Contract = _TFS_Wit_Contract;
                TestContracts = _TestContracts;

                Utils_Date = UtilsDate;
                Utils_String = UtilsString;

                // Get an instance of the client
                var client = Service.VssConnection.getConnection()
                    .getHttpClient(RestClient.TestHttpClient2_3, WebApi_Constants.ServiceInstanceTypes.TFS);

                var witClient = Service.getCollectionClient(TFS_Wit_WebApi.WorkItemTrackingHttpClient);
                srv.ProgressMessage("Fetching test plan");
                client.getPlanById(srv.teamProject, srv.testPlanId).then(function (testPlan) {
                    root = testPlan;
                    var testPlanTag= AddTestPlan2Xml(rootXml.documentElement, testPlan);

                    if(srv.ProgressMessage("Fetching test suites")){
                        client.getTestSuitesForPlan(srv.teamProject, srv.testPlanId, true).then(function (testSuites) {
                            setSuitePath(testSuites, null);
                            if (srv.startSuiteId == null) {
                                root.rootTestSuite = testSuites[0];
                                srv.suiteList = testSuites;
                            }
                            else{
                                root.rootTestSuite = testSuites.filter(function (i) { return i.id === srv.startSuiteId })[0];
                                srv.suiteList = srv.FindChildSuites(root.rootTestSuite, testSuites, [root.rootTestSuite]);
                            }

                            if (srv.ProgressMessage("Fetching test cases")) {
                                srv.preFetchTestCases(client, witClient, srv.suiteList).then(function (tcList) {
                                    srv.testCaseList = tcList;

                                    if (srv.ProgressMessage("Fetching test results ")) {
                                        srv.fetchTestResults(client, witClient, srv.suiteList).then(function (tcResList) {

                                            //MapTestResults(tcResList);
                                            if (srv.ProgressMessage("Building test suites and test case data")) {
                                               
                                                AddFlatSuiteList2PlanXml(srv, rootXml.documentElement);

                                                if (srv.ProgressMessage("Building test suite hierarchy")) {
                                                    var suiteHirerchyRoot = appendTags(testPlanTag, "suiteHierarchy");
                                                    AddSuiteHierarchy2PlanXml(srv, suiteHirerchyRoot, root.rootTestSuite);
                                                }
                                                //console.log((new XMLSerializer()).serializeToString(srv.xml));
                                            }
                                            deffered.resolve(srv.xml);
                                        
                                        }, function (err) {
                                            deffered.reject(err);
                                        }); //Prefetched end
                                    }
                                });
                            }

                        },function (err) {
                            deffered.reject(err);
                        });//getTestSuitesFOr
                    }
                },function (err) {
                    deffered.reject(err);
                });//getTestSuitesFOr
            });

        return deffered.promise()
    };

    this.FindChildSuites = function (suite, inputlist, list) {
        var ctrl = this;
        var childs = inputlist.filter(function (i) { return  i.parent !=null && i.parent.id === suite.id.toString() });
        list = list.concat(childs);
        childs.forEach(function (c) {
            list = ctrl.FindChildSuites(c, inputlist, list);
        })
        return list;
    };

    this.init = function (configutation) {
       
        this.testPlanId = configutation.properties.testPlanId;
        this.startSuiteId = configutation.properties.startSuiteId;
        this.teamProject = configutation.properties.teamProject;
        this.ExportName = configutation.properties.name;

    };


    this.ProgressMessage = function (message) {
        if (this.progressCallback != null) {
            return this.progressCallback(message);

        }
        else {
            return true;
        }
    }

    this.preFetchTestCases = function (client, witClient, suitelist)
    {
        var preFetchdeferred = $.Deferred();
        var idList = [];
        var suiteFetchedCount = suitelist.length;

        suitelist.forEach(function (suite) {
            client.getTestCases(suite.project.name, suite.plan.id, suite.id).then(
                function (tcList) {

                    suite._suiteTestCaseList = tcList;
                    suiteFetchedCount--;
                    idList = idList.concat(tcList.map(function (o) { return o.testCase.id }));

                    if (suiteFetchedCount == 0) {
                        //We got the list now fetch the records.
                        if (idList.length > 0) {
                            var bulkFetcher = new BulkWIFetcher()
                            bulkFetcher.FetchAllWorkItems(witClient, idList, undefined, null, TFS_Wit_Contract.WorkItemExpand.Fields)
                                .then(function (queryResultList) {
                                    preFetchdeferred.resolve(queryResultList);
                                }, function (error) {
                                    preFetchdeferred.reject(error);
                                });
                        }
                        else {
                            preFetchdeferred.resolve(null);
                        }

                    }
                });

        });

        return preFetchdeferred.promise();
    }

    this.fetchTestResultAttatchements=function(suite, tpList){
        var deferredAttatchments = $.Deferred();

        var tpCount= tpList.length;

        var testResultPromises = [];
        var tpResultList = []
        var srv = this;
        if (tpList.length > 0 && FEATURE_TestResultAttatchments) {
            tpList.forEach(function (tp) {
                var testRunId = Number(tp.lastTestRun.id)
                if (testRunId != 0) {
                    var testResultId = Number(tp.lastResult.id)
                    testResultPromises.push(srv.client.getTestResultById(suite.project.name, testRunId, testResultId, TestContracts.ResultDetails.Iterations));
                    tpResultList.push(tp);
                }
            });

            Q.allSettled(testResultPromises).then(

                function (results) {
                    for (var i = 0; i < results.length; i++) {
                        if (results[i].state === "fulfilled") {
                            var tr = results[i].value;
                            if (tr != null) {
                                if (tr.iterationDetails != null && tr.iterationDetails.length > 0 && tr.iterationDetails[0] != null) {
                                    tpResultList[i]["_resultAttatchments"] = tr.iterationDetails[0].attachments;
                                }
                            }
                        }
                    }
                    deferredAttatchments.resolve(tpList);
                },
                function (err) {
                    deferredAttatchments.resolve(tpList);
                }
            );
               
        }
        else {
            deferredAttatchments.resolve(tpList);
        }
        return deferredAttatchments.promise();
    }



    this.fetchTestResults = function (client, witClient, suitelist)
    {
        var preFetchdeferred = $.Deferred();
        var idList = [];
        var suiteFetchedCount = suitelist.length;
        var srv = this;
        srv.client = client;
        var testPointPromises = [];
        //var suiteResultList = []

        suitelist.forEach(function (suite) {
            if (srv.ProgressMessage("Fetching  test results " + suiteFetchedCount)) {
                testPointPromises.push(srv.client.getPoints(suite.project.name, suite.plan.id, suite.id));
            }
            else {
                preFetchdeferred.reject("user canceled");
            }
        });
        
        Q.all(testPointPromises).then(
            function (tpdata) {
                for (var i = 0; i < tpdata.length; i++) {
                    var tpList = tpdata[i];
                    if (srv.ProgressMessage("Fetching test results, recieved " + i + " of " + suitelist.length)) {

                        suitelist[i]._suiteTestPoints = tpList
                        
                        srv.fetchTestResultAttatchements(suitelist[i], tpList).then(
                            function (tp) {
                                suiteFetchedCount--
                                if (suiteFetchedCount == 0) {
                                    preFetchdeferred.resolve(null);
                                }
                            },
                            function (err) {
                                suiteFetchedCount--
                                if (suiteFetchedCount == 0) {
                                    preFetchdeferred.resolve(null);
                                }
                            }
                        );
                    }
                 
                    else {
                        preFetchdeferred.reject("user canceled");
                    }
                }
            },
            function (err) {
                preFetchdeferred.reject(err);
            });

        return preFetchdeferred.promise();
    }

    this.mapTestResults = function (suitelist)
    {
        
    }

});

function AddTestPlan2Xml(rootNode, testPlan) {
    var testPlanTag = appendTags(rootNode, "testPlan");
    testPlanTag.setAttribute("id", testPlan.id.toString());
    testPlanTag.setAttribute("title", testPlan.name);
    testPlanTag.setAttribute('url', getTestPlanUrl(testPlan));

    // Test plan properties 
    VSS.require(["VSS/Utils/Date"], function (Utils_Date) {
        var propertiesTag = appendTags(testPlanTag, "properties");
        addPropertyTag(propertiesTag, 'AreaPath', testPlan.area.name);
        addPropertyTag(propertiesTag, 'Iteration', testPlan.iteration);
        addPropertyTag(propertiesTag, 'Owner', testPlan.owner.displayName);
        addPropertyTag(propertiesTag, 'State', testPlan.state);
        addPropertyTag(propertiesTag, 'StartDate', Utils_Date.localeFormat(testPlan.startDate, "D", true));
        addPropertyTag(propertiesTag, 'EndDate', Utils_Date.localeFormat(testPlan.endDate, "D", true));
    
        //Test Plan Desc 
        if (testPlan.description) {
            //TODO 
            var desc = HtmlNormalizer.normalize(this._selectedPlan.description);
            desc = desc.replace(/\n/g, "<br />");
            addRitchTextTag(rootNode, 'description', TestsOM.HtmlUtils.wrapInDiv(desc))
        }

        // TEst Plan settings
        var testPlanSettingsTag = appendTags(rootNode, "testPlanSettings");

        if (testPlan.manualTestSettings != null) {
            addTestPlanRun(testPlanSettingsTag, 'manualRuns', testPlan.manualTestSettings.testSettingsName, testPlan.manualTestEnvironment.environmentName);
        }
        if (testPlan.automatedTestEnvironment != null) {
            addTestPlanRun(testPlanSettingsTag, 'automatedRuns', testPlan.automatedTestEnvironment.testSettingsName, testPlan.automatedTestEnvironment.environmentName);
        }

        //TODO addBuilds($testPlanSettingsTag);

        //addConfigs($rootDom, callback, errorCallback);
    });
    return testPlanTag;
}

function addTestPlanRun($rootDom, tagName, testSettings, testEnv) {
    var $manualRunsTag = appendTags($rootDom, tagName);
    var $propertiesTag = appendTags($manualRunsTag, "properties");
    var settingsName = "";//Resources.DisplayTextNone, 
    var envName = ""; //Resources.DisplayTextNone;
    if (testSettings) {
        settingsName = testSettings;
    }
    if (testPlan.manualTestEnvironment) {
        envName = testEnv;
    }

    this.addPropertyTag($propertiesTag, "Settings", settingsName);
    this.addPropertyTag($propertiesTag, "Environment", envName);
}

function addBuilds ($rootDom, buildInfo) {
    var $buildsTag = appendTags($rootDom, "builds");
    var $propertiesTag = appendTags($buildsTag, "properties")
    var buildDef = ""; // Resources.DisplayTextNone;
    var buildQuality = ""; // Resources.DisplayTextNone;
    var buildNumber = ""; // Resources.DisplayTextNone;
    if(buildInfo) {
        if (buildInfo.definition) {
            buildDef = buildInfo.definition.name;
        }
        if (buildInfo.quality) {
            buildQuality = buildInfo.quality;
        }
        if (buildInfo.buildNumber) {
            buildNumber =buildInfo.buildNumber;
        }
    }
    addPropertyTag($propertiesTag, "Definition", buildDef);
    addPropertyTag($propertiesTag, "Quality", buildQuality);
    addPropertyTag($propertiesTag, "BuildInUse", buildNumber);
};	

function addConfigs($rootDom, testPlan, testSuite) {
    var configMap = {};
    var $configurations,$planConfiguration, $addtionalConfiguration, $config;
    var allconfigIds, additionalConfigurationIds, rootSuiteConfigIds;

            
    //this._testPlanManager.getTestSuite(this._planId, this._rootSuiteId, function 
    //    (testSuite) {
    //        additionalConfigurationIds = _this._getAdditionalConfigurationIds(testSuite.defaultConfigurations);
    //        rootSuiteConfigIds = $.map(testSuite.defaultConfigurations, function (item, index) {
    //            return parseInt(item.id, 10);
    //        });
    //        allconfigIds = Utils_Array.unique(additionalConfigurationIds.concat(rootSuiteConfigIds));
    //        Diag.logTracePoint("HtmlDocumentGenerator._addCongfigurations.getTestConfigurationsDetail.Start");
    //        _this._testPlanManager.getTestConfigurationsDetail(allconfigIds, _this._planId, function (configDetails) {
    //            Diag.logTracePoint("HtmlDocumentGenerator._addCongfigurations.getTestConfigurationsDetail.End");
    //            for (i = 0; i < configDetails.length; i++) {
    //                configMap[configDetails[i].id] = configDetails[i];
    //            }
    //            $configurations = _this.appendTags($rootDom, "configurations");
    //            $planConfiguration = _this.appendTags($configurations, "planConfiguration");
    //            for (i = 0; i < testSuite.defaultConfigurations.length; i++) {
    //                _this._appendConfigurationsData($planConfiguration, configMap[testSuite.defaultConfigurations[i].id]);
    //            }
    //            if (additionalConfigurationIds.length > 0) {
    //                $addtionalConfiguration = _this.appendTags($configurations, "additionalConfiguration");
    //                for (i = 0; i < additionalConfigurationIds.length; i++) {
    //                    _this._appendConfigurationsData($addtionalConfiguration, configMap[additionalConfigurationIds[i]]);
    //                }
    //            }
    //            callback();
    //        }, errorCallback);
        
}

function AddFlatSuiteList2PlanXml(srv, node)
{

    var suitesList = appendTags(node, 'testSuites');
    srv.suiteList.forEach(function (suite) {

        // Add to flat list of tests suites
        var suiteNode = appendTags(suitesList, 'testSuite');
        suiteNode.setAttribute('id', suite.id);
        suiteNode.setAttribute('title', suite.id + ': ' + suite.name);
        suiteNode.setAttribute('url', getTestSuiteUrl(suite));

        var suiteNodeProperties = appendTags(suiteNode, 'suiteProperties');
        var properties = appendTags(suiteNodeProperties, 'properties');
        addPropertyTag(properties, 'State', suite.state);
        addPropertyTag(properties, 'Type', suite.suiteType);


        addConfigs2SuiteXML(suiteNode, suite);

        AddTCestCases2SuiteXml(srv, suiteNode, suite);
    });
}

function GetOutcomeCnt(tpList)
{
    var cnt = {
        active:         tpList.filter(function (i) { return i.outcome == 'Unspecified' }).length,
        notApplicable:  tpList.filter(function (i) { return i.outcome == 'NotApplicable' }).length,
        passed:         tpList.filter(function (i) { return i.outcome == 'Passed' }).length,
        failed:         tpList.filter(function (i) { return i.outcome == 'Failed' }).length,
        blocked:        tpList.filter(function (i) { return i.outcome == 'Blocked' }).length,
        paused:         tpList.filter(function (i) { return i.outcome == 'Paused' }).length
    };
    return cnt;
}

function AddOutcomeCntAttributes(tag, cnt, prefix) {
    
    AddOutcomeCntAttribute(tag, attribName(prefix, "active"), cnt.active);
    AddOutcomeCntAttribute(tag, attribName(prefix, "notApplicable"), cnt.notApplicable);
    AddOutcomeCntAttribute(tag, attribName(prefix, "passed"), cnt.passed);
    AddOutcomeCntAttribute(tag, attribName(prefix, "failed"), cnt.failed);
    AddOutcomeCntAttribute(tag, attribName(prefix, "blocked"), cnt.blocked);
    AddOutcomeCntAttribute(tag, attribName(prefix,"paused"), cnt.paused);
}

function AddOutcomeCntAttribute(tag, name, value) {
    if(value!=0){
        tag.setAttribute(name, value)
    }
}
function attribName(prefix,name)
{
    if (prefix.length == 0) {
        return name;
    }
    else {
        return prefix + name.charAt(0).toUpperCase() + name.slice(1)
    }
}

function setSuitePath(suiteList, suite) {
    if(suite==null){
        //Find root 
        suite = suiteList.filter(function (o) { return o.parent == null; })[0];
        suite.path = "\\";
        setSuitePath(suiteList, suite);
    }
    else{
        suiteList.filter(function (o) { if (o.parent == null) { return false } else { return o.parent.id == suite.id; } }).forEach(function (s) {
            s.path = suite.path + (suite.path == "\\"?"":"\\") + s.name;
            setSuitePath(suiteList, s);
        });
    }
}


function AddSuiteHierarchy2PlanXml(srv, parent, suite) {

    var cnt = GetOutcomeCnt(suite._suiteTestPoints);
    if (parent.path == null) {
        parent.path = "";
        }

    suite.path = parent.path + "\\" + suite.name;

    var suiteTag = appendTags(parent, "suite");
    suiteTag.setAttribute("id", suite.id);
    suiteTag.setAttribute("title", suite.name);
    suiteTag.setAttribute("type", suite.suiteType);
    

    AddOutcomeCntAttributes(suiteTag, cnt, "");
        
    suiteTag.setAttribute("url", getTestSuiteUrl(suite));

    srv.suiteList.filter(function (o) { if (o.parent == null) { return false } else { return o.parent.id == suite.id; } }).forEach(function (s) {
        if (suite.testSuites == null) {
            suite.testSuites = [];
    }
        suite.testSuites.push(s);

        //add to suiteHierachy
        var cntSuite = AddSuiteHierarchy2PlanXml(srv, suiteTag, s);
        //Add up summs
        cnt.active+= cntSuite.active;
        cnt.notApplicable+= cntSuite.notApplicable;
        cnt.passed+= cntSuite.passed;
        cnt.failed+= cntSuite.failed;
        cnt.blocked+= cntSuite.blocked;
        cnt.paused+= cntSuite.paused;      
    });

  
    AddOutcomeCntAttributes(suiteTag, cnt, "total");
    
    return cnt;
    }


        function addConfigs2SuiteXML(planNode, suite) {
    if (suite.configs != null) {
        //TODO:CHEC IF PLAN NODE (suitenode)
        var suiteNodeConfigs = appendTags(planNode, 'configurations');

        suite.configurations.forEach(function (c) {
            var $config = appendTags(suiteNodeConfigs, 'configuration');
            $config.setAttribute('value', c.name);

        });
        }
    }

        function AddTCestCases2SuiteXml(srv, $suiteNode, suite)
        {
    var testCasesList = suite._suiteTestCaseList;
    var tcPath = "";
    
    if (testCasesList.length != null) {
        var $testCasesNode = appendTags($suiteNode, 'testCases');
        $testCasesNode.setAttribute('count', testCasesList.length);

        testCasesList.forEach(function (tc) {

            var tcWi = srv.testCaseList.filter(function (o) { return o.id == tc.testCase.id })[0];
            // add test case 

            var $tc = appendTags($testCasesNode, 'testCase');
            $tc.setAttribute('id', tc.testCase.id);
            $tc.setAttribute('url', getTestCaseUrl(tcWi));
            $tc.setAttribute('title', tcWi.fields['System.Title']);
            $tc.setAttribute('suitePath', suite.path);

            tcPath = suite.path+ "\\" + tc.testCase.id + "_" + tcWi.fields['System.Title'];

            var fieldsTag = appendTags($tc, "fields");
            for (var f in tcWi.fields) {
                if (tcWi.fields.hasOwnProperty(f)) {

                    switch (f) {
                        case "Microsoft.VSTS.TCM.Steps'":
                            break;
                        case "System.Description":
                            addRitchTextTag($tc, 'summary', tcWi.fields[f]);
                            break;
                        default:
                            addPropertyTag(fieldsTag, f, tcWi.fields[f]);
                            break;
                }

            }
        }


            
            var stepsXml = tcWi.fields['Microsoft.VSTS.TCM.Steps'];
            var steps = parseTestSteps(stepsXml);
            
            var $latestTestOutcomes = appendTags($tc, "latestTestOutcomes")
            var tps = suite._suiteTestPoints.filter(function (i) { return i.testCase.id == tc.testCase.id; });
           
            AddOutcomeCntAttributes($tc,  GetOutcomeCnt(tps), "");

            tps.forEach(function (tp) {
                var $testResult = appendTags($latestTestOutcomes, "testResult")
                var $testResultProperties = appendTags($testResult, "properties")

                //addPropertyTag($testResultProperties, "suitePath", suite.path);
                testPointColumns.forEach(function (p) {
                    addPropertyTag($testResultProperties, p.name, getNestedPropValue(p.property, tp))
                });
              


                if (tp["_resultAttatchments"] != null) {
                    var $testResultAttachemnts = appendTags($testResult, "testResultAttatchments");

                    tp["_resultAttatchments"].forEach(function (r) {
                        var $trAtt = appendTags($testResultAttachemnts, "attatchment");
                        $trAtt.setAttribute('id', r.id);
                        $trAtt.setAttribute('path', tcPath + "\\" + tp.configuration.name );
                        $trAtt.setAttribute('fileName', r.name);
                        $trAtt.setAttribute('url', DownloadArtifactUrl.replace("{id}", r.id));
                        
                    });
            }

            });


            var $stepList = appendTags($tc, "testSteps");

            steps.forEach(function (step) {
                var $step = appendTags($stepList, 'testStep');
                $step.setAttribute("index", step.id);

                addRitchTextTag($step, 'testStepAction', step.action)

                addRitchTextTag($step, 'testStepExpected', step.expectedResult);
             

                var $attatchment= appendTags($step, 'testStepAttachments');
                
            });
        })
        }
    }

        function getNestedPropValue(name, object) {
    if (object == null) {
        return "";
    }
    else {
        if (name.indexOf(".") >= 0) {
            var n = name.split(".")[0];
            name = name.replace(n + ".", "");
            return getNestedPropValue(name, object[n]);
        }
        else {
            var value = object[name];
            if (name == "displayName") {
                if(value!=null && value.indexOf("<")>=0){
                    value=value.split("<")[0];
            }
        }
            if (name == "dateCompleted") {
                value = Utils_Date.localeFormat(value, "D");
        }
            return value == null ? "" : value;
    }
        }
    }

        function parseTestSteps(stepXml) {
    $stepXmlDom = $.parseXML(stepXml);
    var steps = [];
    $($stepXmlDom).find("step").each(function (i, step) {
        steps.push(readStep(step));
    });

    return steps;
    }

        function readParameterizedString(parameterizedString) {
    var stringParts = [];
    if ($(parameterizedString).children().length > 0) {
        $(parameterizedString).children().each(function () {
            switch (this.nodeName.toLowerCase()) {
                case "parameter":
                    stringParts.push(Utils_String.format("@{0}", $(this).text()));
                    break;
                case "outputparameter":
                    stringParts.push(Utils_String.format("@?{0}", $(this).text()));
                    break;
                case "text":
                    stringParts.push($(this).text());
                    break;
        }
        });
    }
    else {
        stringParts.push($(parameterizedString).text());
        }
    return stringParts.join("");
    }

        function readStep(step) {
    var id = $(step).attr("id"), stepType = $(step).attr("type"), action, expectedResult, count = 0, testStep, isFormatted = false, isActionFormatted = false, isExpectedResultFormatted = false;

    $(step).children("parameterizedString").each(function () {
        if (count === 0) {
            action = readParameterizedString(this);
            isActionFormatted = ($(this).attr("isformatted") === "true");
        }
        else {
            expectedResult = readParameterizedString(this);
            isExpectedResultFormatted = ($(this).attr("isformatted") === "true");
    }
        count++;
    });
    if (isActionFormatted && isExpectedResultFormatted) {
        isFormatted = true;
        }
    if (!action && !expectedResult) {


        $(step).children().each(function () {
            switch (this.nodeName.toLowerCase()) {
                case "action":
                    action = $(this).text();
                    break;
                case "expected":
                    expectedResult = $(this).text();
                    break;
        }
        });
        }

    testStep = { 
            id:parseInt(id, 10), 
            stepType: stepType,             
            action: action, 
            expectedResult: expectedResult, 
            isFormatted: isFormatted
        };

    return testStep;
    }

        function addPropertyTag(node, name, value) {
    var tagPropery = node.ownerDocument.createElement("property");
    try {
        tagPropery.setAttribute("name", name);
        tagPropery.setAttribute("value", value);
        node.appendChild(tagPropery);
        return tagPropery;
        } catch (ex) {
        return null;
        }
    }

        function appendTags(node, tagName) {
    try {
        var tag = node.ownerDocument.createElement(tagName);
        node.appendChild(tag);
        return tag;
        } catch (ex) {
        return null;
        }
    }

        function addRitchTextTag(node, tagname, rithcText) {
    var tag = appendTags(node, tagname);
            //$expected.append($(TestsOM.HtmlUtils.wrapInDiv(HtmlNormalizer.normalize(step.expectedResult))));
    try {
        var importXml= $.parseXML(wrapInDiv(secureHTMLisXml(rithcText)));

        var nodeTxt = tag.ownerDocument.importNode(importXml.documentElement, true);
        tag.appendChild(nodeTxt);
        }
    catch (ex) {

        }
    }


        function getTestPlanUrl(plan) {
    var url = getTeamProjectUrl();
    url += "_testManagement#planId=" + plan.id;
    url += "&ampsuiteId=" + plan.rootSuite.id;
    return url;
    }

        function getTestSuiteUrl(suite) {
    var url = getTeamProjectUrl();
    url += "_testManagement#planId=" + suite.plan.id;
    url += "&ampsuiteId="+suite.id;
        
    return url;
    }
        function getTestCaseUrl(tc) {
    var url = getTeamProjectUrl();
    url += "_workitems/edit/" + tc.id;
    return url;  
    }
        function getTeamProjectUrl() {
    return VSS.getWebContext().collection.uri + VSS.getWebContext().project.name + "/"
    }

