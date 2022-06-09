//---------------------------------------------------------------------
// <copyright file="query-fetcher">
//    This code is licensed under the MIT License.
//    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF 
//    ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED 
//    TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
//    PARTICULAR PURPOSE AND NONINFRINGEMENT.
// </copyright>
// <summary>A utility to execute a query, fetch and build all work item data for the query result.</summary>
//---------------------------------------------------------------------

       

    function QueryFetcher() {
    
        this.executeQuery = function (queryId, progressCallback) {

            this.progressCallback = progressCallback;

            var deferred = $.Deferred();

            VSS.require(["VSS/Controls", "VSS/Service", "TFS/WorkItemTracking/RestClient",
           "TFS/WorkItemTracking/Contracts", "VSS/Context"],
           function (Controls, VSS_Service, TFS_Wit_WebApi, TFS_Wit_Contract, Context) {

               var witClient = VSS_Service.getCollectionClient(TFS_Wit_WebApi.WorkItemTrackingHttpClient);
               this.ProgressMessage("Executing query");
               var ctx = VSS.getWebContext();
               var team = null //
               if (Context.getPageContext().webAccessConfiguration.isHosted) {
                   team = ctx.team.name
               }

               witClient.queryById(queryId, ctx.project.name, team).then(
                    function (queryResult) {
                        var queryFetcher = this;
                        queryFetcher.queryResult = queryResult;

                        var wi = [];
                        var ids = [];
                        if (queryResult.queryResultType == 1) {
                            ids = queryResult.workItems.map(function (reference) { return reference.id; });
                        }
                        if (queryResult.queryResultType == 2) {
                            ids = queryResult.workItemRelations.map(function (reference) { return reference.target.id; });
                        }
                
                        var bulkFetcher = new BulkWIFetcher()
                        bulkFetcher.FetchAllWorkItems( witClient, ids, queryResult.columns.map(function (reference) { return reference.referenceName; }), TFS_Wit_Contract.WorkItemExpand.Fields)
                            .then(
                            function (queryData) {
                                queryFetcher.queryData = queryData;
                                ProgressMessage("Retrieved data, building hierarchy");
                                if (queryFetcher.queryResult.queryResultType == 2) {
                                    queryFetcher.treeRoot = [];
                                    var rootItems = queryFetcher.queryResult.workItemRelations.filter(function (o) { return o.source == undefined });

                                    rootItems.forEach(function (n) {
                                        var workItem = queryFetcher.queryData.filter(function (wi) { return wi.id == n.target.id })[0];
                                        queryFetcher.treeRoot.push(workItem);
                                        try{
                                            queryFetcher.BuildTree(queryFetcher, workItem);
                                        }
                                        catch (errBuildTree) {
                                            deferred.reject(errBuildTree);
                                        }
                                    
                                    });
                                    queryFetcher.queryResultList = queryFetcher.treeRoot;
                
                                }
                                if (queryFetcher.queryResult.queryResultType == 1) {
                                    queryFetcher.queryResultList = queryResult.workItems.map(function(i) {
                                        return queryFetcher.queryData.filter(function(d){
                                            return d.id == i.id
                                        })[0];
                                    });
                                }
                                deferred.resolve({ Columns: queryFetcher.queryResult.columns, QueryResults: queryFetcher.queryResultList });
                            }, 
                            function (err) {
                                deferred.reject(err);
                            });
                    }, 
                    function (errGetById) {
                        deferred.reject(errGetById);
                    });
            });
        
            return deferred.promise()
        }

  
        this.BuildTree = function (ctrl, wi) {
            ctrl.queryResult.workItemRelations.filter(function (o) {
                if (o.source == null)
                    return false;
                return o.source.id == wi.id;
            })
            .forEach(function (item) {
                if (wi.childs == null) {
                    wi.childs = [];
                }
                var workItem = ctrl.queryData.filter(function (wi) { return wi.id == item.target.id })[0];

                wi.childs.push(workItem);
                ctrl.BuildTree(ctrl, workItem);
            });
        }

        this.ProgressMessage = function (message)
        {
            if (this.progressCallback != null)
            {
                this.progressCallback(message);
            }
        }

        return this;

    }

    function BulkWIFetcher() {
        this.QueryData = [];

        this.FetchAllWorkItems = function ( witClient, wiIdLst, fields, expand) {
            var fetchWorkItemsDeferred= new $.Deferred();
            var reqIdLst= [];
            var size = 100;
            var asOfDate = new Date();

        
            wiIdLst = wiIdLst.unique();
            this.wiDataCount = wiIdLst.length;

            this.ProgressMessage("Fetching work item data <br/>Fetched 0 out of " + this.wiDataCount);

            var bulkFetcher = this;

  
            while (wiIdLst.length > 0) {
                reqIdLst = wiIdLst.splice(0, size);
                witClient.getWorkItems(reqIdLst, fields).then(function (workItems) {
                    var d = bulkFetcher.AddWorkItemData(bulkFetcher, workItems);
                    if (d != null) {
                        fetchWorkItemsDeferred.resolve(d);
                    }

                });
            }

            return fetchWorkItemsDeferred.promise();
        }


        this.AddWorkItemData = function (bulkFetcher, workItems) {

            bulkFetcher.QueryData = this.QueryData.concat(workItems);
            this.ProgressMessage("Fetching work item data <br/>Fetched " + this.QueryData.length + " out of " + this.wiDataCount);

            if (bulkFetcher.QueryData.length == this.wiDataCount) {
                return bulkFetcher.QueryData;
            }
            else {
                return null;
            }
        }

        this.ProgressMessage = function (message) {
            if (this.progressCallback != null) {
                this.progressCallback(message);
            }
        }

        return this;
    }
