function callWebHookAction(actionName, requestUrl, reloadAfterExecute) {
        var serverURL = Xrm.Page.context.getClientUrl();
        var req = new XMLHttpRequest();
        var objectId = '';
        if (Xrm.Page != null && Xrm.Page.data != null && Xrm.Page.data.entity != null) {
            objectId = Xrm.Page.data.entity.getId().replace("{", "").replace("}", "");
        }

        var data = {
            "Request": window.JSON.stringify({
                "EntityId": objectId
            }),
            "RequestUrl": requestUrl
        }

        // specify name of the entity, record id and name of the action in the Wen API Url
        req.open("POST", serverURL + "/api/data/v9.0/" + actionName, true);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.onreadystatechange = function () {
            if (this.readyState == 4 /* complete */) {
                req.onreadystatechange = null;
                if (this.status == 200 && this.response != null) {
                    var responseObject = JSON.parse(this.response);
                    console.trace(this.response);
                    if (reloadAfterExecute) {
                        Xrm.Utility.openEntityForm(Xrm.Page.data.entity.getEntityName(), Xrm.Page.data.entity.getId());
                    }
                    //callback(responseObject);
                }
                else {
                    if (this.response != null && this.response != "") {
                        var error = JSON.parse(this.response).error;
                        console.trace(error.message);
                        //errHandler(error.message);
                    }
                }
            }
        };
        // send the request with the data for the input parameter
        req.send(window.JSON.stringify(data));
    }
