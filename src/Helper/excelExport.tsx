import * as Excel from 'exceljs';
import React from 'react';
import AuthorizationClient from '../AuthorizationClient';

export class ExcelExporter{

    public static async exportIssuesToExcel(formTemplateId:string, projectId:string, formType:string, exportComments:boolean, statusUpdates:any, logger:any, verboseLogging:boolean){
        // get all the issues related to the formType
        // this is quicker than looping through every issue on the project
        // as you cannot filter directly on the "template" id used :(
        const formTypeFormsMatch:any[] = [];
        const loggingData:JSX.Element[] = [];
        var formTypeLooper = true; //used to force the initial call;
        statusUpdates("Grabbing up to date token");
        loggingData.push(<div><h1>Logger results</h1></div>);
        loggingData.push(<div>Grabbing up to date token</div>);
        const accessToken = await (await AuthorizationClient.oidcClient.getAccessToken()).toTokenString();
        //const accessTokenTime = await (await AuthorizationClient.oidcClient.getAccessToken()).getStartsAt();
        if(verboseLogging){console.log(`Got accesstoken ${accessToken}`)};
        var urlToQuery : string = `https://api.bentley.com/issues/?top=50&projectId=${projectId}&type=${formType}`;
        while (formTypeLooper) {
            statusUpdates("Looking through issues...");
            loggingData.push(<div>Looking through issues...</div>);
            if(verboseLogging){console.log(`Looking at ${urlToQuery}`)};
            const response = await fetch(urlToQuery, {
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': accessToken,
                  },
            })
            const data = await response;
            const json = await data.json();
            if(verboseLogging){console.log("Got data ",json)};

            if (data.status === 500){
                //api failure somewhere!
                statusUpdates(`Bentley had a server failure. Try again later.`);
                loggingData.push(<div className="errRow">Bentley had a server failure. Try again later.</div>);
                logger(loggingData);
                return;
            }
            json.issues.forEach((id: string) => {
                formTypeFormsMatch.push(id);
            });
            //let see if we are continuing.
            try {
                formTypeLooper = true;
                urlToQuery = json._links.next.href;
                if(verboseLogging){console.log(`Continuation found new query is ${urlToQuery}`)};
            } catch (error) {
                // better than === undefined?
                //swallow the missing link error and stop the loop
                formTypeLooper = false;
            }
        }

        if (!(formTypeFormsMatch.length > 0))
        {
            statusUpdates("No exportable form instances found");
            loggingData.push(<div className="errRow">No exportable form instances found</div>);
            if(verboseLogging){console.error(`No form instances were found`)};
            alert("No exportable form instances found");
            return;
        }
        
        //continuing on create template for excel.
        // red columns do not go back into the system during import.
        // yellow columns are mandatory for import
        // fixed header structure -
        statusUpdates("Building excel template...");
        loggingData.push(<div>Building excel template...</div>);
        const wb = new Excel.Workbook();
        var ws = wb.addWorksheet("IRS");
        var currCol = 1;
        ws.getCell(1, currCol++).value = "id"; const xlColid = currCol-1;
        ws.getCell(1, currCol-1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFFFF00'} };; //want yellow
        ws.getCell(1, currCol++).value = "number"; const xlColnumber = currCol-1;
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid',fgColor:{argb:'FFFFFF00'}};; //want yellow
        ws.getCell(1, currCol++).value = "displayName"; const xlColvalue = currCol-1;
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red
        ws.getCell(1, currCol++).value = "subject"; const xlColsubject = currCol-1;
        ws.getCell(1, currCol++).value = "description"; const xlColdescription = currCol-1;
        ws.getCell(1, currCol++).value = "location.latitude"; const xlCollocationlatitude = currCol-1;
        ws.getCell(1, currCol++).value = "location.longitude"; const xlCollocationlogitude = currCol-1;

        ws.getCell(1, currCol++).value = "modelPin.location.x"; const xlColmodelpinglocationx = currCol-1;
        ws.getCell(1, currCol++).value = "modelPin.location.y"; const xlColmodelpinglocationy = currCol-1;
        ws.getCell(1, currCol++).value = "modelPin.location.z"; const xlColmodelpinglocationz = currCol-1;

        ws.getCell(1, currCol++).value = "createdBy"; const xlColcreatedBy = currCol-1;
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red
        ws.getCell(1, currCol++).value = "createdDateTime"; const xlColcreatedDateTime = currCol-1;
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red
        ws.getCell(1, currCol++).value = "status"; const xlColstatus = currCol-1;
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red
        ws.getCell(1, currCol++).value = "assignee.displayName"; const xlColassigneedisplayName = currCol-1;
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red
        ws.getCell(1, currCol++).value = "assignee.email"; const xlColassigneeemail = currCol-1;
        ws.getCell(1, currCol++).value = "assignee.id"; const xlColassigneeid = currCol-1;
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red
        ws.getCell(1, currCol++).value = "dueDate"; const xlColdueDate = currCol-1;
        ws.getCell(1, currCol++).value = "state"; const xlColstate = currCol-1;
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red
        ws.getCell(1, currCol++).value = "assignees"; const xlColassignees = currCol-1;
        ws.getCell(1, currCol++).value = "formGUID"; const xlColformGUID = currCol-1;
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red

        var userPropsStartAtCol = currCol;
        var currRow = 1;
        var gotData = false;
        if(verboseLogging){console.log(`Created all of the default excel columns`)};

        // loop through all the potential matches and extract the properties column headers from the data
        // this is the only way you can actually figure out what controls are on the form
        // and the access keys to formulate for an upload.... very costly! :(
        for(const typeMatchId of formTypeFormsMatch) {
            if(verboseLogging){console.log("Examining form instance for match ", typeMatchId)};
            statusUpdates(`Checking form ${typeMatchId.id}`);
            loggingData.push(<div>Checking form {typeMatchId.id}</div>);
            const response = await fetch(`https://api.bentley.com/issues/${typeMatchId.id}`, {
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': accessToken,
                  },
            })
            const data = await response;
            const json = await data.json();
            if (data.status === 500){
                //api failure somewhere!
                statusUpdates(`Bentley had a server failure. Try again later.`);
                loggingData.push(<div className="errRow">Bentley had a server failure. Try again later.</div>);
                logger(loggingData);
                return;
            }
            if (formTemplateId === json.issue.formId)
            {
                if(verboseLogging){console.log(`This instance is from the selected definition`)};

                currRow++;
                statusUpdates(`Exporting form to excel row: ${currRow}`);
                loggingData.push(<div>Exporting form to excel row: {currRow}</div>);
                gotData = true;
                if ("id" in json.issue){
                    ws.getCell(currRow, xlColid).value = json.issue.id;
                    if (ws.getCell(1, xlColid).note === undefined)
                    {
                        ws.getCell(1, xlColid).note = typeof json.issue.id;
                    }
                }

                if ("number" in json.issue){
                    ws.getCell(currRow, xlColnumber).value = json.issue.number;
                    if (ws.getCell(1, xlColnumber).note === undefined)
                    {
                        ws.getCell(1, xlColnumber).note = typeof json.issue.number;
                    }
                }

                if ("displayName" in json.issue){
                    ws.getCell(currRow, xlColassigneedisplayName).value = json.issue.displayName;
                    if (ws.getCell(1, xlColassigneedisplayName).note === undefined)
                    {
                        ws.getCell(1, xlColassigneedisplayName).note = typeof json.issue.displayName;
                    }
                }

                if ("subject" in json.issue){
                    ws.getCell(currRow, xlColsubject).value = json.issue.subject;
                    if (ws.getCell(1, xlColsubject).note === undefined)
                    {
                        ws.getCell(1, xlColsubject).note = typeof json.issue.subject;
                    }
                }

                if ("description" in json.issue){
                    ws.getCell(currRow, xlColdescription).value = json.issue.description;
                    if (ws.getCell(1, xlColdescription).note === undefined)
                    {
                        ws.getCell(1, xlColdescription).note = typeof json.issue.description;
                    }
                }

                if ("location" in json.issue){
                    if ("latitude" in json.issue.location){
                        ws.getCell(currRow, xlCollocationlatitude).value = json.issue.location.latitude;
                        if (ws.getCell(1, xlCollocationlatitude).note === undefined)
                        {
                            ws.getCell(1, xlCollocationlatitude).note = typeof json.issue.location.latitude;
                        }
                    }

                    if ("longitude" in json.issue.location){
                        ws.getCell(currRow, xlCollocationlogitude).value = json.issue.location.longitude;
                        if (ws.getCell(1, xlCollocationlogitude).note === undefined)
                        {
                            ws.getCell(1, xlCollocationlogitude).note = typeof json.issue.location.longitude;
                        }
                    }
                }

                if ("modelPin" in json.issue){
                    if ("location" in json.issue.modelPin){
                        if ("x" in json.issue.modelPin.location){
                            ws.getCell(currRow, xlColmodelpinglocationx).value = json.issue.modelPin.location.x;
                            if (ws.getCell(1, xlColmodelpinglocationx).note === undefined)
                            {
                                ws.getCell(1, xlColmodelpinglocationx).note = typeof json.issue.modelPin.location.x;
                            }
                        }
                        if ("y" in json.issue.modelPin.location){
                            ws.getCell(currRow, xlColmodelpinglocationy).value = json.issue.modelPin.location.y;
                            if (ws.getCell(1, xlColmodelpinglocationy).note === undefined)
                            {
                                ws.getCell(1, xlColmodelpinglocationy).note = typeof json.issue.modelPin.location.y;
                            }
                        }
                        if ("z" in json.issue.modelPin.location){
                            ws.getCell(currRow, xlColmodelpinglocationz).value = json.issue.modelPin.location.z;
                            if (ws.getCell(1, xlColmodelpinglocationz).note === undefined)
                            {
                                ws.getCell(1, xlColmodelpinglocationz).note = typeof json.issue.modelPin.location.z;
                            }
                        }
                    }
                }

                if ("createdBy" in json.issue){
                    ws.getCell(currRow, xlColcreatedBy).value = json.issue.createdBy;
                    if (ws.getCell(1, xlColcreatedBy).note === undefined)
                    {
                        ws.getCell(1, xlColcreatedBy).note = typeof json.issue.createdBy;
                    }
                }

                if ("createdDateTime" in json.issue){
                    ws.getCell(currRow, xlColcreatedDateTime).value = json.issue.createdDateTime;
                    if (ws.getCell(1, xlColcreatedDateTime).note === undefined)
                    {
                        ws.getCell(1, xlColcreatedDateTime).note = typeof json.issue.createdDateTime;
                    }
                }

                if ("status" in json.issue){
                    ws.getCell(currRow, xlColstatus).value = json.issue.status;
                    if (ws.getCell(1, xlColstatus).note === undefined)
                    {
                        ws.getCell(1, xlColstatus).note = typeof json.issue.status;
                    }
                }

                if ("assignee" in json.issue){
                    if ("displayName" in json.issue.assignee){
                        ws.getCell(currRow, xlColassigneedisplayName).value = json.issue.assignee.displayName;
                        if (ws.getCell(1, xlColassigneedisplayName).note === undefined)
                        {
                            ws.getCell(1, xlColassigneedisplayName).note = typeof json.issue.assignee.displayName;
                        }
                    }

                    // assignee gets some extra treatment...
                    if ("id" in json.issue.assignee){
                        statusUpdates(`Exporting form to excel row: ${currRow} Supplementing assignee data`);
                        loggingData.push(<div>Exporting form to excel row: {currRow} Supplementing assignee data</div>);
                        const usersEmail = await ExcelExporter.getUsersEmailFromGuid(json.issue.assignee.id, projectId,  accessToken);
                        if(verboseLogging){console.log(`Got this email for the user ${usersEmail}`)};

                        if (usersEmail === "0")
                        {
                            // it is a role
                            ws.getCell(currRow, xlColassigneeemail).value = json.issue.assignee.displayName;
                            if (ws.getCell(1, xlColassigneeemail).note === undefined)
                            {
                                ws.getCell(1, xlColassigneeemail).note = typeof json.issue.assignee.displayName;
                            }
                        }
                        else{
                            ws.getCell(currRow, xlColassigneeemail).value = usersEmail;
                        }

                        ws.getCell(currRow, xlColassigneeid).value = json.issue.assignee.id;
                        if (ws.getCell(1, xlColassigneeid).note === undefined)
                        {
                            ws.getCell(1, xlColassigneeid).note = typeof json.issue.assignee.id;
                        }
                    }
                }

                if ("dueDate" in json.issue){
                    ws.getCell(currRow, xlColdueDate).value = json.issue.dueDate;
                    if (ws.getCell(1, xlColdueDate).note === undefined)
                    {
                        ws.getCell(1, xlColdueDate).note = typeof json.issue.dueDate;
                    }
                }

                if ("state" in json.issue){
                    ws.getCell(currRow, xlColstate).value = json.issue.state;
                    if (ws.getCell(1, xlColstate).note === undefined)
                    {
                        ws.getCell(1, xlColstate).note = typeof json.issue.state;
                    }
                }
                // assignees get some extra treatment
                if ("assignees" in json.issue){
                    statusUpdates(`Exporting form to excel row: ${currRow} Supplementing assignees data`);
                    loggingData.push(<div>Exporting form to excel row: {currRow} Supplementing assignees data</div>);
                    if(verboseLogging){console.log("Looking at these assignees ", json.issue.assignees)};
                    // assignees is an array (well should be!)
                    var strAssignees = "";
                    for (var i = 0; i < json.issue.assignees.length; i++)
                    {
                        if (strAssignees===""){
                            strAssignees = json.issue.assignees[i].displayName + "|" + await ExcelExporter.getUsersEmailFromGuid(json.issue.assignees[i].id, projectId,  accessToken) + "|" + json.issue.assignees[i].id
                        }
                        else
                        {
                            strAssignees =  strAssignees + "||" + json.issue.assignees[i].displayName + "|" + await ExcelExporter.getUsersEmailFromGuid(json.issue.assignees[i].id, projectId,  accessToken) + "|" + json.issue.assignees[i].id
                        }
                    }

                    //added error handler for odd occurance where issue res has stored a null value for assignees!
                    if(strAssignees.includes("Check your issue!")){
                        if(verboseLogging){console.warn(`Handled an invalid assignee`)};
                        statusUpdates(`Exporting form to excel row: ${currRow} Invalid data found for assignees.`);
                        loggingData.push(<div className="errRow">Exporting form to excel row: {currRow} Invalid data found for assignees.</div>);
                    }
                    ws.getCell(currRow, xlColassignees).value = strAssignees;
                    if (ws.getCell(1, xlColassignees).note === undefined)
                    {
                        ws.getCell(1, xlColassignees).note = "string";
                    }
                }

                // add in formGUID
                const formGUIDObj = ExcelExporter.decodeCompositeID(ws.getCell(currRow, xlColid).value as string);
                ws.getCell(currRow, xlColformGUID).value = formGUIDObj.instanceGuid;

                // user properties. things get dynamic from here.
                if ("properties" in json.issue)
                {
                    statusUpdates(`Exporting form to excel row: ${currRow} Dumping user properties`);
                    loggingData.push(<div>Exporting form to excel row: {currRow} Dumping user properties</div>);
                    for (var property in json.issue.properties)
                    {
                        var colNum = 0;
                        for (var x = userPropsStartAtCol; x <= currCol; x++){
                            var header = property;
                            if(header.includes("__x0020__"))
                            {
                                header = header.replaceAll("__x0020__"," ");
                            }
                            if (ws.getCell(1,x).value === "properties." + header){
                                ws.getCell(currRow,x).value = json.issue.properties[property]
                                colNum = x;
                                x = currCol + 1;
                            }
                        }

                        if (colNum === 0){
                            var header = property;
                            if(header.includes("__x0020__"))
                            {
                                header = header.replaceAll("__x0020__"," ");
                            }
                            ws.getCell(1,currCol).value = "properties." + header;
                            ws.getCell(currRow, currCol).value = json.issue.properties[property];
                            if (ws.getCell(1,currCol).note === undefined){
                                ws.getCell(1,currCol).note = typeof json.issue.properties[property];
                            }
                            currCol++;
                        }


                    }

                }


            }

        }
        // double check we got some data
        if (gotData === false){
            statusUpdates("No exportable form instances were found");
            loggingData.push(<div className="errRow">No exportable form instances were found</div>);
            alert("No exportable form instances were found");
            logger(loggingData);
            if(verboseLogging){console.error(`No form instances were found`)};
            return;
        }

        // all forms have been matched and exported.
        // check to see if we need to export comments also
        if (gotData && exportComments)
        { 
            //loop through the forms in column a and get the comments from the API.
            var lastCol = currCol-1;
            var x = 2;
            while (!(ws.getCell(x,1).value === null))
            {
                statusUpdates(`Exporting form to excel row: ${x} checking for comment data`);
                loggingData.push(<div>Exporting form to excel row: {x} checking for comment data</div>);
                currCol = lastCol;
                const response = await fetch(`https://api.bentley.com/issues/${ws.getCell(x,1).value}/comments`, {
                    mode: 'cors',
                    headers: {
                        'Content-Type': 'application/json',
                        'Authorization': accessToken,
                      },
                })
                const data = await response;
                const json = await data.json();
                if (data.status === 500){
                    //api failure somewhere!
                    statusUpdates(`Bentley had a server failure. Try again later.`);
                    loggingData.push(<div className="errRow">Bentley had a server failure. Try again later.</div>);
                    logger(loggingData);
                    return;
                }
                if ("comments" in json){
                    statusUpdates(`Exporting form to excel row: ${x} adding in comment data`);
                    loggingData.push(<div>Exporting form to excel row: {x} adding in comment data</div>);
                    for (var i =0; i < json.comments.length; i++){
                        currCol++;
                        ws.getCell(x,currCol).value=json.comments[i].authorDisplayName + "|" + json.comments[i].createdDateTime + "|" + json.comments[i].text;
                    }
                }
                x++;
            }


        }

        //Lastly kick off the download
        statusUpdates("Sending you the excel result");
        loggingData.push(<div>Sending you the excel result</div>);
        logger(loggingData);
        ExcelExporter.downloadExportedIssues(wb);
        return;
    }

    public static async getUsersEmailFromGuid(userGuid:string, projectId:string, accessToken:any){
        if(userGuid === null){
            return "Check your issue!";
        }
        const response = await fetch(`https://api.bentley.com/projects/${projectId}/members/${userGuid}`, {
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': accessToken,
                  },
            })
            const data = await response;
            if (data.status === 404)
            {
                //most likely a role, or use has been ejected :(
                return "0";
            }
            
            const json = await data.json();
            if ("member" in json){
                if ("email" in json.member)
                {
                    if (json.member.email.trim() === "")
                    {
                        return ("0");
                    }
                    else{
                        return (json.member.email);
                    }
                }
            }
            else{
                return ("0");
            }

    }

    private static downloadExportedIssues(theWorkbook:Excel.Workbook){
        theWorkbook.xlsx.writeBuffer( {
            //base64: true
        })
        .then( function (xls64: BlobPart) {
            // build anchor tag and attach file (works in chrome)
            var a = document.createElement("a");
            var data = new Blob([xls64], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

            var url = URL.createObjectURL(data);
            a.href = url;
            a.download = "IssuesExporter.xlsx";
            document.body.appendChild(a);
            a.click();
            setTimeout(function() {
                    document.body.removeChild(a);
                    window.URL.revokeObjectURL(url);
                },
                0);
        })
        .catch(function(error:any ) {
            console.log(error.message);
        });
    }

    private static newGUIDfromBytesArray(guid: Uint8Array){
        var numArray = Array.prototype.slice.call(guid);
        numArray = numArray.slice(0, 4).reverse().concat(numArray.slice(4,6).reverse()).concat(numArray.slice(6,8).reverse()).concat(numArray.slice(8))
        var strGuidArray = numArray.map(function(item) {
            // return hex value with "0" padding
            return ('00'+item.toString(16).toUpperCase()).substr(-2,2);
        })

        var tmpGuid = "";
        for (var i = 0; i < 16; i++){
          tmpGuid += i === 4 || i === 6 || i === 8 || i === 10 ? "-" : "";
          tmpGuid += strGuidArray[i];
        }
        return tmpGuid.toLowerCase();
      } 

      //
      //Get the form GUID from the composite mashed ID
      // returns object .projectGUID and
      //                .instanceGUID
      public static decodeCompositeID(compositeId:string){
        var compositeBase64 = compositeId.replaceAll("-", "+");
        compositeBase64 = compositeBase64.replaceAll("_","/");
        switch (compositeBase64.length % 4) {
          case 2:
            compositeBase64 += "=="; 
            break;
          case 3: 
            compositeBase64 += "="; 
            break;
        }
        const data = atob(decodeURIComponent(compositeBase64));
        const compositeBytes = Uint8Array.from(data, b => b.charCodeAt(0));
        const projectGUID = compositeBytes.slice(0,16);
        const instanceGUID = compositeBytes.slice(16);
        var obj = {
          "projectGuid" : this.newGUIDfromBytesArray(projectGUID),
          "instanceGuid" : this.newGUIDfromBytesArray(instanceGUID)
        }
        return obj;
      }

}