import * as Excel from 'exceljs';
import { saveAs } from 'file-saver'
import { url } from 'inspector';
import { IMSLoginProper } from './IMSLoginProper';
import React from 'react';

export class ExcelExporter{

    public static tester(){
        const wb = new Excel.Workbook();
        var ws = wb.addWorksheet("IRS");
        console.log(ws.getCell(1,1).value);
        return;
    }

    public static async exportIssuesToExcel(formTemplateId:string, projectId:string, formType:string, exportComments:boolean, statusUpdates:any, logger:any){
        // get all the issues related to the formType
        // this is quicker than looping through every issue on the project
        // as you cannot filter directly on the "template" id used :(
        const formTypeFormsMatch:any[] = [];
        const loggingData:JSX.Element[] = [];
        var formTypeLooper = true; //used to force the initial call;
        statusUpdates("Grabbing up to date token");
        loggingData.push(<div><h1>Logger results</h1></div>);
        loggingData.push(<div>Grabbing up to date token</div>);
        const accessToken = await IMSLoginProper.getAccessTokenForBentleyAPI();
        var urlToQuery : string = `https://api.bentley.com/issues/?top=50&projectId=${projectId}&type=${formType}`;
        while (formTypeLooper) {
            statusUpdates("Looking through issues...");
            loggingData.push(<div>Looking through issues...</div>);
            const response = await fetch(urlToQuery, {
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': accessToken,
                  },
            })
            const data = await response;
            const json = await data.json();
            json.issues.forEach((id: string) => {
                formTypeFormsMatch.push(id);
            });
            //let see if we are continuing.
            try {
                formTypeLooper = true;
                urlToQuery = json._links.next.href;
            } catch (error) {
                // better than === undefined?
                //swallow the missing link error and stop the loop
                formTypeLooper = false;
            }
        }

        if (!(formTypeFormsMatch.length > 0))
        {
            statusUpdates("No exportable form instances found");
            loggingData.push(<div>No exportable form instances found</div>);
            alert("No exportable form instances found");
            return;
        }
        
        //continuing on create up or excel.
        // fixed header structure -
        statusUpdates("Building excel template...");
        loggingData.push(<div>Building excel template...</div>);
        const wb = new Excel.Workbook();
        var ws = wb.addWorksheet("IRS");
        var currCol = 1;
        ws.getCell(1, currCol++).value = "id";
        ws.getCell(1, currCol-1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFFFFF00'} };; //want yellow
        ws.getCell(1, currCol++).value = "number";
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid',fgColor:{argb:'FFFFFF00'}};; //want yellow
        ws.getCell(1, currCol++).value = "displayName";
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red
        ws.getCell(1, currCol++).value = "subject";
        ws.getCell(1, currCol++).value = "description";
        ws.getCell(1, currCol++).value = "location.latitude";
        ws.getCell(1, currCol++).value = "location.longitude";
        ws.getCell(1, currCol++).value = "container.id";
        ws.getCell(1, currCol++).value = "container.url";
        ws.getCell(1, currCol++).value = "container.displayName";
        ws.getCell(1, currCol++).value = "item.id";
        ws.getCell(1, currCol++).value = "item.displayName";
        ws.getCell(1, currCol++).value = "item.url";
        ws.getCell(1, currCol++).value = "elementId";
        ws.getCell(1, currCol++).value = "modelPin.location.x";
        ws.getCell(1, currCol++).value = "modelPin.location.y";
        ws.getCell(1, currCol++).value = "modelPin.location.z";

        ws.getCell(1, currCol++).value = "createdBy";
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red
        ws.getCell(1, currCol++).value = "createdDateTime";
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red
        ws.getCell(1, currCol++).value = "status";
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red
        ws.getCell(1, currCol++).value = "assignee.displayName";
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red
        ws.getCell(1, currCol++).value = "assignee.email";
        ws.getCell(1, currCol++).value = "assignee.id";
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red
        ws.getCell(1, currCol++).value = "dueDate";
        ws.getCell(1, currCol++).value = "state";
        ws.getCell(1, currCol - 1).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FF0000'}};; //want red
        ws.getCell(1, currCol++).value = "assignees";

        var userPropsStartAtCol = currCol;
        var currRow = 1;
        var gotData = false;

        // loop through all the potential matches and extract the properties column headers from the data
        // this is the only way you can actually figure out what controls are on the form
        // and the access keys to formulate for an upload.... very costly! :(
        for(const typeMatchId of formTypeFormsMatch) {
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
            if (formTemplateId === json.issue.formId)
            {
                currRow++;
                statusUpdates(`Exporting form to excel row: ${currRow}`);
                loggingData.push(<div>Exporting form to excel row: {currRow}</div>);
                gotData = true;
                if ("id" in json.issue){
                    ws.getCell(currRow, 1).value = json.issue.id;
                    if (ws.getCell(1, 1).note === undefined)
                    {
                        ws.getCell(1, 1).note = typeof json.issue.id;
                    }
                }

                if ("number" in json.issue){
                    ws.getCell(currRow, 2).value = json.issue.number;
                    if (ws.getCell(1, 2).note === undefined)
                    {
                        ws.getCell(1, 2).note = typeof json.issue.number;
                    }
                }

                if ("displayName" in json.issue){
                    ws.getCell(currRow, 3).value = json.issue.displayName;
                    if (ws.getCell(1, 3).note === undefined)
                    {
                        ws.getCell(1, 3).note = typeof json.issue.displayName;
                    }
                }

                if ("subject" in json.issue){
                    ws.getCell(currRow, 4).value = json.issue.subject;
                    if (ws.getCell(1, 4).note === undefined)
                    {
                        ws.getCell(1, 4).note = typeof json.issue.subject;
                    }
                }

                if ("description" in json.issue){
                    ws.getCell(currRow, 5).value = json.issue.description;
                    if (ws.getCell(1, 5).note === undefined)
                    {
                        ws.getCell(1, 5).note = typeof json.issue.description;
                    }
                }

                if ("location" in json.issue){
                    if ("latitude" in json.issue.location){
                        ws.getCell(currRow, 6).value = json.issue.location.latitude;
                        if (ws.getCell(1, 6).note === undefined)
                        {
                            ws.getCell(1, 6).note = typeof json.issue.location.latitude;
                        }
                    }

                    if ("longitude" in json.issue.location){
                        ws.getCell(currRow, 7).value = json.issue.location.longitude;
                        if (ws.getCell(1, 7).note === undefined)
                        {
                            ws.getCell(1, 7).note = typeof json.issue.location.longitude;
                        }
                    }
                }

                if ("container" in json.issue){
                    if ("id" in json.issue.container){
                        ws.getCell(currRow, 8).value = json.issue.container.id;
                        if (ws.getCell(1, 8).note === undefined)
                        {
                            ws.getCell(1, 8).note = typeof json.issue.container.id;
                        }
                    }
 
                    if ("url" in json.issue.container){
                        ws.getCell(currRow, 9).value = json.issue.container.url;
                        if (ws.getCell(1, 9).note === undefined)
                        {
                            ws.getCell(1, 9).note = typeof json.issue.container.url;
                        }
                    }

                    if ("displayName" in json.issue.container){
                        ws.getCell(currRow, 10).value = json.issue.container.displayName;
                        if (ws.getCell(1, 10).note === undefined)
                        {
                            ws.getCell(1, 10).note = typeof json.issue.container.displayName;
                        }
                    }
                }
                if ("item" in json.issue){
                    if ("id" in json.issue.item){
                        ws.getCell(currRow, 11).value = json.issue.item.id;
                        if (ws.getCell(1, 11).note === undefined)
                        {
                            ws.getCell(1, 11).note = typeof json.issue.item.id;
                        }
                    }

                    if ("displayName" in json.issue.item){
                        ws.getCell(currRow, 12).value = json.issue.item.displayName;
                        if (ws.getCell(1, 12).note === undefined)
                        {
                            ws.getCell(1, 12).note = typeof json.issue.item.displayName;
                        }
                    }

                    if ("url" in json.issue.item){
                        ws.getCell(currRow, 13).value = json.issue.item.url;
                        if (ws.getCell(1, 13).note === undefined)
                        {
                            ws.getCell(1, 13).note = typeof json.issue.item.url;
                        }
                    }
                }

                if ("elementId" in json.issue){
                    ws.getCell(currRow, 14).value = json.issue.elementId;
                    if (ws.getCell(1, 14).note === undefined)
                    {
                        ws.getCell(1, 14).note = typeof json.issue.elementId;
                    }
                }

                if ("modelPin" in json.issue){
                    if ("location" in json.issue.modelPin){
                        if ("x" in json.issue.modelPin.location){
                            ws.getCell(currRow, 15).value = json.issue.modelPin.location.x;
                            if (ws.getCell(1, 15).note === undefined)
                            {
                                ws.getCell(1, 15).note = typeof json.issue.modelPin.location.x;
                            }
                        }
                        if ("y" in json.issue.modelPin.location){
                            ws.getCell(currRow, 16).value = json.issue.modelPin.location.y;
                            if (ws.getCell(1, 16).note === undefined)
                            {
                                ws.getCell(1, 16).note = typeof json.issue.modelPin.location.y;
                            }
                        }
                        if ("z" in json.issue.modelPin.location){
                            ws.getCell(currRow, 17).value = json.issue.modelPin.location.z;
                            if (ws.getCell(1, 17).note === undefined)
                            {
                                ws.getCell(1, 17).note = typeof json.issue.modelPin.location.z;
                            }
                        }
                    }
                }

                if ("createdBy" in json.issue){
                    ws.getCell(currRow, 18).value = json.issue.createdBy;
                    if (ws.getCell(1, 18).note === undefined)
                    {
                        ws.getCell(1, 18).note = typeof json.issue.createdBy;
                    }
                }

                if ("createdDateTime" in json.issue){
                    ws.getCell(currRow, 19).value = json.issue.createdDateTime;
                    if (ws.getCell(1, 19).note === undefined)
                    {
                        ws.getCell(1, 19).note = typeof json.issue.createdDateTime;
                    }
                }

                if ("status" in json.issue){
                    ws.getCell(currRow, 20).value = json.issue.status;
                    if (ws.getCell(1, 20).note === undefined)
                    {
                        ws.getCell(1, 20).note = typeof json.issue.status;
                    }
                }

                if ("assignee" in json.issue){
                    if ("displayName" in json.issue.assignee){
                        ws.getCell(currRow, 21).value = json.issue.assignee.displayName;
                        if (ws.getCell(1, 21).note === undefined)
                        {
                            ws.getCell(1, 21).note = typeof json.issue.assignee.displayName;
                        }
                    }

                    // assignee get some extra treatment...
                    if ("id" in json.issue.assignee){
                        statusUpdates(`Exporting form to excel row: ${currRow} Supplementing assignee data`);
                        loggingData.push(<div>Exporting form to excel row: {currRow} Supplementing assignee data</div>);
                        const usersEmail = await ExcelExporter.getUsersEmailFromGuid(json.issue.assignee.id, projectId);
                        if (usersEmail === "0")
                        {
                            // it is a role
                            ws.getCell(currRow, 22).value = json.issue.assignee.displayName;
                            if (ws.getCell(1, 22).note === undefined)
                            {
                                ws.getCell(1, 22).note = typeof json.issue.assignee.displayName;
                            }
                        }
                        else{
                            ws.getCell(currRow, 22).value = usersEmail;
                        }

                        ws.getCell(currRow, 23).value = json.issue.assignee.id;
                        if (ws.getCell(1, 23).note === undefined)
                        {
                            ws.getCell(1, 23).note = typeof json.issue.assignee.id;
                        }
                    }
                }

                if ("dueDate" in json.issue){
                    ws.getCell(currRow, 24).value = json.issue.dueDate;
                    if (ws.getCell(1, 24).note === undefined)
                    {
                        ws.getCell(1, 24).note = typeof json.issue.dueDate;
                    }
                }

                if ("state" in json.issue){
                    ws.getCell(currRow, 25).value = json.issue.state;
                   // console.log(ws.getCell(1, 25).note,(ws.getCell(1, 25).note === undefined))
                    if (ws.getCell(1, 25).note === undefined)
                    {
                        ws.getCell(1, 25).note = typeof json.issue.state;
                    }
                }
               // console.log(json.issue);
                // assignees get some extra treatment
                if ("assignees" in json.issue){
                    statusUpdates(`Exporting form to excel row: ${currRow} Supplementing assignees data`);
                    loggingData.push(<div>Exporting form to excel row: {currRow} Supplementing assignees data</div>);
                    //console.log("assigneesssss", json.issue.assignees);
                    // assignees is an array (well should be!)
                    var strAssignees = "";
                    for (var i = 0; i < json.issue.assignees.length; i++)
                    {
                        if (strAssignees===""){
                            strAssignees = json.issue.assignees[i].displayName + "|" + await ExcelExporter.getUsersEmailFromGuid(json.issue.assignees[i].id, projectId) + "|" + json.issue.assignees[i].id
                        }
                        else
                        {
                            strAssignees =  strAssignees + "||" + json.issue.assignees[i].displayName + "|" + await ExcelExporter.getUsersEmailFromGuid(json.issue.assignees[i].id, projectId) + "|" + json.issue.assignees[i].id
                        }
                    }

                    ws.getCell(currRow, 26).value = strAssignees;
                    if (ws.getCell(1, 26).note === undefined)
                    {
                        ws.getCell(1, 26).note = "string";
                    }
                }

                // user properties. things get dynamic from here.
                if ("properties" in json.issue)
                {
                    for (var property in json.issue.properties)
                    {
                        statusUpdates(`Exporting form to excel row: ${currRow} Dumping user properties`);
                        loggingData.push(<div>Exporting form to excel row: {currRow} Dumping user properties</div>);
                        var colNum = 0;
                      //  console.log(property, json.issue.properties[property]);
                        for (var x = userPropsStartAtCol; x <= currCol; x++){
                            var header = property;
                            if(header.includes("__x0020__"))
                            {
                                header = header.replace("__x0020__"," ");
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
                                header = header.replace("__x0020__"," ");
                            }
                            ws.getCell(1,currCol).value = "properties." + header;
                         //   console.log(`Adding property to excel named ${header}` );
                         //   console.log("properties." + header);
                            ws.getCell(currRow, currCol).value = json.issue.properties[property];
                            if (ws.getCell(1,currCol).note === undefined){
                                ws.getCell(1,currCol).note = typeof json.issue.properties[property];
                                //console.log ("property typeof=",typeof json.issue.properties[property]);
                            }
                            currCol++;
                        }


                    }

                }


            }

        }
        // double check we go some data
        if (gotData === false){
            statusUpdates("No exportable form instances were found");
            loggingData.push(<div>No exportable form instances were found</div>);
            alert("No exportable form instances were found");
            logger(loggingData);
            return;
        }

        // all forms have been matched and exported.
        // check to see if we need to export comments also
        if (gotData && exportComments)
        { //new__x0020__number
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

    public static async getUsersEmailFromGuid(userGuid:string, projectId:string){
        const accessToken = await IMSLoginProper.getAccessTokenForBentleyAPI();
        const response = await fetch(`https://api.bentley.com/projects/${projectId}/members/${userGuid}`, {
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': accessToken,
                  },
            })
         //   console.log("queried for use ", `https://api.bentley.com/projects/${projectId}/members/${userGuid}`);
            const data = await response;
            if (data.status === 404)
            {
                //most likely a role, or use has been ejected :(
                return "0";
            }
            const json = await data.json();
         //   console.log("member ",json);
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

}