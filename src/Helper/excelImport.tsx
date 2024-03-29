// @ts-nocheck

import * as Excel from 'exceljs';
import React from 'react';
import AuthorizationClient from '../AuthorizationClient';

interface LooseObject {
    [key: string]: any
}

export class ExcelImporter{

    public static async importIssuesFromExcel(selectedFile:FileReader, selectedProject:string, selectedFormTemplate:string, statusUpdates:any, logger:any, verboseLogging:boolean){
        const accessToken = await (await AuthorizationClient.oidcClient.getAccessToken()).toTokenString();
        const timeStarted = new Date();
        const timeStartedHolder = timeStarted.toLocaleTimeString([],{hour: '2-digit', minute: '2-digit', second: '2-digit'});
        var errorCount = 0;
        var processCount = 0;
        const loggingData:JSX.Element[] = [];
        const wb = new Excel.Workbook();
        const buffer = selectedFile.result;        
        statusUpdates("Loading file");
        loggingData.push(<div>Started Processing at {timeStartedHolder}</div>);
        loggingData.push(<div><h1>Logger results</h1></div>);
        loggingData.push(<div>Loading file</div>);
        await wb.xlsx.load(buffer);
        statusUpdates("Looking for IRS worksheet");
        loggingData.push(<div>Looking for worksheet named IRS</div>);
        const ws = wb.getWorksheet("IRS");
        if (ws.name==="IRS"){
            var currColumn =1;
            var currRow =2;
            var logColumn = 1;
            //var x = 1;
            statusUpdates("Locating the logging column");
            loggingData.push(<div>Locating the logging column</div>);
            //find the column to log stuff into
            while(!(ws.getCell(1,logColumn).value===null)){
                logColumn++;
            }
            statusUpdates("Processing row data");
            loggingData.push(<div>Processing row data....</div>);
            while((!(ws.getCell(currRow, currColumn).value === null)) || (!(ws.getCell(currRow, currColumn + 1).value === null) )){
                var processThisRow = true;
                loggingData.push(<div>Processing row {currRow}</div>);
                const data : LooseObject = {};
                currColumn = 1;
                while (!(ws.getCell(1,currColumn).value === null)){
                    statusUpdates(`Processing row ${currRow}`);
                    
                    processThisRow = true;
                    if (!(ws.getCell(currRow,currColumn).value === null)){
                    // if the header has a dot then make sure the full property exists!
                    let currentCellHeaderValue:string = ws.getCell(1,currColumn).value?.toString() as string;
                        if (currentCellHeaderValue.includes(".")){
                            //make sure the full notation exists
                            const dsx = currentCellHeaderValue.split(".");
                            
                            switch (dsx.length) {
                            case 3:
                                ExcelImporter.updateObject(data,"object",dsx[0]); //modelpin
                                var tmpObj = data[dsx[0]];
                                ExcelImporter.updateObject(tmpObj, "object", dsx[1]) //location
                                tmpObj = tmpObj[dsx[1]];
                                ExcelImporter.updateObject(tmpObj, "value", dsx[2], ws.getCell(currRow,currColumn).value) //x OR y OR z
                                break;
            
                            case 2:
                                ExcelImporter.updateObject(data,"object",dsx[0]); //modelpin
                                var tmpObj = data[dsx[0]];
                                ExcelImporter.updateObject(tmpObj, "value", dsx[1], ws.getCell(currRow,currColumn).value) //x OR y OR z
                                break;
                            
                            default:
                                break;
                            }
            
                        }
                            else{
                                //blindly throw it in
                                data[currentCellHeaderValue] = ws.getCell(currRow,currColumn).value;
                        }
                    }
                    currColumn++;
                }
                //#region remove invalid upload props
                // some things we cannot upload
                // get rid of those
                if ("id" in data){
                    delete data.id;
                }
                if ("number" in data){
                    delete data.number;
                }
                if ("displayName" in data){
                    delete data.displayName;
                }
                if ("createdBy" in data){
                    delete data.createdBy;
                }
                if ("createdDateTime" in data){
                    delete data.createdDateTime;
                }
                if ("status" in data){
                    delete data.status;
                }

                if ("state" in data){
                    delete data.state;
                }

                if("formGUID" in data){
                    delete data.formGUID;
                }
                //#endregion

                if("assignee" in data){
                    if ("email" in data.assignee){
                        //figure out who the assignee is.
                        if(verboseLogging){console.log("Assignee email type:",typeof data.assignee.email)};
                        if(typeof data.assignee.email == "undefined"){
                            if(verboseLogging){console.log("Cleaning out assignee completely")};
                            delete data.assignee
                        }
                        if(typeof(data.assignee.email) === "object" && "hyperlink" in data.assignee.email){
                            //bloody hyperlinks!
                            const tmpEmail = data.assignee.email.text;
                            delete data.assignee.email;
                            ExcelImporter.updateObject(data.assignee,"value","email", tmpEmail);
                        }
                        const assigneeEmail:string = data.assignee.email;
                        statusUpdates(`Processing row ${currRow} sorting out assignee details`);
                        loggingData.push(<div>Processing row {currRow} : sorting out assignee details</div>);
                        if (!(assigneeEmail.trim() === "")){
                            if (assigneeEmail.includes("@")){
                                const usersGUID = await this.getUsersGUIDfromEmail(assigneeEmail, selectedProject, accessToken);
                                if (usersGUID === "")
                                {   
                                    processThisRow = false;
                                    statusUpdates("FAILED User could not be determined for assignee email address. Row Skipped");
                                    loggingData.push(<div className="errRow">Processing row {currRow} : FAILED User could not be determined for assignee email address. Row Skipped</div>);
                                    ws.getCell(currRow, logColumn).value = "FAILED User could not be determined for assignee email address. Row Skipped";
                                }
                                else
                                {
                                    //clean up the data for the load
                                    delete data.assignee;
                                    ExcelImporter.updateObject(data,"object","assignee", "");
                                    ExcelImporter.updateObject(data.assignee,"value","id", usersGUID);
                                    const usersDisplayName= await this.getUsersDisplayNameFromGUID(usersGUID,selectedProject, accessToken);
                                    ExcelImporter.updateObject(data.assignee,"value","displayName", usersDisplayName);
                                    loggingData.push(<div>Processing row {currRow} : Assignee data assimilated</div>);
                                }
                            }
                            else{ //it is a role
                                const roleGUID = await this.getRoldIdFromDisplayName(assigneeEmail,selectedProject, accessToken);
                                if (!(roleGUID === "")){
                                    delete data.assignee;
                                    ExcelImporter.updateObject(data,"object","assignee", "");
                                    ExcelImporter.updateObject(data.assignee,"value","id", roleGUID);
                                    ExcelImporter.updateObject(data.assignee,"value","displayName", assigneeEmail);
                                    loggingData.push(<div>Processing row {currRow} : Assignee role data assimilated</div>);
                                }
                                else{
                                    //log an error?
                                    processThisRow = false;
                                    ws.getCell(currRow, logColumn).value = "FAILED Role could not be determined for assignee email address field. Row Skipped";
                                    statusUpdates("FAILED User could not be determined for assignee email address. Row Skipped");
                                    loggingData.push(<div className="errRow">Processing row {currRow} : FAILED User could not be determined for assignee email address. Row Skipped</div>);
                                }

                            }
                        }
                        else{
                            //clean up data incase user left in some entries!
                            delete data.assignee;
                            if(data.id === "" || data.id === undefined || data.id === null){
                                statusUpdates(`No assignee data entered for row ${currRow} You will be the assignee.`);
                                loggingData.push(<div>No assignee data entered for row {currRow} You will be the assignee</div>);
                            }
                        }
                    }
                    else{
                        if(data.id === "" || data.id === undefined || data.id === null){
                            statusUpdates(`No assignee data entered for row ${currRow} You will be the assignee.`);
                            loggingData.push(<div>No assignee data entered for row {currRow} You will be the assignee</div>);
                        }
                    }
                }
                //work on the assignees
                //first split them up by ||
                if(typeof(data.assignees) === "string"){
                    if(data.assignees.trim() === ""){
                        delete data.assignees
                    }
                }
                if (typeof(data.assignees) === "object"){
                    const tmpAssigneesHolder = data.assignees.text;
                    delete data.assignees;
                    this.updateObject(data, "value", "assignees", tmpAssigneesHolder);
                }
                try {
                    statusUpdates(`Processing row ${currRow} sorting out assignees details`);
                    loggingData.push(<div>Processing row {currRow} : Sorting out assignees data</div>);
                    if ("assignees" in data){
                        const assignees:string = data.assignees;
                        const assignee = assignees.split("||");
                        var assigneesData:{displayName:string;id:string;isRole:false;}[] = [];
                        for (var i=0;i<assignee.length;i++){
                            if (!(assignee[i] === "")){
                                const elementData = assignee[i].split("|");
                                var assDisplayName = "";
                                var assEmail = "";
                                var assGuid = "";
                                var assIsRole = false;
                                if (elementData.length === 3){
                                    assDisplayName = elementData[0];
                                    assEmail = elementData[1];
                                    //get the real guid
                                    if (assEmail.includes("@")){
                                        assGuid = await this.getUsersGUIDfromEmail(assEmail, selectedProject, accessToken);
                                        if (assGuid === ""){
                                            processThisRow = false;
                                            statusUpdates(`FAILED User could not be determined for assignees email: ${assEmail}. Row Skipped`);
                                            loggingData.push(<div className="errRow">Processing row {currRow} : FAILED User could not be determined for assignees email: {assEmail}. Row Skipped</div>);
                                            ws.getCell(currRow, logColumn).value = `FAILED User could not be determined for assignees email: ${assEmail}. Row Skipped`;
                                        }
                                    }
                                    else{
                                        //assume role
                                        assGuid = await this.getRoldIdFromDisplayName(assDisplayName,selectedProject, accessToken);
                                        assIsRole = true;
                                        if (assGuid === ""){
                                            statusUpdates(`FAILED ROLE could not be determined for assignees role name: ${assDisplayName}. Row Skipped`);
                                            loggingData.push(<div className="errRow">Processing row {currRow} : FAILED ROLE could not be determined for assignees role name: {assDisplayName}. Row Skipped</div>);
                                            ws.getCell(currRow, logColumn).value = `FAILED ROLE could not be determined for assignees role name: ${assDisplayName}. Row Skipped`;
                                            processThisRow=false;
                                        }
                                    }
                                }
                                else{
                                    statusUpdates(`FAILED Invalid definition for assignees, Check your entry for row ${currRow}. Row Skipped`);
                                    loggingData.push(<div className="errRow">FAILED Invalid definition for assignees, Check your entry for row {currRow}. Row Skipped</div>);
                                    processThisRow = false;
                                }
                            }
                            else{
                                //nothing to process. field is blank.
                            }
                            //before we loop stuff the variables :)
                            if (!(assDisplayName === "") && !(assEmail === "")){
                                assigneesData.push({displayName:assDisplayName, id:assGuid, isRole:assIsRole});
                            }
                        }
                    }
                    else{
                        statusUpdates(`No assignees data entered for row ${currRow}`);
                        loggingData.push(<div>No assignees data entered for row {currRow}</div>);
                    }
                    if("assignees" in data){
                        delete data.assignees;
                        if (assigneesData.length > 0 ){
                            ExcelImporter.updateObject(data,"value","assignees", assigneesData);
                        }
                        loggingData.push(<div>Processing row {currRow} : Assignees data assimilated</div>);
                    }
                
                } catch (error) {
                    statusUpdates(`Processing row ${currRow} ooooops ${error}`);
                    loggingData.push(<div>Processing row {currRow} : Something weird happened</div>);
                    console.log(`Processing row ${currRow} ooooops ${error}`);
                }
                const existingFormId = ws.getCell(currRow,1).value;
                if(existingFormId===null || existingFormId===""){
                    ExcelImporter.updateObject(data,"value","formId", selectedFormTemplate);
                }
                //push it to the service!
                if(processThisRow){
                    statusUpdates(`Processing row ${currRow} uploading form now...`);
                    loggingData.push(<div>Processing row {currRow} : Uploading form</div>);
                    logger([...loggingData]);
                    await this.pushInTheChanges(JSON.stringify(data),existingFormId===null?"":existingFormId,selectedFormTemplate, loggingData, verboseLogging, accessToken, errorCount, processCount);
                }
                else{
                    statusUpdates(`Processing row ${currRow} - row skipped due to errors`);
                    errorCount++;
                    loggingData.push(<div>Processing row {currRow} : Row skipped due to errors</div>);
                    logger([...loggingData]);
                }
                currRow++;
                currColumn=1;
                }    
            logger([...loggingData]);
        }
        else
        {
            statusUpdates("IRS Sheet not found in the selected workbook.");
            loggingData.push(<div className="errRow">IRS Sheet not found in the selected workbook.</div>);
            alert("There was no matiching IRS sheet in the workbook selected.")
        }
        
        statusUpdates("Upload completed");
        loggingData.push(<div>Upload complete</div>);
        loggingData.push(<div><h2>Summary Data</h2></div>);
        loggingData.push(<div>Total issues created/updated: {currRow-2}</div>);
        loggingData.push(<div>Total issues with errors: {errorCount}</div>);
        loggingData.push(<div>Time started: {timeStartedHolder}</div>);
        const finishedTime = new Date();
        const finishTimeHolder = finishedTime.toLocaleTimeString([],{hour: '2-digit', minute: '2-digit', second: '2-digit'});
        loggingData.push(<div>Time ended: {finishTimeHolder}</div>);
        logger([...loggingData]);
    }

    public static async pushInTheChanges(theIssue:string, existingId:string, selectedFormId:string, loggingData:any, verboseLogging:boolean, accessToken:any, errorCount:any, processCount:any ){
        if(verboseLogging){console.log("Preparing to upload this", theIssue);}
                
        if (existingId===""){
            //new issue
            const response = await fetch(`https://api.bentley.com/issues/`, {
                method: 'POST',
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': accessToken,
                  },
                body: theIssue,
            })
            const data = await response;
            if (data.status === 201)
            {
                loggingData.push(<div>Issue created successfully</div>);
                processCount++;
                return "1";
            }
            else{
                //TODO get a better error descriptor
                loggingData.push(<div className="errRow">Issue creation failed: {data.status} </div>);
                errorCount++;
            }
        }
        else{
            //existing issue
            const response = await fetch(`https://api.bentley.com/issues/${existingId}`, {
                method: 'PATCH',
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': accessToken,
                  },
                  body: theIssue,
            })
            const data = await response;
            if (data.status === 200 || data.status === 201)
            {
                loggingData.push(<div>Issue updated successfully</div>);
                processCount++;
                return "1";
            }
            else{
                //TODO get a better error descriptor
                errorCount++;
                loggingData.push(<div className="errRow">Issue update failed: {data.status}</div>);
            }
        }
    }

    public static updateObject(myObj:Object, doWhat:string, value:any, stuff:any){
        switch (doWhat.toLowerCase()) {
          case "object":
            if (value in myObj){
            }
            else
            {
              myObj[value] = {};
            }
            break;
        
          case "value":
            if (value.includes(" "))
            {
                value = value.replaceAll(" ", "__x0020__");
            }
            myObj[value] = stuff;
            break;
          default:
            break;
        }
        return;
    }

    public static async getUsersDisplayNameFromGUID(userGuid:string, projectId:string, accessToken:any ){
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
            return json.member.givenName + " " + json.member.surname;
        }
        else{
            return ("0");
        }
    }

    public static async getUsersGUIDfromEmail(userEmail:string, projectId:string , accessToken:any){
        var looper = true; //used to force the initial call;
        var memberGUID = "";
        var urlToQuery : string = `https://api.bentley.com/projects/${projectId}/members`;
        while (looper && (memberGUID === "")) {
            const response = await fetch(urlToQuery, {
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': accessToken,
                  },
            })
            const data = await response;
            const json = await data.json();
            json.members.forEach(member => {
                const memberEmail:string = member.email;
                if (memberEmail.toLowerCase() === userEmail.toLowerCase()){
                    memberGUID = member.userId;
                    return;
                }
            });
            //let see if we are continuing.
            try {
                if (memberGUID === ""){
                    looper = true;
                    urlToQuery = json._links.next.href;
                }
                else{
                    looper = false;
                }
            } catch (error) {
                // better than === undefined?
                //swallow the missing link error and stop the loop
                looper = false;
            }
        }
        //let's check if it was a role or missing
        if (memberGUID === ""){
            memberGUID = await this.getRoldIdFromDisplayName(userEmail, projectId, accessToken);
        }
        return memberGUID;
    }

    public static async getRoldIdFromDisplayName(userEmail:string, projectId:string, accessToken:any){
        var memberGUID = "";
        var looper=true;
        var urlToQuery : string = `https://api.bentley.com/projects/${projectId}/roles`;
        while (looper && (memberGUID === "")) {
            const response = await fetch(urlToQuery, {
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': accessToken,
                  },
            })
            const data = await response;
            const json = await data.json();
            json.roles.forEach(role => {
                const roleDN:string = role.displayName;
                if (roleDN.toLowerCase() === userEmail.toLowerCase()){
                    memberGUID = role.id;
                    return;
                }
            });
            //let see if we are continuing.
            try {
                if (memberGUID === ""){
                    looper = true;
                    urlToQuery = json._links.next.href;
                }
                else{
                    looper = false;
                }
            } catch (error) {
                // better than === undefined?
                //swallow the missing link error and stop the loop
                looper = false;
            }
        }
        return memberGUID;
    }
      
}
