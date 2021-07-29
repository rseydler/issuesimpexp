// @ts-nocheck

import * as Excel from 'exceljs';
import { IMSLoginProper } from './IMSLoginProper';
import React, { useCallback } from 'react'
import { ResponseError } from '@bentley/itwin-client';

interface LooseObject {
    [key: string]: any
}

export class ExcelImporter{

    public static async importIssuesFromExcel(selectedFile:FileReader, selectedProject:string, selectedFormTemplate:string, statusUpdates:any, logger:any, verboseLogging:boolean){
        const loggingData:JSX.Element[] = [];
        const wb = new Excel.Workbook;
        const buffer = selectedFile.result;
        statusUpdates("Loading file");
        loggingData.push(<div><h1>Logger results</h1></div>);
        loggingData.push(<div>Loading file</div>);
        await wb.xlsx.load(buffer);
        statusUpdates("Looking for IRS worksheet");
        loggingData.push(<div>Looking for worksheet named IRS</div>);
        const ws = wb.getWorksheet("IRS");
        console.log(`Extra logging enabled: ${verboseLogging}`);

        if (ws.name==="IRS"){
            var currColumn =1;
            var currRow =2;
            var logColumn = 1;
            var x = 1;
            statusUpdates("Locating the logging column");
            loggingData.push(<div>Locating the logging column</div>);
            //find the column to log stuff into
            while(!(ws.getCell(1,logColumn).value===null)){
                logColumn++;
            // console.log("columns..", logColumn);
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
                // console.log("headercellvalue",currentCellHeaderValue);
                        if (currentCellHeaderValue.includes(".")){
                            //make sure the full notation exists
                            const dsx = currentCellHeaderValue.split(".");
                            // largest is 3 Modelpin.location.x y z
                            // 0 is at the object level. easy
                            // does it exist?
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
                        // console.log("2nd/3rd level",data)
            
                        }
                            else{
                                //blindly throw it in
                                data[currentCellHeaderValue] = ws.getCell(currRow,currColumn).value;
                            // console.log("DataIs",data);
                        }
                    }
                    //logger(loggingData);
                    currColumn++;
                }
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

                if(verboseLogging){console.log("Assignee email type:",typeof data.assignee.email)};
                if(typeof data.assignee.email == "undefined"){
                    if(verboseLogging){console.log("Cleaning out assignee completely")};
                    delete data.assignee
                }
                if("assignee" in data){
                    if ("email" in data.assignee){
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
                                const usersGUID = await this.getUsersGUIDfromEmail(assigneeEmail, selectedProject);
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
                                    const usersDisplayName= await this.getUsersDisplayNameFromGUID(usersGUID,selectedProject);
                                    ExcelImporter.updateObject(data.assignee,"value","displayName", usersDisplayName);
                                    loggingData.push(<div>Processing row {currRow} : Assignee data assimilated</div>);
                                }
                            }
                            else{ //it is a role
                                const roleGUID = await this.getRoldIdFromDisplayName(assigneeEmail,selectedProject);
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
                // SEYDLER|0|91d7cf3f-8ba1-406b-b0de-9a49c1a86d05||Ron Seydler|Ron.Seydler@bentley.com|fe0bd0e6-d9dc-4dec-b013-0bcfbc05a66c
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
                                        assGuid = await this.getUsersGUIDfromEmail(assEmail, selectedProject);
                                        if (assGuid === ""){
                                            processThisRow = false;
                                            statusUpdates(`FAILED User could not be determined for assignees email: ${assEmail}. Row Skipped`);
                                            loggingData.push(<div className="errRow">Processing row {currRow} : FAILED User could not be determined for assignees email: {assEmail}. Row Skipped</div>);
                                            ws.getCell(currRow, logColumn).value = `FAILED User could not be determined for assignees email: ${assEmail}. Row Skipped`;
                                        }
                                    }
                                    else{
                                        //assume role
                                        assGuid = await this.getRoldIdFromDisplayName(assDisplayName,selectedProject);
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
                //console.log("Upload is:",JSON.stringify(data));
                const existingFormId = ws.getCell(currRow,1).value;
                if(existingFormId===null || existingFormId===""){
                    ExcelImporter.updateObject(data,"value","formId", selectedFormTemplate);
                }
                //push it to the service!
                if(processThisRow){
                    statusUpdates(`Processing row ${currRow} uploading form now...`);
                    loggingData.push(<div>Processing row {currRow} : Uploading form</div>);
                    //statusUpdates(loggingData);
                    logger([...loggingData]);
                    //console.log("Pushing this data",JSON.stringify(data));
                    await this.pushInTheChanges(JSON.stringify(data),existingFormId===null?"":existingFormId,selectedFormTemplate, loggingData, verboseLogging);
                }
                else{
                    statusUpdates(`Processing row ${currRow} - row skipped due to errors`);
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
           // console.log("no match for IRS");
        }
        
        statusUpdates("Upload completed");
        loggingData.push(<div>Upload complete</div>);
        logger([...loggingData]);
    }

    public static async pushInTheChanges(theIssue:string, existingId:string, selectedFormId:string, loggingData:any, verboseLogging:boolean ){
        const accessToken = await IMSLoginProper.getAccessTokenForBentleyAPI();
        if(verboseLogging){
            console.log("Preparing to upload this", theIssue);
        }
       // return;
        if (existingId===""){
            
            //new issue pump it out!
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
            const json = await data.json;
            if (data.status === 201)
            {
                loggingData.push(<div>Issue created successfully</div>);
                return "1";
            }
            else{
                console.log("GOT an error",data);
                console.log("GOT an error",json);
                loggingData.push(<div className="errRow">Issue creation failed: {data.status} </div>);
            }
        }
        else{
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
            const json = await data.json;
            if (data.status === 200 || data.status === 201)
            {
                loggingData.push(<div>Issue updated successfully</div>);
                return "1";
            }
            else{
                console.log("GOT an error",data);
                console.log("GOT an error",json);
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
                value = value.replace(" ", "__x0020__");
            }
            myObj[value] = stuff;
            break;
          default:
            break;
        }
        return;
    }

    public static async getUsersDisplayNameFromGUID(userGuid:string, projectId:string ){
       // console.log("starting getUsersDisplayNameFromGUID", userGuid, projectId);
        const accessToken = await IMSLoginProper.getAccessTokenForBentleyAPI();
       // console.log("got token", `https://api.bentley.com/projects/${projectId}/members/${userGuid}`);
        const response = await fetch(`https://api.bentley.com/projects/${projectId}/members/${userGuid}`, {
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': accessToken,
                  },
        })
       // console.log("queried for use ", `https://api.bentley.com/projects/${projectId}/members/${userGuid}`);
        const data = await response;
        if (data.status === 404)
        {
            //most likely a role, or use has been ejected :(
           // console.log("got a 404");
            return "0";
        }
        const json = await data.json();
       // console.log("member: ",json);
        if ("member" in json){
            return json.member.givenName + " " + json.member.surname;
        }
        else{
            return ("0");
        }
    }

    public static async getUsersGUIDfromEmail(userEmail:string, projectId:string ){
        const accessToken = await IMSLoginProper.getAccessTokenForBentleyAPI();
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
        // after all that let's check if it was a role or missing
        if (memberGUID === ""){
            memberGUID = await this.getRoldIdFromDisplayName(userEmail, projectId);
        }
      //  console.log(memberGUID);
        return memberGUID;
    }

    public static async getRoldIdFromDisplayName(userEmail:string, projectId:string){
        const accessToken = await IMSLoginProper.getAccessTokenForBentleyAPI();
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
