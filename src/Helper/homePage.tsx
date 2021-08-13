// @ts-nocheck
import {Button,  Checkbox , CheckboxProps} from '@itwin/itwinui-react';
import React, { useCallback, useState } from 'react'
import { ExcelExporter } from './excelExport';
import { ExcelImporter } from './excelImport';


interface homePageStuff{
  selectedProjectId:string;
  selectedFormsId:string;
  selectedFormsType:string;
  verboseLogging:boolean;
}

var projId = "";
var formId = "";

function MyHomePage({selectedProjectId, selectedFormsId, selectedFormsType, verboseLogging }:homePageStuff) {
  const [checked, setChecked] = React.useState(true);
  const [doing, setDoing] = useState("");
  const [logger, setLogger] = useState<JSX.Element[]>([]);
  //logger stuff

    projId = selectedProjectId;
    formId = selectedFormsId;
    const info:JSX.Element[]= [];
    const working:JSX.Element[]= [];
 
    if (selectedProjectId.trim() === "")
    {
        info.push(<div key="comments1">Please select a project to begin</div>);
        return(<div>{info}</div>)
    }
    if (selectedFormsId.trim() === "")
    {
        info.push(<div key="comments2">Please select a form definition</div>);
        return(<div>{info}</div>)
    }

    function resetFileSelected(){
      document.getElementById('input').value = "";
    }

    function startExelImport(){
      try {
        const fileName = document.getElementById('input').files[0];
        setDoing(`Started processing excel file`);
        const fileReader = new FileReader();
        fileReader.readAsArrayBuffer(fileName);
        fileReader.onload = async() =>{
          await ExcelImporter.importIssuesFromExcel(fileReader, projId, formId,setDoing, setLogger, verboseLogging);
        }
      } catch (error) {
        setDoing(`There was an error reading the file: ${error}`);
      }
    }

    function clickFileButton(){
      document.getElementById('input').click();
    }

    //right we have enough lets make some buttons!
    info.push(<h2 key="h2">Please select from the following options</h2>)
    info.push(<Button size="large" key="Export Button" name="Export Button" onClick={() => ExcelExporter.exportIssuesToExcel(selectedFormsId, selectedProjectId,selectedFormsType, checked, setDoing, setLogger, verboseLogging)}>Export Issues</Button>);
    info.push(<Checkbox label="Export Comments" defaultChecked={checked} key="export comments" onChange={() => setChecked(!checked)} />)
    info.push(<br key="br"></br>);
    info.push(<Button size="large" input="file" key="Upload Excel" name="Upload Excel" onClick={() => clickFileButton()}>Upload Excel</Button>);
    info.push(<p key="p"></p>);
    info.push(<Button size="small" key="Clean Logger" name="Clean Logger" onClick={() => setLogger([<div></div>])}>Clear Log Display</Button>);
    info.push(<input hidden={true} key="hidden button" type="file" id="input" onClick={() => {resetFileSelected()}} onChange={() => {startExelImport()}}/>)
    if(verboseLogging){
      info.push(<div>Extra logging enabled</div>);
    }
   
  return (
    <div>
      {info}
      <span>
        <div>
          {doing}
          {logger}          
        </div>
      </span>
    </div>
    
  )
}

function ExportIssuesCallBack(isExportCommentsSelected:boolean){
  alert(formId + " " + projId + " " + isExportCommentsSelected);
}

export default MyHomePage