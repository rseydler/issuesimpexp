// @ts-nocheck
//import { Button } from '@bentley/ui-core';
import { ImageCheckBox } from '@bentley/ui-core';
import {Button,  Checkbox , CheckboxProps} from '@itwin/itwinui-react';
import React, { useCallback, useState } from 'react'
import { useEffect } from 'react'
import {SvgHelpCircularHollow,SvgNotification, SvgAdd,SvgDelete, SvgFlag, SvgHome,SvgNetwork, SvgSearch, SvgSettings, SvgExport, SvgImport} from "@itwin/itwinui-icons-react";
import { ExcelExporter } from './excelExport';
import { ExcelImporter } from './excelImport';
import logo from './logo.svg';


interface homePageStuff{
  selectedProjectId:string;
  selectedFormsId:string;
  selectedFormsType:string;
}

var projId = "";
var formId = "";
var commentsChecked = false;

function MyHomePage({selectedProjectId, selectedFormsId, selectedFormsType }:homePageStuff) {
  const [checked, setChecked] = React.useState(true);
  const [doing, setDoing] = useState("");
  const [logger, setLogger] = useState<JSX.Element[]>([]);
  //logger stuff
  const toLog = [];
  useEffect(() => {

  },[doing])

    projId = selectedProjectId;
    formId = selectedFormsId;
    const info:JSX.Element[]= [];
    const working:JSX.Element[]= [];
 
    if (selectedProjectId.trim() === "")
    {
        info.push(<div>Please select a project to begin</div>);
        return(<div>{info}</div>)
    }
    if (selectedFormsId.trim() === "")
    {
        info.push(<div>Please select a form definition</div>);
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
          await ExcelImporter.importIssuesFromExcel(fileReader, projId, formId,setDoing, setLogger);
        }
      } catch (error) {
        setDoing(`There was an error reading the file: ${error}`);
      }
    }

    function clickFileButton(){
      document.getElementById('input').click();
    }

    //right we have enough lets make some buttons!
    info.push(<h2>Please select from the following options</h2>)
    info.push(<Button size="large" key="Export Button" name="Export Button" onClick={() => ExcelExporter.exportIssuesToExcel(selectedFormsId, selectedProjectId,selectedFormsType, checked, setDoing, setLogger)}>Export Issues</Button>);
    info.push(<Checkbox label="Export Comments" defaultChecked={checked} key="export comments" onChange={() => setChecked(!checked)} />)
    info.push(<br></br>);
    info.push(<Button size="large" input="file" key="Upload Excel" name="Upload Excel" onClick={() => clickFileButton()}>Upload Excel</Button>);
    info.push(<p></p>);
    info.push(<Button size="small" key="Clean Logger" name="Clean Logger" onClick={() => setLogger([<div></div>])}>Clear Log Display</Button>);
    info.push(<input hidden={true} type="file" id="input" onClick={() => {resetFileSelected()}} onChange={() => {startExelImport()}}/>)
   
  return (
    <div>
      {info}
      <span>
        <div>
          {doing}
          <p></p>
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