//import { Button } from '@bentley/ui-core';
import { ImageCheckBox } from '@bentley/ui-core';
import {Button,  Checkbox , CheckboxProps} from '@itwin/itwinui-react';
import React, { useCallback, useState } from 'react'
import { useEffect } from 'react'
import {SvgHelpCircularHollow,SvgNotification, SvgAdd,SvgDelete, SvgFlag, SvgHome,SvgNetwork, SvgSearch, SvgSettings, SvgExport, SvgImport} from "@itwin/itwinui-icons-react";
import { ExcelExporter } from './excelExport';


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
    projId = selectedProjectId;
    formId = selectedFormsId;
    const info:JSX.Element[]= [];

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
    //right we have enough lets make some buttons!
    info.push(<h2>Please select from the following option</h2>)
    info.push(<Button size="large" key="Export Button" name="Export Button" onClick={() => ExcelExporter.exportIssuesToExcel(selectedFormsId, selectedProjectId,selectedFormsType, checked)}>Create Excel for defintion / Export Issues</Button>);
    info.push(<Checkbox label="Export Comments" defaultChecked={checked} key="export comments" onChange={() => setChecked(!checked)} />)
   /* info.push(<Button size="large" key="Upload Excel" name="Upload Excel" onClick={() => ExcelExporter.tester()}>Upload Excel</Button>);
    info.push(<Button size="large" key="Export Tester" name="Export Tester" onClick={() => ExcelExporter.exportIssuesToExcel(selectedFormsId, selectedProjectId,selectedFormsType, checked)}>Export Tester</Button>);
    */

  return (
    <div>
      {info}
    </div>
  )
}

function ExportIssuesCallBack(isExportCommentsSelected:boolean){
  alert(formId + " " + projId + " " + isExportCommentsSelected);
}

function ImportIssueCallBack(){
  
}

export default MyHomePage