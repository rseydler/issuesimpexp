import "./App.scss";

import React, { useCallback, useEffect, useState } from "react";

import AuthorizationClient from "./AuthorizationClient";
import { SvgHome,SvgNetwork, SvgSettings, SvgExport, } from "@itwin/itwinui-icons-react";
import {  Header, HeaderBreadcrumbs, HeaderButton, HeaderLogo, IconButton,  MenuItem, SidenavButton, SideNavigation,  UserIcon} from "@itwin/itwinui-react";
import { ProjectData } from "./Helper/ProjectData";
import {ThemeButton} from "./Helper/ThemeButton"
import MyHomePage from "./Helper/homePage"
import { Checkbox, UnderlinedButton } from "@bentley/ui-core";

const App: React.FC = () => {
  const [isAuthorized, setIsAuthorized] = useState(
    AuthorizationClient.oidcClient
      ? AuthorizationClient.oidcClient.isAuthorized
      : false
  );
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [sideBarChosen, setsideBarChosen] = useState("issues");
  const [bodyData, setbodyData] = useState<JSX.Element[]>([]);
  const [verboseLogging, setVerboseLogging] = useState(false);
//#region  LoginStuff

  useEffect(() => {
    const initOidc = async () => {
      if (!AuthorizationClient.oidcClient) {
        await AuthorizationClient.initializeOidc();
      }

      try {
        // attempt silent signin
        await AuthorizationClient.signInSilent();
        setIsAuthorized(AuthorizationClient.oidcClient.isAuthorized);
      } catch (error) {
        // swallow the error. User can click the button to sign in
      }
    };
    initOidc().catch((error) => console.error(error));
  }, []);

  useEffect(() => {
    if (isLoggingIn && isAuthorized) {
      setIsLoggingIn(false);
    }
  }, [isAuthorized, isLoggingIn]);

  /*
  useEffect(() => {
    if (isAuthorized){
      // do the other login silent like
      (async () => {
        var token = await IMSLoginProper.getAccessTokenForBentleyAPI();
        setimsToken(token);
      })();
    }
  }, [isAuthorized])
*/
  const onLoginClick = async () => {
    setIsLoggingIn(true);
    await AuthorizationClient.signIn();
  };

  const onLogoutClick = async () => {
    setIsLoggingIn(false);
    await AuthorizationClient.signOut();
    setIsAuthorized(false);
  };
//#endregion
  
//#region BreadCrumb1 Projects Drop Down
//get the recent projects and stuff them into the 1st breadcrumb element
const [projectDetails, setprojectDetails] = useState([{name: "Pick a Project", description: "...to Start", id: ""}]);
const [projectLabel, setprojectLabel] = useState({displayName:"Loading....", description: "....", id:""});
var selectedProjectId:string = "";

 //get the list of recent projects once when authorized
 useEffect(() =>{
  if (!isAuthorized){
    setprojectLabel({displayName:"Waiting on login",description:"", id:""});
    return;
  }
  setprojectLabel({displayName:"Loading...",description:"", id:""});
  // clear any other dependant drop downs also.
  (async () => {
    setprojectDetails([]);
    const recentProjects = [{name: "", description: "", id: ""}];
    recentProjects.shift();
    //get the iModel List and stuff it
    if(selectedProjectId === "" || selectedProjectId === undefined){
      const response = await ProjectData.getRecentProjects();
      const data = await response;
      if(data.length === 0){
        recentProjects.push({name:"No recent projects found",description:"", id:""});
      }
      else{
        for (var x = 0; x < data.length; x++){
          //loop through the projects result
          recentProjects.push({name:data[x].projectNumber, description:data[x].displayName, id:data[x].id});
        }
        setprojectDetails(recentProjects);
      }
    }
    setprojectLabel({displayName:"Select a project from the list",description:"", id:""});
  })();
},[isAuthorized])


// formulate the menu entries - update when something is picked or the contextid changes in the url
const recentProject = useCallback((close: () => void) => {
  var menuItemsToReturn : JSX.Element[] = [];
  for (var x = 0; x < projectDetails.length; x++){
      menuItemsToReturn.push(
      <MenuItem
      key={projectDetails[x].id}
      title={projectDetails[x].name}
      value={projectDetails[x].id}
      id={projectDetails[x].id}
      onClick={(value: any) => {         
        handleProjectChange(value);
          close(); // close the dropdown menu
      }}
      isSelected={(projectDetails[x].name === projectLabel.displayName) ? true : false}
      >
      {projectDetails[x].name}
      </MenuItem>
      )
  }

  return(menuItemsToReturn);
},[projectLabel]) 

const handleProjectChange = useCallback(value =>{
  //got the ID(GUID) so find the displayName
  projectDetails.forEach(project =>{
    if(project.id === value)
    {
      setprojectLabel({displayName:project.name, description:project.description, id:project.id});
     // setContextId(project.id);
    }
  })
},[projectDetails])

//Forms template chooser
const [formDetails, setformDetails] = useState([{name: "Pick a Form", description: "...to Start", id: "", type:""}]);
const [formLabel, setformLabel] = useState({displayName:"Loading....", description: "....", id:"", type:""});

 //get the list of recent projects once when authorized
 useEffect(() =>{
  if (!isAuthorized){
    setformLabel({displayName:"Waiting on login",description:"", id:"", type:""});
    return;
  }
  setformLabel({displayName:"Loading...",description:"", id:"", type:""});
  // clear any other dependant drop downs also.
  (async () => {
    setformDetails([]);
    const formDefs = [{name: "", description: "", id: "", type: ""}];
    formDefs.shift();
    //get the iModel List and stuff it
    if(projectLabel.id.trim() !== ""){
      const response = await ProjectData.getFormTemplatesFromProject(projectLabel.id);
      const data = await response;
      if(data.length === 0){
        formDefs.push({name:"No recent projects found",description:"", id:"", type:""});
      }
      else{
        for (var x = 0; x < data.length; x++){
          //loop through the forms result
          formDefs.push({name:data[x].displayName, description:data[x].displayName, id:data[x].formId, type:data[x].type});
        }
        setformDetails(formDefs);
      }
    }
    if (projectLabel.id === ""){
      setformLabel({displayName:"Select a project first",description:"", id:"", type:""});
    }
    else
    {
      setformLabel({displayName:"Select a form definition",description:"", id:"", type:""});
    }
  })();
},[projectLabel,isAuthorized])


// formulate the menu entries - update when something is picked or the contextid changes in the url
const issueTemplates = useCallback((close: () => void) => {
  var menuItemsToReturn : JSX.Element[] = [];
  for (var x = 0; x < formDetails.length; x++){
      menuItemsToReturn.push(
      <MenuItem
      key={formDetails[x].id}
      title={formDetails[x].name}
      value={formDetails[x].id}
      id={formDetails[x].id}
      onClick={(value: any) => {         
        handleFormChange(value);
          close(); // close the dropdown menu
      }}
      isSelected={(formDetails[x].name === formLabel.displayName) ? true : false}
      >
      {formDetails[x].name}
      </MenuItem>
      )
  }

  return(menuItemsToReturn);
},[formLabel]) 

const handleFormChange = useCallback(value =>{
  //got the ID(GUID) so find the displayName
  formDetails.forEach(form =>{
    if(form.id === value)
    {
      setformLabel({displayName:form.name, description:form.description, id:form.id, type:form.type});
    }
  })
},[formDetails])

// setup for the sidebar
useEffect(() =>  {
  if (sideBarChosen === "issues") // Home Page
  {
    const bodyData: JSX.Element[] = [];
    bodyData.push(<MyHomePage key="HomePageStuff" selectedProjectId={projectLabel.id}  selectedFormsId={formLabel.id} selectedFormsType={formLabel.type} verboseLogging={verboseLogging}/>);
    setbodyData(bodyData);
  }

  if (sideBarChosen === "settings") // Settings Page
  {
    const bodyData: JSX.Element[] = [];
    bodyData.push(<h1 key="settingsPage">Optional Settings</h1>);
    bodyData.push(<Checkbox label="Enable Logging to Console" defaultChecked={verboseLogging} key="verboseLogging" onChange={() => setVerboseLogging(!verboseLogging)} />);
    bodyData.push(<h3><a href={`${window.location.origin}/Formsloader.docx`}>Download Guide</a></h3>)
    setbodyData(bodyData);
  }
},[sideBarChosen, formLabel, projectLabel, verboseLogging])

  return (
    <div className="app">
       <Header
       appLogo={<HeaderLogo logo={<SvgExport />}>Issue Export and Import</HeaderLogo>}
        breadcrumbs={
          <HeaderBreadcrumbs
            items={[
                <HeaderButton
                  key="projectBreadcrumb"
                  menuItems={recentProject}
                  name={projectLabel.displayName}
                  description={projectLabel.description}
                  startIcon={<SvgNetwork />}
                />,
                <HeaderButton
                key="formBreadcrumb"
                menuItems={issueTemplates}
                name={formLabel.displayName}
                description={formLabel.type}
                startIcon={<SvgNetwork />}
              />
            ]}
          />
        }
        actions={[<ThemeButton key="themeSwitched" />]}
        userIcon={
          <IconButton styleType="borderless"  onClick={() => {isAuthorized ? onLogoutClick() : onLoginClick()} }>
            <UserIcon
            className={isAuthorized===true ? "App-logo-noSpin" : "App-logo"} 
              size="medium"
              status={isAuthorized ? "online" : "offline"}
              image={
                <img
                  src="https://itwinplatformcdn.azureedge.net/iTwinUI/user-placeholder.png"
                  alt="User icon"
                />
              }
            />
          </IconButton>
        }
      />
      <div className="app-body">
        <SideNavigation
          items={[
            <SidenavButton onClick={() => {setsideBarChosen("issues")}} isActive={sideBarChosen==="issues" ? true : false} startIcon={<SvgHome />} key="Home">
              Home
            </SidenavButton>
        ]}
        secondaryItems={[
          <SidenavButton onClick={() => {setsideBarChosen("settings")}} isActive={sideBarChosen==="settings" ? true : false} startIcon={<SvgSettings />} key="settings">
            Settings
          </SidenavButton>
        ]}
        />       
     
       <div className="app-container">
          <div className="app-content">
          {isLoggingIn ? ( <span>"Logging in...."</span> ) : (!isAuthorized ? (<h1>You need to <UnderlinedButton title="login" onClick={() => { isAuthorized ? onLogoutClick() : onLoginClick(); } } children={"login"}></UnderlinedButton> first.</h1>) : (bodyData))}
          </div>
          {//place your footer here <Footer />
          }
        </div>
    </div>
  </div>
    
  );
};

export default App;
