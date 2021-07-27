import "./App.scss";

import { Viewer } from "@bentley/itwin-viewer-react";
import React, { useCallback, useEffect, useState } from "react";

import AuthorizationClient from "./AuthorizationClient";
import { Header as OrigHeader } from "./Header";
import {SvgHelpCircularHollow,SvgNotification, SvgAdd,SvgDelete, SvgFlag, SvgHome,SvgNetwork, SvgSearch, SvgSettings, SvgExport, SvgImport} from "@itwin/itwinui-icons-react";
import { Button, ButtonGroup, DropdownMenu, Footer, Header, HeaderBreadcrumbs, HeaderButton, HeaderLogo, IconButton, LabeledInput, MenuItem, SidenavButton, SideNavigation, Table, Title, toaster, UserIcon, useTheme} from "@itwin/itwinui-react";
import { IMSLoginProper } from "./Helper/IMSLoginProper";
import { ProjectData } from "./Helper/ProjectData";
import {ThemeButton} from "./Helper/ThemeButton"
import MyHomePage from "./Helper/homePage"

const App: React.FC = () => {
  const [isAuthorized, setIsAuthorized] = useState(
    AuthorizationClient.oidcClient
      ? AuthorizationClient.oidcClient.isAuthorized
      : false
  );
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [imsToken, setimsToken] = useState("");
  const [BC1Details, setBC1Details] = useState({name: "Pick a Project", description: "...to Start", id: ""});
  const [BC2Details, setBC2Details] = useState({displayName: "Pick a Project First", type: "Then Pick a form Template", proJid: "", formId:"", status:""});
  const [homeFlag, sethomeFlag] = useState(true);
  const [settingsFlag, setsettingsFlag] = useState(false);
  const [sideBarChosen, setsideBarChosen] = useState(0);
  const [bodyData, setbodyData] = useState<JSX.Element[]>([]);
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

  useEffect(() => {
    if (isAuthorized){
      // do the other login silent like
      (async () => {
        var token = await IMSLoginProper.getAccessTokenForBentleyAPI();
        setimsToken(token);
      })();
    }
  }, [isAuthorized])

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

const recentProject = useCallback((close: () => void) => {
  var menuItemsToReturn : JSX.Element[] = [];
  if (isAuthorized){
  ProjectData.getRecentProjects().then(res => {
    for (var x = 0; x < res.length; x++){
        menuItemsToReturn.push(
        <MenuItem
        key={res[x].id}
        title={res[x].projectNumber}
        value={res[x].id} 
        id={res[x].id}
        onClick={(value) => {         
            handleProjectInputChange(value);
            close(); // close the dropdown menu
        }}
        isSelected={(res[x].id === BC1Details.id) ? true : false}
        >
        {res[x].projectNumber} -- {res[x].displayName}
        </MenuItem>
        )
    }
  })
  .catch(error => {console.log("Caught this nasty", error.message); 
      menuItemsToReturn.push(
          <MenuItem
          key="1" title={error.essage}
          onClick={(value) => {close()}}
          >Ooops - {error.message}</MenuItem>)
          })
  return (menuItemsToReturn);
  }
  else
  {
    return ([<MenuItem key="1">Please Login</MenuItem>])
  }
},[isAuthorized, BC1Details])

//handle the change to breadcrumb1
 const handleProjectInputChange = useCallback(value => {
  setBC2Details({displayName: "Loading...." ,type: "Won't be a moment" , formId:"" ,status: "", proJid: ""});
  ProjectData.getProjectData(value).then(projData => {
    setBC1Details({name: projData.projectNumber ,description: projData.displayName , id: projData.id});
    setBC2Details({displayName: "Pick a Project First" ,type: "Then Pick a form Template" , formId: BC2Details.formId ,status: BC2Details.status , proJid: projData.id})
    })
}, [BC1Details]);
//#endregion



//Forms template chooser

  const issueTemplates = useCallback((close: () => void) => {
    var menuItemsToReturn : JSX.Element[] = [];
    if (!isAuthorized)
    {
      return ([<MenuItem key="1">Please Login</MenuItem>])
    }
  
    if (BC1Details.id !== ""){
    ProjectData.getFormTemplatesFromProject(BC1Details.id).then(res => {
      for (var x = 0; x < res.length; x++){
          menuItemsToReturn.push(
          <MenuItem
          key={res[x].formId}
          title={res[x].displayName}
          value={res[x].formId} 
          id={res[x].formId}
          onClick={(value) => {         
              handleFormSelectionChange(value);
              close(); // close the dropdown menu
          }}
          isSelected={(res[x].formId == BC2Details.formId) ? true : false}
          >
          {res[x].displayName} -- {res[x].type}
          </MenuItem>
          )
      }
    })
    .catch(error => {console.log("Caught this error", error.message); 
        menuItemsToReturn.push(
            <MenuItem
            key="1" title={error.essage}
            onClick={(value) => {close()}}
            >Ooops - {error.message}</MenuItem>)
    })
    return (menuItemsToReturn);
    }
    else
    {
      return ([<MenuItem key="1">Please Select a Project First</MenuItem>])
    }
  },[BC1Details, BC2Details]) 

//end forms template chooser

//handle the change to breadcrumb2
const handleFormSelectionChange = useCallback(value => {
  //get form details
  ProjectData.getFormDetailsFromId(value).then(formData => {
    setBC2Details({displayName: formData.displayName ,type: formData.type , formId: formData.id ,status: formData.status , proJid: BC1Details.id})
  });
}, [BC1Details]);

// setup for the sidebar
useEffect(() =>  {
  if (sideBarChosen === 0) // Home Page
  {
    const bodyData: JSX.Element[] = [];
    bodyData.push(<MyHomePage key="HomePageStuff" selectedProjectId={BC1Details.id}  selectedFormsId={BC2Details.formId} selectedFormsType={BC2Details.type}/>);
    setbodyData(bodyData);
    sethomeFlag(true);
    setsettingsFlag(false);
  }

  if (sideBarChosen === 99) // Settings Page
  {
    const bodyData: JSX.Element[] = [];
    bodyData.push(<h1 key="settingsPage">This page not currently in use</h1>);
    setbodyData(bodyData);
    sethomeFlag(false);
    setsettingsFlag(true);
  }
},[sideBarChosen, BC1Details, BC2Details])

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
                  name={BC1Details.name}
                  description={BC1Details.description}
                  startIcon={<SvgNetwork />}
                />,
                <HeaderButton
                key="formBreadcrumb"
                menuItems={issueTemplates}
                name={BC2Details.displayName}
                description={BC2Details.type}
                startIcon={<SvgNetwork />}
              />
            ]}
          />
        }
        actions={[<ThemeButton key="themeSwitched" />]}
        userIcon={
          <IconButton styleType="borderless"  onClick={() => {isAuthorized ? onLogoutClick() : onLoginClick()} }>
            <UserIcon
            className="App-logo" 
              size="medium"
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
            <SidenavButton onClick={() => {setsideBarChosen(0)}} isActive={homeFlag} startIcon={<SvgHome />} key="Home">
              Home
            </SidenavButton>
        ]}
        secondaryItems={[
          <SidenavButton onClick={() => {setsideBarChosen(99)}} isActive={settingsFlag} startIcon={<SvgSettings />} key="settings">
            Settings
          </SidenavButton>
        ]}
        />       
     
       <div className="app-container">
          <div className="app-content">
            {isLoggingIn ? ( <span>"Logging in...."</span> ) : (!isAuthorized ? (<h1>You need to login first</h1>) : (bodyData))}
          </div>
          <Footer />
        </div>
    </div>
  </div>
    
  );
};

export default App;
