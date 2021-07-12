import { IMSLoginProper } from "./IMSLoginProper";

export class ProjectData{
    public static async getRecentProjects(){
        const accessToken = await IMSLoginProper.getAccessTokenForBentleyAPI();
          const response = await fetch("https://api.bentley.com/projects/recents", {
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': accessToken,
                  //  'Access-Control-Allow-Origin' : 'http://localhost:3000',
                  },
                })
            //const dataRaw = await response.text(); // work around for invalid newline escaping that is present in some public api
            //const json = JSON.parse(dataRaw);
            // normally we would just do below but Bentley API is a little unsafe *2021/07/01*
            const data = await response;
            const json = await data.json();
            var info: { id: string; displayName: string; projectNumber: string; }[] = [];
            for (var i = 0; i < json.projects.length; i++)
            {
              info.push({id: json.projects[i].id, displayName: json.projects[i].displayName , projectNumber: json.projects[i].projectNumber });
            }
            return  info;
    }

    public static async getProjectData(projectId: string){
        const accessToken = await IMSLoginProper.getAccessTokenForBentleyAPI();
        const response = await fetch("https://api.bentley.com/projects/" + projectId, { //https://api.bentley.com/projects/favorites?top=1000
            mode: 'cors',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': accessToken,
                // 'Access-Control-Allow-Origin' : 'http://localhost:3000',
                },
            })
        const data = await response;
        const json = await data.json();
        const info = ({id:projectId, displayName:json.project.displayName, projectNumber:json.project.projectNumber})
        return  info;
      }

      public static async getFormTemplatesFromProject(projectId: string){
          const accessToken = await IMSLoginProper.getAccessTokenForBentleyAPI();
          const response = await fetch("https://api.bentley.com/issues/formDefinitions?projectId=" + projectId, { 
          mode: 'cors',
          headers: {
              'Content-Type': 'application/json',
              'Authorization': accessToken,
              //'Access-Control-Allow-Origin' : 'http://localhost:3000',
            },
          })
      const data = await response; 
      const json = await data.json();
      var info: { id: string; formId: string; displayName: string; type: string; }[] = [];
      for (var i = 0; i < json.formDefinitions.length; i++)
      {
        info.push({id:projectId, formId:json.formDefinitions[i].id ,displayName:json.formDefinitions[i].displayName, type:json.formDefinitions[i].type})
      }
      return info;
      }

      public static async getFormDetailsFromId(formId: string){
          const accessToken = await IMSLoginProper.getAccessTokenForBentleyAPI();
          const response = await fetch("https://api.bentley.com/issues/formDefinitions/" + formId, { //GET https://api.bentley.com/issues/formDefinitions/{id}
                mode: 'cors',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': accessToken,
                   // 'Access-Control-Allow-Origin' : 'http://localhost:3000',
                  },
                })
            const data = await response;
            const json = await data.json();
           // console.log("formId",formId)
           // console.log(json);
            const info = ({id:formId, displayName:json.formDefinition.displayName, type:json.formDefinition.type, status:json.formDefinition.status});
            return  info;
      }

}