import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
require('./test.css');
import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import "jquery"; 
import * as bootstrap from "bootstrap";
export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPartWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public render(): void {
    var url = "https://capgeminitls.sharepoint.com/sites/Cockpit-DEV";
    //var url = this.context.pageContext;
    
    this.domElement.innerHTML = `
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons"
    rel="stylesheet">
    <div>
      <table id="TableBorderColor" style='width:100%' >
        <tr BGCOLOR="#D9D9D9" style="font-size: 16px;">
          <td id ="Nameoftheproject">Name of the Project : </td>
          <td id ="Nameoftheproject2"></td>
          <td id ="Typeofproject">Type of Project : </td>
          <td id ="Typeofproject2"></td>
        </tr>
        <tr BGCOLOR="#D9D9D9" style="font-size: 16px;">
          <td id ="Nameoftheprojetcchief">Name of the Project's Chief : </td>
          <td id ="Nameoftheprojetcchief2"></td>
          <td id ="NameofthePMO">Name of the PMO : </td>
          <td id ="NameofthePMO2"></td>
          <td id ="Nameofthesponsor">Name of the Sponsor : </td>
          <td id ="Nameofthesponsor2"></td>
        </tr>
      </table>
    <div style="padding-top: 20px;">
      <table style='width:50%' align='left' class="table">
        <tr BGCOLOR="#D9D9D9" style="height: 60px;">
          <th colspan="4"  style="text-align: left; height: 60px; font-size: 22px;" >Project Indicators</th>
        </tr>
        <tr BGCOLOR="#F2F2F2">
          <td id="ColorGates" HEIGHT=100 width="16%" align=center><i class="material-icons">chevron_right</i><i class="material-icons">chevron_right</i><i class="material-icons">chevron_right</i><br />Gates<hr color="white"">
          </td>
          <td id="ColorDeliverables" HEIGHT=100 width="8%" align=center><i class="material-icons">access_time</i><br />Times<hr color="white">
          </td>
          <td id="ColorRisks" HEIGHT=100 width="8%" align=center><i class="material-icons">help_outline</i><br />Risks<hr color="white"> 
          </td>
          <td id="ColorActions" HEIGHT=100 width="8%" align=center><i class="material-icons">info_outline</i><br />Actions<hr color="white"> 
          </td>
        </tr>
      </table>
    </div>
  <div style="padding-left: 20px;">
    <table style='width:50%' align='right'>
      <tr>
        <td onclick="document.location.href='`+url+`/Lists/Deliverables'" id="deliverables" HEIGHT=230 width="30%" align=center ><FONT color="white" style="font-weight: bold; font-size: 14px;"># Pending <br />deliverables<br /> in retard</FONT><hr style="border: 1px dashed white;">
        </td>
        <td onclick="document.location.href='`+url+`/Lists/Actions'" id="actions" HEIGHT=100 width="30%" align=center><FONT color="white" style="font-weight: bold; font-size: 14px;"># Pending <br />actions<br /> in retard</FONT><hr style="border: 1px dashed white;"> 
        </td>
        <td onclick="document.location.href='`+url+`/Lists/Risks'" id="Risks" HEIGHT=100 width="30%" align=center><FONT color="white" style="font-weight: bold; font-size: 14px;"># Risks <br />not covered<br /> (Criticality 3+)</FONT><hr style="border: 1px dashed white;"> 
        </td>
      </tr>
    </table>
    </div>
    </div>
    </div>`;

    
      var Retard;
      var Criticality;
      var Id;
      $(document).ready(() => {
        $.ajax({
          // Rest Query KPI 3 List Actions
          url: this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Actions')/items?$filter=((bxpx eq null) and (Retard eq 'Retard'))",
          method: "GET",
          contentType: "application/json",
          headers: { "Accept": "application/json; odata=verbose" },
          success:function(data){
            $.each(data.d.results, function(index, item){   

              //$("#Retard").append('<div>'+item.Retard+'</div>');
              Retard = item.Retard;
            });

            if (data.d.results.length<= 2){
              //Couleur Verte Actions
              $("#actions").append('<td id="actions"><font size="12" color="white" align=center>'+data.d.results.length+'</font></td>');
              $("#actions").css("background-color","#00B050");

              $("#ColorActions").append('<td id="ColorActions" HEIGHT=100 width="10%" align=center><div id="carreVert"></div></td>'); 
              
            }
            else if (data.d.results.length>= 3 && data.d.results.length<= 4){
              //Couleur Orange Actions
              $("#actions").append('<td id="actions"><font size="12" color="white" align=center>'+data.d.results.length+'</font></td>');
              $("#actions").css("background-color","#ffc000"); 

              $("#ColorActions").append('<td id="ColorActions" HEIGHT=100 width="10%" align=center><div id="carreOrange"></div></td>');
              
            }
            else if (data.d.results.length>5){
              //Couleur Rouge Actions
              $("#actions").append('<td id="ColorActions"><font size="12" color="white" align=center>'+data.d.results.length+'</font></td>');
              $("#actions").css("background-color","#ff0000");

              $("#ColorActions").append('<td id="ColorActions" HEIGHT=100 width="10%" align=center><div id="carreRouge"></div></td>');            }
              
          },
          error: function () {
            console.log("erreur Actions");
          }
         
            });
       
          $.ajax({
            // Rest Query KPI 1 List Delivrables
            url: this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Deliverables')/items?$filter=((Realdateofdelivery eq null) and (Retard eq 'Retard'))",
            method: "GET",
            contentType: "application/json",
            headers: { "Accept": "application/json; odata=verbose" },
            success:function(data){
              $.each(data.d.results, function(index, item){   

                Retard = item.Retard;
              });
              
              if (data.d.results.length<= 1){
                //Couleur Verte deliverables
                $("#deliverables").append('<td id="deliverables"><font size="12" color="white">'+data.d.results.length+'</font></td>');
                $("#deliverables").css("background-color","#00B050");

                $("#ColorDeliverables").append('<td id="ColorDeliverables" HEIGHT=100 width="10%" align=center><div id="carreVert"></div></td>');
              }
              else if (data.d.results.length>= 2 && data.d.results.length<= 3){
                //Couleur Orange deliverables
                $("#deliverables").append('<td id="deliverables"><font size="12" color="white">'+data.d.results.length+'</font></td>');
                $("#deliverables").css("background-color","#ffc000"); 
  
                $("#ColorDeliverables").append('<td id="ColorDeliverables" HEIGHT=100 width="10%" align=center><div id="carreOrange"></div></td>');

              }
              else if (data.d.results.length>4){
                //Couleur Rouge deliverables
                $("#deliverables").append('<td id="deliverables"><font size="12" color="white">'+data.d.results.length+'</font></td>');
                $("#deliverables").css("background-color","#ff0000");

                $("#ColorDeliverables").append('<td id="ColorDeliverables" HEIGHT=100 width="10%" align=center><div id="carreRouge"></div></td>');
  
              }

            },
            error: function () {
              console.log("erreur Delivrables");
            }
            
              });
            

              $.ajax({
                // Rest Query KPI 2 List Risks
                url: this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Risks')/items?$filter=((Status eq 'Not covered') and (Criticality ge '3'))",
                method: "GET",
                contentType: "application/json",
                headers: { "Accept": "application/json; odata=verbose" },
                success:function(data){
                  $.each(data.d.results, function(index, item){   
    
                    Criticality = item.Criticality;
                    //Id = item.Id;
                    //console.log(Id);
                  });
                  
                  //console.log(data.d.results.length);
                  
                  if (data.d.results.length == 0){
                    //Couleur Verte Risks
                    $("#Risks").append('<td id="Risks"><font size="12" color="white">'+data.d.results.length+'</font></td>');
                    $("#Risks").css("background-color","#00B050");

                    $("#ColorRisks").append('<td id="ColorRisks" HEIGHT=100 width="10%" align=center><div id="carreVert"></div></td>');            
                  }
                  else if (data.d.results.length == 1){
                    //Couleur Orange Risks
                    $("#Risks").append('<td id="Risks"><font size="12" color="white">'+data.d.results.length+'</font></td>');
                    $("#Risks").css("background-color","#ffc000"); 
                     
                    $("#ColorRisks").append('<td id="ColorRisks" HEIGHT=100 width="10%" align=center><div id="carreOrange"></div></td>');                    
                  }
                  else if (data.d.results.length>=2){
                    //Couleur Rouge Risks
                    $("#Risks").append('<td id="Risks"><font size="12" color="white">'+data.d.results.length+'</font></td>');
                    $("#Risks").css("background-color","#ff0000");

                    $("#ColorRisks").append('<td id="ColorRisks" HEIGHT=100 width="10%" align=center><div id="carreRouge"></div></td>');
                    
                  }

                },
                error: function () {
                  console.log("erreur Risks");
                }
                
                  });
                  var Phase;
                  var PhaseNbr;
                  var Status;
                  
                  $.ajax({
                    // Rest Query Phase delivrable
                    url: this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Deliverables')/items",
                    method: "GET",
                    contentType: "application/json",
                    headers: { "Accept": "application/json; odata=verbose" },
                    success:function(data){
                      $.each(data.d.results, function(index, item){   
        
                        Phase = item.Phase;
                        Status = item.Status;

                        if (Phase == "1 Launch" && Status != "Validated"){
                          PhaseNbr = "Phase1";

                          return false;
                        }
                        if(Phase == "2 Analysis and conception" && Status != "Validated"){
                          PhaseNbr = "Phase2";

                          return false;
                        }
                        if(Phase == "3 Development" && Status != "Validated"){
                          PhaseNbr = "Phase3";
                          return false;
                        }
                        

                       
                        
                      });
                      
                      if (PhaseNbr == "Phase1"){
                        //Phase 1
                        $("#ColorGates").append('<td id="ColorGates" HEIGHT=100 width="10%" align=center><img src="'+url+'/img/Phase1.png" /><br />Jalon 1</td>');
                      }
                      else if (PhaseNbr == "Phase2"){
                        //Phase 2
                        $("#ColorGates").append('<td id="ColorGates" HEIGHT=100 width="10%" align=center><img src="'+url+'/img/Phase2.png" /><br />Jalon 2</td>');
                      }
                      else if (PhaseNbr == "Phase3"){
                        //Phase 3
                        $("#ColorGates").append('<td id="ColorGates" HEIGHT=100 width="10%" align=center><img src="'+url+'/img/Phase3.png" /><br />Jalon 3</td>');

                      }
    
                    },
                    error: function () {
                      console.log("erreur Risks");
                    }
                    
                      });


                  $.ajax({
                    // Rest Query List Project Name
                    url: this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('Project name')/items?$filter=(ID eq  '1')",
                    method: "GET",
                    contentType: "application/json",
                    headers: { "Accept": "application/json; odata=verbose" },
                    success:function(data){
                      $.each(data.d.results, function(index, item){   
                        $("#Nameoftheproject2").append(item.Nameoftheproject);
                        $("#Typeofproject2").append(item.Typeofproject);
                        $("#Nameoftheprojetcchief2").append(item.Nameoftheprojetcchief);
                        $("#NameofthePMO2").append(item.NameofthePMO);
                        $("#Nameofthesponsor2").append(item.Nameofthesponsor); 
                      });
                      
                    },
                    error: function () {
                      console.log("erreur Risks");
                    }
                    
                      });
                  
           
          });
          
        };
 




  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
