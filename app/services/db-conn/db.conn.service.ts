

/// 

import { Injectable } from '@angular/core';
import { Http,Headers,RequestOptions } from '@angular/http';

import 'rxjs/add/operator/map';
declare var jquery:any;
declare var $:any;

@Injectable()
export class DbConnService { 
    
    constructor(private http:Http) {}

    getAgents(){
      
        //return this.http.get('https://www.kansanmedtrip.com/getData.php?module=treatment').map(res=>res.json());
    
    }


    getFiles(token:any){
        
        let headers = new Headers({
            "Authorization":"Bearer "+ token,
            "Accept":"application/json;odata=verbose",
            "Access-Control-Allow-Origin":"https://localhost:3000/settings/"
        });

        var options:any = new RequestOptions({
            headers:headers,
            withCredentials:true
        })
        console.log("get files called");
        return this.http.get("https://progressivedigital.sharepoint.com/_api/web/lists/GetByTitle('client')/items",options).map(res=>{console.log("Response  "+res.json())}).subscribe(res=>{console.log(res)});
    }

    getFilesFromSharepoint(token:any){
        try 
        { 
          this.isCorsCompatible(token);
          var endpointUrl = "https://progressivedigital.sharepoint.com/_api/web/lists/GetByTitle('client')/items"; 
          var xhr = new XMLHttpRequest(); 
          xhr.open("GET", endpointUrl); 
            
          // The APIs require an OAuth access token in the Authorization header, formatted like this: 'Authorization: Bearer <token>'. 
          xhr.setRequestHeader("Authorization", "Bearer " + token);
          xhr.setRequestHeader("accept", "application/json;odata=verbose");
          xhr.setRequestHeader("Access-control-allow-origin","*"); 
          
          // Process the response from the API.  
          xhr.onload = function () { 
            if (xhr.status == 200) { 
              var formattedResponse = JSON.stringify(JSON.parse(xhr.response), undefined, 2);
              console.log(formattedResponse) 
            } else { 
            //   document.getElementById("results").textContent = "HTTP " + xhr.status + "<br>" + xhr.response; 
            } 
          } 
      
          // Make request.
          xhr.send(); 
        } 
        catch (err) 
        {  
        //   document.getElementById("results").textContent = "Exception: " + err.message; 
        console.log(err);
        } 
    }

    isCorsCompatible = function(token:any) {
        try
        {
            var xhr = new XMLHttpRequest();
            xhr.open("GET","https://progressivedigital.sharepoint.com/_api/web/lists/GetByTitle('client')/items");
            xhr.setRequestHeader("authorization", "Bearer " + token);
            xhr.setRequestHeader("accept", "application/json");
            xhr.onload = function () {
               // CORS is working.
               console.log("Browser is CORS compatible."); 
            }
            xhr.send();
        }
        catch (e)
        {
            if (e.number == -2147024891)
            {
                console.log("Internet Explorer users must use Internet Explorer 11 with MS15-032: Cumulative security update for Internet Explorer (KB3038314) installed for this sample to work.");
            }
            else
            {
                console.log("An unexpected error occurred. Please refresh the page."); 
            }

        }
      };

    
    
}
