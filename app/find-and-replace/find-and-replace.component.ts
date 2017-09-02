// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

/*
  This file defines a component that enables a search-and-replace functionality for
  the Word document. 
*/

import { Component, } from '@angular/core';
import { Router,ActivatedRoute } from '@angular/router';

import { FabricTextFieldWrapperComponent } from '../shared/office-fabric-component-wrappers/fabric.textfield.wrapper.component';
import { ButtonComponent } from '../shared/button/button.component';
import { NavigationHeaderComponent } from '../shared/navigation-header/navigation.header.component';
import { BrandFooterComponent } from '../shared/brand-footer/brand.footer.component';

// The WordDocumentService provides methods for manipulating the document.
import { WordDocumentService } from '../services/word-document/word.document.service';

// The SettingsStorageService provides CRUD operations on application settings..
import { SettingsStorageService } from '../services/settings-storage/settings.storage.service';
import { DbConnService } from '../services/db-conn/db.conn.service';


@Component({
    templateUrl: 'app/find-and-replace/find-and-replace.component.html',
    styleUrls: ['app/find-and-replace/find-and-replace.component.css'],
})
export class FindAndReplaceComponent {

    private searchString: string;
    private replaceString: string;
    private excludedParagraph: number;
    private subscription: any;
    data: Boolean;
    loadControls:Boolean;
    controls:Array<any>;
    keys:Array<any>;

    agents: [{
        "id": number,
        "tName": string,
        "tLink": string,

    }];

    customers:[{
        "company":string,
        "email":string,
        "phone":string,
        "address":string

    }];

    constructor(private wordDocument: WordDocumentService,
        private settingsStorage: SettingsStorageService,
        private router: Router, private dbConn: DbConnService,public actroute:ActivatedRoute) {
        this.data = true;
        console.log("Bahar");
        this.wordDocument.getAgents().subscribe((data)=>{
            console.log("Shubham");
            console.log(data);
        });
        this.customers=[
            {
              "company": "EMTRAC",
              "email": "patcantrell@emtrac.com",
              "phone": "+1 (869) 531-3434",
              "address": "626 Hawthorne Street, Beaverdale, Palau, 546"
            },
            {
              "company": "EXODOC",
              "email": "patcantrell@exodoc.com",
              "phone": "+1 (858) 467-3380",
              "address": "178 Monroe Place, Tyhee, South Carolina, 6293"
            },
            {
              "company": "ZOMBOID",
              "email": "patcantrell@zomboid.com",
              "phone": "+1 (923) 553-3570",
              "address": "265 Granite Street, Nelson, Wyoming, 345"
            },
            {
              "company": "MELBACOR",
              "email": "patcantrell@melbacor.com",
              "phone": "+1 (960) 444-3049",
              "address": "562 Harbor Lane, Brambleton, Nebraska, 1270"
            },
            {
              "company": "CEDWARD",
              "email": "patcantrell@cedward.com",
              "phone": "+1 (943) 402-2747",
              "address": "809 Leonard Street, Clay, Puerto Rico, 6673"
            },
            {
              "company": "HOMELUX",
              "email": "patcantrell@homelux.com",
              "phone": "+1 (866) 508-2407",
              "address": "137 Ovington Avenue, Greensburg, California, 8205"
            }
          ];
    }

    ngOnInit(){
        let route = window.location.href;
        
        console.log(route);
        this.splitQueryString(route);
       
    }

    // Handle the event of a user entering text in the search box.
    onSearchTextEntered(message: string): void {
        this.searchString = message;
    }

    // Handle the event of a user entering text in the replace box.
    onReplaceTextEntered(message: string): void {
        this.replaceString = message;
    }

    // Handle the event of a user entering a number in the box for excluded paragraphs.
    onParagraphNumeralEntered(message: number): void {
        this.excludedParagraph = message;
    }

    replace(): void {
        this.wordDocument.replaceFoundStringsWithExceptions(this.searchString, this.replaceString, this.excludedParagraph);
    }

    loadData() {
        this.data = false;
    }

    pushData(){
        var agnts:Array<string>;
        agnts =[];
        for(var i=0;i<this.agents.length;i++){
            agnts[i] = this.agents[i].tName;
        }
        console.log(this.agents);
        console.log(agnts);
        this.wordDocument.replaceDocumentContent(agnts);
    }

    loadControl(){
        this.loadControls  = true;
        this.wordDocument.loadingContentControl();
        
    }

    createNew(){
        this.wordDocument.createNewPlaceholder();
    }
    
    loadCustomerData(customer:any){
        
        for(var i=0;i<this.customers.length;i++){
            if(this.customers[i].company == customer){
                
                this.keys=[];
                for(var j in this.customers[i]){
                    this.keys.push(j);
                    
                }
                this.wordDocument.populateData(this.customers[i],this.keys);
                
            }
        }
        
    }

    loadOoxml(){
        this.wordDocument.loadOoxml();
    }

    splitQueryString(queryStringFormattedString:string){
        var str = queryStringFormattedString.toString();
        var codeIndex = str.indexOf("code=");
        var sessionIndex = str.indexOf("session");
        var token = str.substring((codeIndex+5),(sessionIndex-1));
        console.log(token);
        this.getFilesFromO365(token);
             
    }

    printSomething(){
        console.log("Something")
    }


    getFilesFromO365(token:any) { 
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
