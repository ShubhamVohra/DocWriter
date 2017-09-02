// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

/*
  This file defines an instructions component for a task pane page. It is based on
  the instruction-step sample, created by the Modern Assistance Experience Developer 
  Docs team. Along with other samples, it is in the Office-Add-in-UX-Design-Patterns-Code 
  repo:  https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code
*/

import { Component } from '@angular/core';
import { Router,ActivatedRoute } from '@angular/router';

import { ButtonComponent } from '../shared/button/button.component';
import { IInstructionStep } from './IInstructionStep';
var queryStringParameters:any;
@Component({
    templateUrl: 'app/instructions/instruction-steps.component.html',
    styleUrls: ['app/instructions/instruction-steps.component.css']
})
export class InstructionStepsComponent {
    
    private title: string = "EY TEMPLATE DESIGNER";
    
    public token:any;
    private addin_description: string = "Template Designer enables you to enforce style rules while exempting paragraphs that you specify from the rules.";
    private steps_intro: string = "Just take these steps:";
    private steps: Array<IInstructionStep> =
    [{ step_number: 1, content: "Enter a string in the Find box." },
        { step_number: 2, content: "Enter a replacement string in the Replace With box." },
        { step_number: 3, content: "Enter the zero-based numbers of the parapgraphs that should be exempt in the Skip Paragraphs box." },
        { step_number: 4, content: "Press Replace." }];

    constructor(private router: Router,private route:ActivatedRoute) { }

    private parseQueryString = function(url:any) {
		var params = {}, queryString = url.substring(1),
		regex = /([^&=]+)=([^&]*)/g, m;
		while (m = regex.exec(queryString)) {
			params[decodeURIComponent(m[1])] = decodeURIComponent(m[2]);
		}
		return params;
	}
    private params = this.parseQueryString(location.hash);
    public access_token:string = null;
    
    ngOnInit(){
        let urout = this.router.url;
        console.log(urout);
        console.log(window.location.pathname);
        
        var clientId    = '44d3241c-4d7a-439b-a37c-df8f6e8d689d';
        var replyUrl    = 'https://10.80.0.236:3000/settings/'; 
        var endpointUrl = 'https://progressivedigital.sharepoint.com/_api/v1.0/me/files';
        var resource = "https://progressivedigital.sharepoint.com"; 
      
        var authServer  = 'https://login.windows.net/common/oauth2/authorize?';  
        var responseType = 'token'; 
      
        var url = authServer + 
                  "response_type=" + encodeURI(responseType) + "&" + 
                  "client_id=" + encodeURI(clientId) + "&" + 
                  "resource=" + encodeURI(resource) + "&" + 
                  "redirect_uri=" + encodeURI(replyUrl) + "&"+ "state=" +12345  ; 
            console.log(url);
            // alert(url);
        window.open(url,"_self");
        var token=this.params['access_token'];
        var a =this.splitQueryString(url);
        console.log(token);

    }
     
      
    splitQueryString(queryStringFormattedString:any){
          var split = queryStringFormattedString.split('&');
          if(split == ""){
              return {};
          }

          var results = {};

          for (var i=0 ;i<split.length ;i++){
              var p = split[i].split("=",2);

              if(p.length ==1){
                  results[p[0]]="";
              }
              else{
                  results[p[0]] = decodeURIComponent(p[1].replace(/\+/g, " "));
              }
          }
          return results;
               
      }

      printSomething(){
          console.log("Something")
      }

        
}

