// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo.

/*
  This file defines a service for manipulating the Word document. 
*/

/// <reference path="../../../typings/index.d.ts" />

import { Injectable } from '@angular/core';

import { IReplacementCandidate } from './IReplacementCandidate';
import { Http,Headers } from '@angular/http';
import 'rxjs/add/operator/map';



@Injectable()
export class WordDocumentService {

    /// <summary>
    /// Performs a search and replace, but makes no changes to text in the excluded paragraphs.
    /// </summary>
    cust:any;
    controls:Array<any>;
    keys:Array<string>;
    constructor(public http:Http){
        
    }

    loadingContentControl(){
        
        Word.run(function(context){
            let placeholder:Word.ContentControl;
            
            var document = context.document;
            var app = context.application.context;
            var body = document.body;
            var contentControls = document.contentControls;
            
            //placeholder.appearance.ti = "BoundingBox";
            
            contentControls.load('tag,title');
            
            
            return context.sync()
            .then(function(){
                for(var i=0;i<contentControls.items.length;i++){
                    contentControls.items[i].insertText("Shubham",Word.InsertLocation.replace);
                    
                }
                context.load(body);
                
            });

        }).catch(this.errorHandler);
    }


    populateData(customer:any,keys:Array<any>){
        Word.run(function(context){
            let placeholder:Word.ContentControl;
            var document = context.document;
            
            var body = document.body;
            var contentControls = document.contentControls;
            
            contentControls.load('tag,title');
            
            
            return context.sync()
            .then(function(){

                // for(var i=0;i<contentControls.items.length;i++){
                //     contentControls.items[i].insertText(keys.length.toString(),Word.InsertLocation.replace);
                    
                // }

                // body.insertText(keys[0],"End");
                // context.load(body);
                var i,j;
                for (i = 0; i < keys.length; i++) {
                    for (j = 0; j < contentControls.items.length; j++) {

                        // Matching content control tag with the tag set as the id on each input element.
                        // Set the content text to the text value of the INPUT element.
                        if (contentControls.items[j].title === keys[i]) {
                             var shubh = keys[i];
                            contentControls.items[j].insertText(customer[shubh], Word.InsertLocation.replace)
                        }
                    }
                }
                
            });

        }).catch(this.errorHandler);
    }

    loadOoxml(){
        Word.run(function (context) {
            
            // Create a proxy object for the content controls collection.
            var contentControls = context.document.contentControls;
            var body = context.document.body;
            
            // Queue a command to load the id property for all of the content controls. 
            context.load(contentControls, 'id');
             
            // Synchronize the document state by executing the queued-up commands, 
            // and return a promise to indicate task completion.
            return context.sync().then(function () {
                if (contentControls.items.length === 0) {
                    console.log('No content control found.');
                    
                }
                else {
                    // Queue a command to get the OOXML contents of the first content control.
                    //var ooxml = contentControls.items[0].getOoxml();
                    
                    contentControls.items[0].insertOoxml("<w:p xmlns:w='http://schemas.microsoft.com/office/word/2003/wordml'><w:r><w:rPr><w:b/><w:b-cs/><w:color w:val='FF0000'/><w:sz w:val='28'/><w:sz-cs w:val='28'/></w:rPr><w:t>Hello world (this should be bold, red, size 14).</w:t></w:r></w:p>", "End");
                    // Synchronize the document state by executing the queued-up commands, 
                    // and return a promise to indicate task completion.
                    return context.sync()
                        .then(function () {
                            
                    });
                }
                context.load(body);
            });  
        })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
    }
    

    createNewPlaceholder(){
        Word.run(function (context) {
            
                // Create a proxy object for the document body.
                var body = context.document.body;
                var contents = body.contentControls;
                var range = context.document.getSelection();
                var mycontrol = range.insertContentControl();

                
                // Queue a commmand to wrap the body in a content control.
                mycontrol.tag = "Today's Date";
                mycontrol.title = "Enter today's date:";
                //mycontrol.insertText(mycontrol.tag,"Replace")
                mycontrol.appearance = 'BoundingBox';
                mycontrol.color = "gray";
                context.load(mycontrol,'tag');
            
                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                   // context.load(body);
                    console.log('Wrapped the body in a content control.');
                });
            })
            .catch(function (error) {
                
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }

    writeContent(fileName:any) {
            console.log("Shubham");
            
        
    }

    replaceFoundStringsWithExceptions(searchString: string, replaceString: string, excludedParagraph: number
                                      ) {

        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            let http:Http;
            // Find and load all ranges that match the search string, and then all paragraphs in the document.
            // Only the 'items' property of each is needed, no properties on the items are needed, so add any string 
            // after the 'items/' part of the load parameter.
            let foundItems: Word.SearchResultCollection = context.document.body.search(searchString, { matchCase: false, matchWholeWord: true }).load('items/NoPropertiesNeeded');
            let paras : Word.ParagraphCollection = context.document.body.paragraphs.load('items/NoPropertiesNeeded');

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync()

            .then(function () {          

                // Create an array of paragraphs that have been excluded.
                let excludedRanges: Array<Word.Range> = [];
                excludedRanges.push(paras.items[excludedParagraph].getRange('Whole'));

                let replacementCandidates : Array<IReplacementCandidate> = [];

                // For each instance of the search string, record whether or not it is in an
                // excluded paragraph.
                for (let i = 0; i < foundItems.items.length; i++) {
                    for (let j = 0; j < excludedRanges.length; j++) {                 
                        replacementCandidates.push({
                            range: foundItems.items[i],
                            locationRelation: foundItems.items[i].compareLocationWith(excludedRanges[j])
                        });
                    }
                }
                // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
                return context.sync()
                
                .then(function () {

                    // Replace instances of the search string with the replace string only if they are
                    // not inside of (or identical to) an excluded range.
                    replacementCandidates.forEach(function (item) {

                        switch (item.locationRelation.value) {
                            case "Inside":
                            case "Equal":
                                break;
                            default:
                                item.range.insertText(replaceString, 'Replace');
                        }
                    });
                });
            });
        })
        .catch(this.errorHandler);


    }
    
    
    /// <summary>
    /// Inserts sample content for testing the find-and-replace functionality..
    /// </summary>
    replaceDocumentContent(paragraphs: Array<any>) {
        
        // Run a batch operation against the Word object model.
        Word.run(function (context) {

            // Create a proxy object for the document body.
            let body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            body.clear();

            // Queue commands to insert text into the end of the Word document body.
            // Use insertText for the first to prevent a line break from being inserted 
            // at the top of the document.
            body.insertText(paragraphs[0], "End");

            // Use insertParagrpah for all the others.
            for (let i=1; i < paragraphs.length; i++) {
                 body.insertParagraph(paragraphs[i], 'End');
            }
           
            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
        .catch(this.errorHandler);
    }

    getAgents(){
        var headers = new Headers;
        var accessToken = '0x5EB8690626ED38C5C334D0ACC34CBB48175641C184CFE308EDE550D6F03D8869D3393F4E81919CE3D8713866475DBAF4252E921781B77DF0FD0DEC3638FE3380,19 Aug 2017 11:20:20 -0000';
        //var accessToken = "rtFa=WORLS/jbMVbB+y3AYpoQYwKtGnsPk4TkLe9SLymsKaFabRax8jDGW+c+cSMsWc+d53CVN0hAujI+eaWJbSKE0V4BWc1MqDSsntl28Y2WnVBxj+wYbjB+edUplP3xV4dYuADBpiX4Ql33wb2aDAYGRsxc7mW0CZtWbr+i/fLPupSzS18DGLisY2mJHws8162VH6AGH56/4+fH9/oBp/1bTm5lDERU6PNm1c30X2prRnNtNHePF6HMII1D/3SGDx60nhmk8iJWdtr+xtfdUyvUnLqS8eEz8rbFVEPMbjHzWHDv/lCfTXy4/Kx9jEQ2PtKuJ0x1yDTY2C5udxj65q/E/Y7Eq3cQbkNl6+iJMoo3c5KClUxP4UiNCgFbHpRJ7qmUIAAAAA==;FedAuth=77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48U1A+VjMsMGguZnxtZW1iZXJzaGlwfDEwMDMzZmZmYTQyMDk3NTVAbGl2ZS5jb20sMCMuZnxtZW1iZXJzaGlwfHNodWJoYW1AcHJvZ3Jlc3NpdmVkaWdpdGFsLm9ubWljcm9zb2Z0LmNvbSwxMzE0NzYwNzQzOTAwMDAwMDAsMTMxNDc1MDY3MjYwMDAwMDAwLDEzMTQ4MDQzNTY4NDIyMzE0NiwwLjAuMC4wLDIsMGUxNWIxNjctMGY4Zi00MmVhLTgyYmItMTQzY2RlNThkZTAxLCwsZTBjMjEwOWUtOTA4YS00MDAwLWVlNGMtNTIyZWFmM2JlNTg1LGUwYzIxMDllLTkwOGEtNDAwMC1lZTRjLTUyMmVhZjNiZTU4NSwsMCxmS0xtdmk1Wk1FYWx5UmJIVTVkOTF2Q1ZuZzI2bFdtNEtDeGtkOGc1Z2tGcmtrbUZtb2VxdnpuN2dXcHBjNzdmbndJbG5DVVRjTTZ6NU1aM2RoZHJITmdMeXM4cmRXRHdreW1HT0VjSFFESlRTZVplK0dqblg4REt1T2p5Qlk4MHRtQVZpaFNkM2xZVU5BUjBEZWdjcHFUL0E4c2Y2d0lhNGEzQTRlOTlaR1E5ZVVCOFB2dnk4dWc2U1BMRTBLOXdpbExnbGVDUmhMTjZoVVVpNkxPQmIrZ0xjK05YcDQ3RVIva0p0aEtDNFdoTU96TzEySDdxZXQrM3hvRWtSMStTYng0dFJOVUJsK1huSExwYlJTY1hsLzFvcUdqT0gyQ29kVmdFbG9US0NGYkUxZHRrdmtWb1l5WTNuR3ZoYysrUUVsME5ZU2JQSFFrQWlJTnpHRDdIRHc9PTwvU1A+;FedAuth=77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48U1A+VjMsMGguZnxtZW1iZXJzaGlwfDEwMDMzZmZmYTQyMDk3NTVAbGl2ZS5jb20sMCMuZnxtZW1iZXJzaGlwfHNodWJoYW1AcHJvZ3Jlc3NpdmVkaWdpdGFsLm9ubWljcm9zb2Z0LmNvbSwxMzE0NzYwNzQzOTAwMDAwMDAsMTMxNDc1MDY3MjYwMDAwMDAwLDEzMTQ4MDQzNTY4NDIyMzE0NiwwLjAuMC4wLDIsMGUxNWIxNjctMGY4Zi00MmVhLTgyYmItMTQzY2RlNThkZTAxLCwsZTBjMjEwOWUtOTA4YS00MDAwLWVlNGMtNTIyZWFmM2JlNTg1LGUwYzIxMDllLTkwOGEtNDAwMC1lZTRjLTUyMmVhZjNiZTU4NSwsMCxmS0xtdmk1Wk1FYWx5UmJIVTVkOTF2Q1ZuZzI2bFdtNEtDeGtkOGc1Z2tGcmtrbUZtb2VxdnpuN2dXcHBjNzdmbndJbG5DVVRjTTZ6NU1aM2RoZHJITmdMeXM4cmRXRHdreW1HT0VjSFFESlRTZVplK0dqblg4REt1T2p5Qlk4MHRtQVZpaFNkM2xZVU5BUjBEZWdjcHFUL0E4c2Y2d0lhNGEzQTRlOTlaR1E5ZVVCOFB2dnk4dWc2U1BMRTBLOXdpbExnbGVDUmhMTjZoVVVpNkxPQmIrZ0xjK05YcDQ3RVIva0p0aEtDNFdoTU96TzEySDdxZXQrM3hvRWtSMStTYng0dFJOVUJsK1huSExwYlJTY1hsLzFvcUdqT0gyQ29kVmdFbG9US0NGYkUxZHRrdmtWb1l5WTNuR3ZoYysrUUVsME5ZU2JQSFFrQWlJTnpHRDdIRHc9PTwvU1A+";
        headers['Content-Type']= "application/json;odata=verbose";
        //headers.append('Cookie','rtFa=WORLS/jbMVbB+y3AYpoQYwKtGnsPk4TkLe9SLymsKaFabRax8jDGW+c+cSMsWc+d53CVN0hAujI+eaWJbSKE0V4BWc1MqDSsntl28Y2WnVBxj+wYbjB+edUplP3xV4dYuADBpiX4Ql33wb2aDAYGRsxc7mW0CZtWbr+i/fLPupSzS18DGLisY2mJHws8162VH6AGH56/4+fH9/oBp/1bTm5lDERU6PNm1c30X2prRnNtNHePF6HMII1D/3SGDx60nhmk8iJWdtr+xtfdUyvUnLqS8eEz8rbFVEPMbjHzWHDv/lCfTXy4/Kx9jEQ2PtKuJ0x1yDTY2C5udxj65q/E/Y7Eq3cQbkNl6+iJMoo3c5KClUxP4UiNCgFbHpRJ7qmUIAAAAA==;FedAuth=77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48U1A+VjMsMGguZnxtZW1iZXJzaGlwfDEwMDMzZmZmYTQyMDk3NTVAbGl2ZS5jb20sMCMuZnxtZW1iZXJzaGlwfHNodWJoYW1AcHJvZ3Jlc3NpdmVkaWdpdGFsLm9ubWljcm9zb2Z0LmNvbSwxMzE0NzYwNzQzOTAwMDAwMDAsMTMxNDc1MDY3MjYwMDAwMDAwLDEzMTQ4MDQzNTY4NDIyMzE0NiwwLjAuMC4wLDIsMGUxNWIxNjctMGY4Zi00MmVhLTgyYmItMTQzY2RlNThkZTAxLCwsZTBjMjEwOWUtOTA4YS00MDAwLWVlNGMtNTIyZWFmM2JlNTg1LGUwYzIxMDllLTkwOGEtNDAwMC1lZTRjLTUyMmVhZjNiZTU4NSwsMCxmS0xtdmk1Wk1FYWx5UmJIVTVkOTF2Q1ZuZzI2bFdtNEtDeGtkOGc1Z2tGcmtrbUZtb2VxdnpuN2dXcHBjNzdmbndJbG5DVVRjTTZ6NU1aM2RoZHJITmdMeXM4cmRXRHdreW1HT0VjSFFESlRTZVplK0dqblg4REt1T2p5Qlk4MHRtQVZpaFNkM2xZVU5BUjBEZWdjcHFUL0E4c2Y2d0lhNGEzQTRlOTlaR1E5ZVVCOFB2dnk4dWc2U1BMRTBLOXdpbExnbGVDUmhMTjZoVVVpNkxPQmIrZ0xjK05YcDQ3RVIva0p0aEtDNFdoTU96TzEySDdxZXQrM3hvRWtSMStTYng0dFJOVUJsK1huSExwYlJTY1hsLzFvcUdqT0gyQ29kVmdFbG9US0NGYkUxZHRrdmtWb1l5WTNuR3ZoYysrUUVsME5ZU2JQSFFrQWlJTnpHRDdIRHc9PTwvU1A+');
        //headers.append('Cookie','FedAuth=77u/PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0idXRmLTgiPz48U1A+VjMsMGguZnxtZW1iZXJzaGlwfDEwMDMzZmZmYTQyMDk3NTVAbGl2ZS5jb20sMCMuZnxtZW1iZXJzaGlwfHNodWJoYW1AcHJvZ3Jlc3NpdmVkaWdpdGFsLm9ubWljcm9zb2Z0LmNvbSwxMzE0NzYwNzQzOTAwMDAwMDAsMTMxNDc1MDY3MjYwMDAwMDAwLDEzMTQ4MDQzNTY4NDIyMzE0NiwwLjAuMC4wLDIsMGUxNWIxNjctMGY4Zi00MmVhLTgyYmItMTQzY2RlNThkZTAxLCwsZTBjMjEwOWUtOTA4YS00MDAwLWVlNGMtNTIyZWFmM2JlNTg1LGUwYzIxMDllLTkwOGEtNDAwMC1lZTRjLTUyMmVhZjNiZTU4NSwsMCxmS0xtdmk1Wk1FYWx5UmJIVTVkOTF2Q1ZuZzI2bFdtNEtDeGtkOGc1Z2tGcmtrbUZtb2VxdnpuN2dXcHBjNzdmbndJbG5DVVRjTTZ6NU1aM2RoZHJITmdMeXM4cmRXRHdreW1HT0VjSFFESlRTZVplK0dqblg4REt1T2p5Qlk4MHRtQVZpaFNkM2xZVU5BUjBEZWdjcHFUL0E4c2Y2d0lhNGEzQTRlOTlaR1E5ZVVCOFB2dnk4dWc2U1BMRTBLOXdpbExnbGVDUmhMTjZoVVVpNkxPQmIrZ0xjK05YcDQ3RVIva0p0aEtDNFdoTU96TzEySDdxZXQrM3hvRWtSMStTYng0dFJOVUJsK1huSExwYlJTY1hsLzFvcUdqT0gyQ29kVmdFbG9US0NGYkUxZHRrdmtWb1l5WTNuR3ZoYysrUUVsME5ZU2JQSFFrQWlJTnpHRDdIRHc9PTwvU1A+');
        // //headers.append('Origin','""');
        headers.append("Accept","application/json;odata=verbose");
        //headers['Accept']= "application/json;odata=verbose";
        // //headers.append('Content-Length','0');
        headers.append("Authorization","Bearer"+accessToken);
        // headers['Authorization']
        return this.http.get("https://progressivedigital.sharepoint.com/_api/web/lists/GetByTitle('client')/items",{headers:headers}).map(res=>res.json());
        // return context.sync();
        //return this.http.get("http://progserpsrv2:3000/_api/web/lists/getByTitle('GroupMaster')/items",{headers:headers}).map(res=>res.json());
        
        
    }

    errorHandler(error: any){
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }
}