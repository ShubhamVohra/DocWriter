"use strict";
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
Object.defineProperty(exports, "__esModule", { value: true });
/*
  This file defines a component that enables a search-and-replace functionality for
  the Word document.
*/
var core_1 = require("@angular/core");
var router_1 = require("@angular/router");
// The WordDocumentService provides methods for manipulating the document.
var word_document_service_1 = require("../services/word-document/word.document.service");
// The SettingsStorageService provides CRUD operations on application settings..
var settings_storage_service_1 = require("../services/settings-storage/settings.storage.service");
var db_conn_service_1 = require("../services/db-conn/db.conn.service");
var FindAndReplaceComponent = (function () {
    function FindAndReplaceComponent(wordDocument, settingsStorage, router, dbConn, actroute) {
        this.wordDocument = wordDocument;
        this.settingsStorage = settingsStorage;
        this.router = router;
        this.dbConn = dbConn;
        this.actroute = actroute;
        this.isCorsCompatible = function (token) {
            try {
                var xhr = new XMLHttpRequest();
                xhr.open("GET", "https://progressivedigital.sharepoint.com/_api/web/lists/GetByTitle('client')/items");
                xhr.setRequestHeader("authorization", "Bearer " + token);
                xhr.setRequestHeader("accept", "application/json");
                xhr.onload = function () {
                    // CORS is working.
                    console.log("Browser is CORS compatible.");
                };
                xhr.send();
            }
            catch (e) {
                if (e.number == -2147024891) {
                    console.log("Internet Explorer users must use Internet Explorer 11 with MS15-032: Cumulative security update for Internet Explorer (KB3038314) installed for this sample to work.");
                }
                else {
                    console.log("An unexpected error occurred. Please refresh the page.");
                }
            }
        };
        this.data = true;
        console.log("Bahar");
        this.wordDocument.getAgents().subscribe(function (data) {
            console.log("Shubham");
            console.log(data);
        });
        this.customers = [
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
    FindAndReplaceComponent.prototype.ngOnInit = function () {
        var route = window.location.href;
        console.log(route);
        this.splitQueryString(route);
    };
    // Handle the event of a user entering text in the search box.
    FindAndReplaceComponent.prototype.onSearchTextEntered = function (message) {
        this.searchString = message;
    };
    // Handle the event of a user entering text in the replace box.
    FindAndReplaceComponent.prototype.onReplaceTextEntered = function (message) {
        this.replaceString = message;
    };
    // Handle the event of a user entering a number in the box for excluded paragraphs.
    FindAndReplaceComponent.prototype.onParagraphNumeralEntered = function (message) {
        this.excludedParagraph = message;
    };
    FindAndReplaceComponent.prototype.replace = function () {
        this.wordDocument.replaceFoundStringsWithExceptions(this.searchString, this.replaceString, this.excludedParagraph);
    };
    FindAndReplaceComponent.prototype.loadData = function () {
        this.data = false;
    };
    FindAndReplaceComponent.prototype.pushData = function () {
        var agnts;
        agnts = [];
        for (var i = 0; i < this.agents.length; i++) {
            agnts[i] = this.agents[i].tName;
        }
        console.log(this.agents);
        console.log(agnts);
        this.wordDocument.replaceDocumentContent(agnts);
    };
    FindAndReplaceComponent.prototype.loadControl = function () {
        this.loadControls = true;
        this.wordDocument.loadingContentControl();
    };
    FindAndReplaceComponent.prototype.createNew = function () {
        this.wordDocument.createNewPlaceholder();
    };
    FindAndReplaceComponent.prototype.loadCustomerData = function (customer) {
        for (var i = 0; i < this.customers.length; i++) {
            if (this.customers[i].company == customer) {
                this.keys = [];
                for (var j in this.customers[i]) {
                    this.keys.push(j);
                }
                this.wordDocument.populateData(this.customers[i], this.keys);
            }
        }
    };
    FindAndReplaceComponent.prototype.loadOoxml = function () {
        this.wordDocument.loadOoxml();
    };
    FindAndReplaceComponent.prototype.splitQueryString = function (queryStringFormattedString) {
        var str = queryStringFormattedString.toString();
        var codeIndex = str.indexOf("code=");
        var sessionIndex = str.indexOf("session");
        var token = str.substring((codeIndex + 5), (sessionIndex - 1));
        console.log(token);
        this.getFilesFromO365(token);
    };
    FindAndReplaceComponent.prototype.printSomething = function () {
        console.log("Something");
    };
    FindAndReplaceComponent.prototype.getFilesFromO365 = function (token) {
        try {
            this.isCorsCompatible(token);
            var endpointUrl = "https://progressivedigital.sharepoint.com/_api/web/lists/GetByTitle('client')/items";
            var xhr = new XMLHttpRequest();
            xhr.open("GET", endpointUrl);
            // The APIs require an OAuth access token in the Authorization header, formatted like this: 'Authorization: Bearer <token>'. 
            xhr.setRequestHeader("Authorization", "Bearer " + token);
            xhr.setRequestHeader("accept", "application/json;odata=verbose");
            xhr.setRequestHeader("Access-control-allow-origin", "*");
            // Process the response from the API.  
            xhr.onload = function () {
                if (xhr.status == 200) {
                    var formattedResponse = JSON.stringify(JSON.parse(xhr.response), undefined, 2);
                    console.log(formattedResponse);
                }
                else {
                    //   document.getElementById("results").textContent = "HTTP " + xhr.status + "<br>" + xhr.response; 
                }
            };
            // Make request.
            xhr.send();
        }
        catch (err) {
            //   document.getElementById("results").textContent = "Exception: " + err.message; 
            console.log(err);
        }
    };
    FindAndReplaceComponent = __decorate([
        core_1.Component({
            templateUrl: 'app/find-and-replace/find-and-replace.component.html',
            styleUrls: ['app/find-and-replace/find-and-replace.component.css'],
        }),
        __metadata("design:paramtypes", [word_document_service_1.WordDocumentService,
            settings_storage_service_1.SettingsStorageService,
            router_1.Router, db_conn_service_1.DbConnService, router_1.ActivatedRoute])
    ], FindAndReplaceComponent);
    return FindAndReplaceComponent;
}());
exports.FindAndReplaceComponent = FindAndReplaceComponent;
//# sourceMappingURL=find-and-replace.component.js.map