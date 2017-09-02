"use strict";
/// 
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
var core_1 = require("@angular/core");
var http_1 = require("@angular/http");
require("rxjs/add/operator/map");
var DbConnService = (function () {
    function DbConnService(http) {
        this.http = http;
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
    }
    DbConnService.prototype.getAgents = function () {
        //return this.http.get('https://www.kansanmedtrip.com/getData.php?module=treatment').map(res=>res.json());
    };
    DbConnService.prototype.getFiles = function (token) {
        var headers = new http_1.Headers({
            "Authorization": "Bearer " + token,
            "Accept": "application/json;odata=verbose",
            "Access-Control-Allow-Origin": "https://localhost:3000/settings/"
        });
        var options = new http_1.RequestOptions({
            headers: headers,
            withCredentials: true
        });
        console.log("get files called");
        return this.http.get("https://progressivedigital.sharepoint.com/_api/web/lists/GetByTitle('client')/items", options).map(function (res) { console.log("Response  " + res.json()); }).subscribe(function (res) { console.log(res); });
    };
    DbConnService.prototype.getFilesFromSharepoint = function (token) {
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
    DbConnService = __decorate([
        core_1.Injectable(),
        __metadata("design:paramtypes", [http_1.Http])
    ], DbConnService);
    return DbConnService;
}());
exports.DbConnService = DbConnService;
//# sourceMappingURL=db.conn.service.js.map