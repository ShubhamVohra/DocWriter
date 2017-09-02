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
  This file defines an instructions component for a task pane page. It is based on
  the instruction-step sample, created by the Modern Assistance Experience Developer
  Docs team. Along with other samples, it is in the Office-Add-in-UX-Design-Patterns-Code
  repo:  https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code
*/
var core_1 = require("@angular/core");
var router_1 = require("@angular/router");
var queryStringParameters;
var InstructionStepsComponent = (function () {
    function InstructionStepsComponent(router, route) {
        this.router = router;
        this.route = route;
        this.title = "EY TEMPLATE DESIGNER";
        this.addin_description = "Template Designer enables you to enforce style rules while exempting paragraphs that you specify from the rules.";
        this.steps_intro = "Just take these steps:";
        this.steps = [{ step_number: 1, content: "Enter a string in the Find box." },
            { step_number: 2, content: "Enter a replacement string in the Replace With box." },
            { step_number: 3, content: "Enter the zero-based numbers of the parapgraphs that should be exempt in the Skip Paragraphs box." },
            { step_number: 4, content: "Press Replace." }];
        this.parseQueryString = function (url) {
            var params = {}, queryString = url.substring(1), regex = /([^&=]+)=([^&]*)/g, m;
            while (m = regex.exec(queryString)) {
                params[decodeURIComponent(m[1])] = decodeURIComponent(m[2]);
            }
            return params;
        };
        this.params = this.parseQueryString(location.hash);
        this.access_token = null;
    }
    InstructionStepsComponent.prototype.ngOnInit = function () {
        var urout = this.router.url;
        console.log(urout);
        console.log(window.location.pathname);
        var clientId = '44d3241c-4d7a-439b-a37c-df8f6e8d689d';
        var replyUrl = 'https://10.80.0.236:3000/settings/';
        var endpointUrl = 'https://progressivedigital.sharepoint.com/_api/v1.0/me/files';
        var resource = "https://progressivedigital.sharepoint.com";
        var authServer = 'https://login.windows.net/common/oauth2/authorize?';
        var responseType = 'token';
        var url = authServer +
            "response_type=" + encodeURI(responseType) + "&" +
            "client_id=" + encodeURI(clientId) + "&" +
            "resource=" + encodeURI(resource) + "&" +
            "redirect_uri=" + encodeURI(replyUrl) + "&" + "state=" + 12345;
        console.log(url);
        // alert(url);
        window.open(url, "_self");
        var token = this.params['access_token'];
        var a = this.splitQueryString(url);
        console.log(token);
    };
    InstructionStepsComponent.prototype.splitQueryString = function (queryStringFormattedString) {
        var split = queryStringFormattedString.split('&');
        if (split == "") {
            return {};
        }
        var results = {};
        for (var i = 0; i < split.length; i++) {
            var p = split[i].split("=", 2);
            if (p.length == 1) {
                results[p[0]] = "";
            }
            else {
                results[p[0]] = decodeURIComponent(p[1].replace(/\+/g, " "));
            }
        }
        return results;
    };
    InstructionStepsComponent.prototype.printSomething = function () {
        console.log("Something");
    };
    InstructionStepsComponent = __decorate([
        core_1.Component({
            templateUrl: 'app/instructions/instruction-steps.component.html',
            styleUrls: ['app/instructions/instruction-steps.component.css']
        }),
        __metadata("design:paramtypes", [router_1.Router, router_1.ActivatedRoute])
    ], InstructionStepsComponent);
    return InstructionStepsComponent;
}());
exports.InstructionStepsComponent = InstructionStepsComponent;
//# sourceMappingURL=instruction-steps.component.js.map