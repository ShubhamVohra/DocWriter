"use strict";
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo.
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
  This file defines a settings view. It is based on
  the settings sample, created by the Modern Assistance Experience Developer
  Docs team. Along with other samples, it is in the Office-Add-in-UX-Design-Patterns-Code
  repo:  https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code
*/
var core_1 = require("@angular/core");
var router_1 = require("@angular/router");
var db_conn_service_1 = require("../services/db-conn/db.conn.service");
// The SettingsStorageService provides CRUD operations on application settings.
var settings_storage_service_1 = require("../services/settings-storage/settings.storage.service");
var SettingsComponent = (function () {
    function SettingsComponent(settingsStorage, activatedRoute, dbconn) {
        this.settingsStorage = settingsStorage;
        this.activatedRoute = activatedRoute;
        this.dbconn = dbconn;
        var routeFragment = activatedRoute.fragment;
        routeFragment.subscribe(function (fragment) {
            var token = fragment.match(/^(.*?)&/)[1].replace('access_token=', '');
        });
    }
    SettingsComponent.prototype.ngAfterViewInit = function () {
        var currentInstructionSetting = this.settingsStorage.fetch("StyleCheckerAddinShowInstructions");
        // Ensure that when the settings view loads, the radio button selection matches
        // the user's current setting.
        if (currentInstructionSetting === "OnlyFirstTime") {
            this.alwaysRadioButton.nativeElement.removeAttribute("checked");
            this.onlyFirstTimeRadioButton.nativeElement.setAttribute("checked", "checked");
        }
    };
    SettingsComponent.prototype.ngOnInit = function () {
        var url = window.location.href;
        console.log(url);
        var access_token = this.splitQueryString(url);
        this.getFiles(access_token);
    };
    SettingsComponent.prototype.onRadioButtonSelected = function (specificSetting, value) {
        this.settingsStorage.store(specificSetting, value);
    };
    SettingsComponent.prototype.splitQueryString = function (queryStringFormattedString) {
        var str = queryStringFormattedString.toString();
        var codeIndex = str.indexOf("access_token=");
        var sessionIndex = str.indexOf("token_type=");
        var token = str.substring((codeIndex + 13), (sessionIndex - 1));
        console.log("token is " + token);
    };
    SettingsComponent.prototype.getFiles = function (token) {
        this.dbconn.getFilesFromSharepoint(token);
    };
    __decorate([
        core_1.ViewChild('always'),
        __metadata("design:type", core_1.ElementRef)
    ], SettingsComponent.prototype, "alwaysRadioButton", void 0);
    __decorate([
        core_1.ViewChild('onlyFirstTime'),
        __metadata("design:type", core_1.ElementRef)
    ], SettingsComponent.prototype, "onlyFirstTimeRadioButton", void 0);
    SettingsComponent = __decorate([
        core_1.Component({
            templateUrl: 'app/settings/settings.component.html',
            styleUrls: ['app/settings/settings.component.css']
        }),
        __metadata("design:paramtypes", [settings_storage_service_1.SettingsStorageService, router_1.ActivatedRoute,
            db_conn_service_1.DbConnService])
    ], SettingsComponent);
    return SettingsComponent;
}());
exports.SettingsComponent = SettingsComponent;
//# sourceMappingURL=settings.component.js.map