// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo.

/*
  This file defines a settings view. It is based on
  the settings sample, created by the Modern Assistance Experience Developer 
  Docs team. Along with other samples, it is in the Office-Add-in-UX-Design-Patterns-Code 
  repo:  https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code
*/

import { Component, AfterViewInit, ElementRef, ViewChild } from '@angular/core';
import { Router,ActivatedRoute } from '@angular/router';

import { NavigationHeaderComponent} from '../shared/navigation-header/navigation.header.component';
import { ButtonComponent } from '../shared/button/button.component';
import { BrandFooterComponent} from '../shared/brand-footer/brand.footer.component';
import { DbConnService } from '../services/db-conn/db.conn.service';

import {Observable} from 'rxjs/Observable';

// The SettingsStorageService provides CRUD operations on application settings.
import { SettingsStorageService } from '../services/settings-storage/settings.storage.service';

@Component({
    templateUrl: 'app/settings/settings.component.html',
    styleUrls: ['app/settings/settings.component.css']
})
export class SettingsComponent {
   
   // Get references to the radio buttons so we can toggle which is selected.
   @ViewChild('always') alwaysRadioButton: ElementRef;
   @ViewChild('onlyFirstTime') onlyFirstTimeRadioButton: ElementRef;

  constructor(private settingsStorage: SettingsStorageService,private activatedRoute: ActivatedRoute,
              private dbconn:DbConnService) {
    const routeFragment: Observable<string> = activatedRoute.fragment;
    routeFragment.subscribe(fragment => {
      let token: string = fragment.match(/^(.*?)&/)[1].replace('access_token=', '');
      
      
    });
  }

  ngAfterViewInit() {
    let currentInstructionSetting: string = this.settingsStorage.fetch("StyleCheckerAddinShowInstructions");
    
    // Ensure that when the settings view loads, the radio button selection matches
    // the user's current setting.


    if (currentInstructionSetting === "OnlyFirstTime") { 
      this.alwaysRadioButton.nativeElement.removeAttribute("checked");
      this.onlyFirstTimeRadioButton.nativeElement.setAttribute("checked", "checked");
    }
  }

  ngOnInit(){
    var url = window.location.href;
    console.log(url);
    var access_token = this.splitQueryString(url);
    this.getFiles(access_token);
  }

  onRadioButtonSelected(specificSetting: string, value: string){
    this.settingsStorage.store(specificSetting, value);
  }

  splitQueryString(queryStringFormattedString:string){
    var str = queryStringFormattedString.toString();
    var codeIndex = str.indexOf("access_token=");
    var sessionIndex = str.indexOf("token_type=");
    var token = str.substring((codeIndex+13),(sessionIndex-1));
    console.log("token is "+token);
  }

  getFiles(token:any){
    this.dbconn.getFilesFromSharepoint(token);
  }

  
}

