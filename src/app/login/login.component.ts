/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import { Component } from '@angular/core';

import { AuthService } from '../auth/auth.service';

@Component({
  selector: 'app-login',
  template: `
  <div class="ms-Grid-row">
  <!-- App navigation bar markup. -->
  <div class="ms-NavBar">
    <ul class="ms-NavBar-items">
      <li class="ms-NavBar-item">Microsoft Graph Connect sample</li>
    </ul>
  </div>

  <!-- App main content markup. -->
    <div class="ms-Grid-col ms-u-mdPush1 ms-u-md9 ms-u-lgPush1 ms-u-lg6">
    <div>
    <p class="ms-font-xl">Use the button below to connect to Microsoft Graph.</p>
    <button class="ms-Button" (click)="onLogin()">
      <span class="ms-Button-label">Connect</span>
    </button>
    </div>
  </div>
</div>
  `
})
export class LoginComponent {
  constructor(private authService: AuthService) {}

  onLogin() {
    this.authService.login();
  }
}
