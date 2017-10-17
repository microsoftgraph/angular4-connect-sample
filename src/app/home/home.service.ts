/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import 'rxjs/add/operator/catch';
import 'rxjs/add/operator/map';
import 'rxjs/add/observable/fromPromise';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"
import * as MicrosoftGraphClient from "@microsoft/microsoft-graph-client"
import { Injectable } from '@angular/core';
import { Http } from '@angular/http';
import { Observable } from 'rxjs/Observable';

import { HttpService } from '../shared/http.service';

@Injectable()
export class HomeService {
  url = 'https://graph.microsoft.com/v1.0';
  file = 'demo.xlsx';
  table = 'Table1';

  constructor(
    private http: Http,
    private httpService: HttpService) {
  }

  getClient(): MicrosoftGraphClient.Client
  {
    var client = MicrosoftGraphClient.Client.init({
      authProvider: (done) => {
          done(null, this.httpService.getAccessToken()); //first parameter takes an error if you can't get an access token
      }
    });
    return client;
  }

  getMe(): Observable<MicrosoftGraph.User>
  {
    var client = this.getClient();
    return Observable.fromPromise(client
    .api('me')
    .select("displayName, mail, userPrincipalName")
    .get()
    .then ((res => {
      return res;
    } ) )
    );
  }

  sendMail(mail: MicrosoftGraph.Message) {
    var client = this.getClient();
    return Observable.fromPromise(client
    .api('me/sendmail')
    .post({message: mail})
  );
  } 


}
