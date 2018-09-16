import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { AppContainer } from 'react-hot-loader';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
// import { setup as pnpSetup } from '@pnp/common';
import {
  getRandomString,
} from '@pnp/common';

//import { sp } from '@pnp/sp';
// import { SPFetchClient } from '@pnp/nodejs';
//import * as MicrosoftGraph from '@microsoft/microsoft-graph-client';


import App from './components/App';

import './styles.less';
import 'office-ui-fabric-react/dist/css/fabric.min.css';

initializeIcons();

console.log(getRandomString(20));


let isOfficeInitialized = false;

const title = 'Atvero1';

const render = (Component, Attachments) => {
  console.log(Attachments);
  ReactDOM.render(
    <AppContainer>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} attachments={Attachments} />
    </AppContainer>,
    document.getElementById('container')
  );
};






/* Render application after Office initializes */
Office.initialize = () => {
  isOfficeInitialized = true;
  let _mailbox = Office.context.mailbox;
  console.log('Item initialised');

  console.log(Office.context.mailbox.ewsUrl);


  let _Item = _mailbox.item;
  console.log(_Item);


  let attachments = (_Item as Office.MessageRead).attachments;
  let realAttachments = attachments.filter(a => !a.isInline);
  // if (attachments.length > 0) {
  //   console.log(attachments);
  //   let att = attachments[0];
  //   console.log(att.name);
  //   console.log(att.id);
  // } else {
  //   console.log('no attachments');
  // }

  console.log('Entities:');
  let entities = (_Item as Office.MessageRead).getEntities();
  console.log(entities);
  console.log(entities.contacts);
  // Check to make sure that address entities are present.
  if (null != entities && null != entities.addresses && undefined !== entities.addresses) {
    //Addresses are present, so use them here.
    console.log('addresses');
    console.log(entities.addresses);
  }



  // Office.context.mailbox.getCallbackTokenAsync(result => {
  //   if (result.status === Office.AsyncResultStatus.Succeeded) {
  //     // Use this token to call Web API
  //     console.log('SSO succeeded');
  //     let ssoToken = result.value;
  //     console.log(ssoToken);
  //     fetch('https://graph.microsoft.com/beta/sites/atverodev/sites',
  //       {
  //         method: 'GET',
  //         headers: {
  //           'Authorization': 'Bearer ' + ssoToken,
  //           'Accept': 'application/json',
  //           'Content-Type': 'application/json'
  //         }
  //       })
  //       .then(response => console.log(response));
  //     // let client = MicrosoftGraph.Client.init({
  //     //   authProvider: (done) => {
  //     //     done(null, ssoToken); //first parameter takes an error if you can't get an access token
  //     //   }
  //     // });
  //     // client
  //     //   .api('/me')
  //     //   .get((err, res) => {
  //     //     console.log(res); // prints info about authenticated user
  //     //     console.log(err);
  //     //   });
  //   } else {
  //     if (result.error.code === 13003) {
  //       console.log('SSO not supported');
  //       // SSO is not supported for domain user accounts, only
  //       // work or school (Office 365) or Microsoft Account IDs.
  //     } else {
  //       // Handle error
  //       console.log('generic error');
  //     }
  //   }
  // });

  // const mySP = sp.configure({
  //   headers: {
  //     'X-Header': 'My header'
  //   }
  // }, 'https://atverodev.sharepoint.com');

  // mySP.web.lists.get().then(l => console.log(l))
  //   .catch(e => console.log(e));





  render(App, realAttachments);
};

/* Initial render showing a progress bar */
render(App, []);

if ((module as any).hot) {
  (module as any).hot.accept('./components/App', () => {
    const NextApp = require('./components/App').default;
    render(NextApp, []);
  });
}
