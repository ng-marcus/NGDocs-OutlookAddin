import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { AppContainer } from 'react-hot-loader';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';

import App from './components/App';

import './styles.less';
import 'office-ui-fabric-react/dist/css/fabric.min.css';

initializeIcons();

let isOfficeInitialized = false;

const title = 'Atvero1';

const render = (Component, Attachments) => {
  console.log(Attachments);
  ReactDOM.render(
    <AppContainer>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} />
    </AppContainer>,
    document.getElementById('container')
  );
};

const attachmentTokenCallback = (asyncResult) => {
  if (asyncResult.status === 'succeeded') {
    // Cache the result from the server.
    //serviceRequest.attachmentToken = asyncResult.value;
    // serviceRequest.state = 3;
    // testAttachments();
    console.log('token received');
    console.log(asyncResult.value);
  } else {
    // showToast("Error", "Could not get callback token: " + asyncResult.error.message);
    console.log(asyncResult.error.message);
  }
};





/* Render application after Office initializes */
Office.initialize = () => {
  isOfficeInitialized = true;
  let _mailbox = Office.context.mailbox;
  console.log('Item initialised');

  console.log(Office.context.mailbox.ewsUrl);

  Office.context.mailbox.getCallbackTokenAsync(attachmentTokenCallback);


  let _Item = _mailbox.item;
  console.log(_Item);
  // var MyEntities = _Item.getEntities();
  // console.log(MyEntities);

  let attachments = (_Item as Office.MessageRead).attachments;
  if (attachments.length > 0) {
    console.log(attachments);
    let att = attachments[0];
    console.log(att.name);
    console.log(att.id);
  } else {
    console.log('no attachments');
  }

  render(App, attachments);
};

/* Initial render showing a progress bar */
render(App, []);

if ((module as any).hot) {
  (module as any).hot.accept('./components/App', () => {
    const NextApp = require('./components/App').default;
    render(NextApp, []);
  });
}
