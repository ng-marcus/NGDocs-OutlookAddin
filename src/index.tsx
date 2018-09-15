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
