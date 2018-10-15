/**
 * @license MIT
 * @author Nadeem Douba <ndouba@redcanari.com>
 * @copyright Red Canari, Inc. 2018
 */

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {AppContainer} from 'react-hot-loader';
import {initializeIcons} from 'office-ui-fabric-react/lib/Icons';

import App from './components/App';

import './styles.less';

initializeIcons();

let isOfficeInitialized = false;

const title = 'Markdown';

const render = (Component) => {
    ReactDOM.render(
        <AppContainer>
            <Component title={title} isOfficeInitialized={isOfficeInitialized}/>
        </AppContainer>,
        document.body.firstElementChild
    );
};

/* Render application after Office initializes */
Office.initialize = () => {
    isOfficeInitialized = true;

    if (!Office.context.mailbox)
        setInterval(() => !window.opener && window.close(), 100);

    render(App);
};

/* Initial render showing a progress bar */
render(App);

if ((module as any).hot) {
    (module as any).hot.accept('./components/App', () => {
        const NextApp = require('./components/App').default;
        render(NextApp);
    });
}
