/**
 * @license MIT
 * @author Nadeem Douba <ndouba@redcanari.com>
 * @copyright Red Canari, Inc. 2018
 */

import * as React from 'react';
import Progress from './Progress';
import Header from "./Header";
import Main from "./Main";
import Footer from "./Footer";
import {EditorCommands} from "../lib/commands";
import {Fabric} from "office-ui-fabric-react";
import {ProgressIndicator} from "office-ui-fabric-react/lib";

export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

export interface AppState {
    powerMode: boolean;
    dialog: any;
    documentId: string;
}

export default class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);

        this.state = {
            powerMode: false,
            dialog: undefined,
            documentId: EditorCommands.getDocumentId()
        };
    }

    listenForUpdates = (event) => {
        console.log(event);
        if (event.key === this.state.documentId && event.newValue !== event.oldValue)
            EditorCommands.loadDocument(this.state.documentId);
    };

    componentDidMount() {
        (window as any).addEventListener("storage", this.listenForUpdates);
    }

    closeDialog = () => {
        this.state.dialog.close();
        this.setState({dialog: undefined});
    };

    handleMessage = (message) => {
        console.log(message);
        switch(message.message) {
            case 'insertMarkdown':
                EditorCommands.getInstance().insertMarkdown();
                this.closeDialog();
                break;
            case 'close':
                this.closeDialog();
                break;
        }
    };

    render() {
        const {
            title,
            isOfficeInitialized,
        } = this.props;

        if (!isOfficeInitialized) {
            return (
                <Fabric id={'container'} className='ms-bgColor-black'>
                    <Progress
                        title={title}
                        logo='assets/watermark-white.png'
                        message='Supercharging Outlook!'
                    />
                </Fabric>
            );
        }

        return (
            <Fabric id={'container'}>
                <Header
                    onPowerMode={() => this.setState({powerMode: !this.state.powerMode})}
                    onOpenWindow={(dialog) => {this.setState({dialog})}}
                    onWindowEvent={(event) => {console.log(event)}}
                    onWindowMessage={this.handleMessage}
                    documentId={this.state.documentId}
                />
                {this.state.powerMode && <ProgressIndicator description="Power-mode Enabled!" />}
                <Main
                    editorDidMount={editor => EditorCommands.setEditorInstance(editor)}
                    powerMode={this.state.powerMode}
                    documentId={this.state.documentId}
                />
                <Footer/>
            </Fabric>
        );
    }
}
