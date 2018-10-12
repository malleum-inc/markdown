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


export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

export interface AppState {
    powerMode: boolean;
}

export default class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            powerMode: false
        };
    }

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
                <Header onPowerMode={() => this.setState({powerMode: !this.state.powerMode})}/>
                <Main
                    editorDidMount={editor => EditorCommands.setEditorInstance(editor)}
                    powerMode={this.state.powerMode}
                />
                <Footer/>
            </Fabric>
        );
    }
}
