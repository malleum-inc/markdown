import * as React from 'react';
import Progress from './Progress';
import Header from "./Header";
import Main from "./Main";
import Footer from "./Footer";
import {EditorCommands} from "../lib/commands";


export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

export interface AppState {
}

export default class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
        };
    }

    render() {
        const {
            title,
            isOfficeInitialized,
        } = this.props;

        if (!isOfficeInitialized) {
            return (
                <Progress
                    title={title}
                    logo='assets/icon0.png'
                    message='Loading...'
                />
            );
        }

        return (
            <>
                <Header/>
                <Main editorDidMount={editor => EditorCommands.setEditorInstance(editor)}/>
                <Footer/>
            </>
        );
    }
}
