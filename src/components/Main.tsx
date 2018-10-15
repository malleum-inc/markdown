/**
 * @license MIT
 * @author Nadeem Douba <ndouba@redcanari.com>
 * @copyright Red Canari, Inc. 2018
 */

import * as React from 'react';

(window as any).CodeMirror = require('codemirror/lib/codemirror');
import '../lib/code-blast';
import {UnControlled as CodeMirror} from 'react-codemirror2';
import 'codemirror/mode/markdown/markdown';
import 'codemirror/mode/gfm/gfm';
import 'codemirror/keymap/sublime';
import 'codemirror/addon/edit/matchbrackets';
import 'codemirror/addon/edit/closebrackets';
import 'codemirror/addon/fold/foldgutter';
import 'codemirror/addon/comment/comment';
import 'codemirror/addon/search/match-highlighter';
import 'codemirror/addon/display/placeholder';
import 'codemirror/lib/codemirror.css';
import {EditorCommands} from "../lib/commands";

export interface MainProps {
    platform?: string;
    editorDidMount?;
    powerMode?: boolean;
    documentId: string;
}


export default class Main extends React.PureComponent<MainProps, {}> {

    static baseListRegex = '^\\s*' +
        '(' +
        '[-*>](?:\\s*\\[.])?' +
        '|' +
        '\\[\\^?[^\\]]+]:?' +
        '|' +
        '\\d+\\.' +
        ')' +
        '\\s*';

    static isListRegex = new RegExp(Main.baseListRegex);

    static isOnlyListRegex = new RegExp(`${Main.baseListRegex}$`);

    static hasEmailHeader = new RegExp('^((?:to|cc|bcc|subject)\\s*:[^\\n]+\\n+)*?---[\\n\\s]*', 'smi');

    static MAX_HEADER_LEVEL = 6;

    private commands = EditorCommands.getInstance();

    private keyMap = {
        'Cmd-B': this.commands.bold,
        'Ctrl-B': this.commands.bold,
        'Cmd-I': this.commands.italic,
        'Ctrl-I': this.commands.italic,
        'Cmd-U': this.commands.underline,
        'Ctrl-U': this.commands.underline,
        'Cmd-D': this.commands.strike,
        'Ctrl-D': this.commands.strike,
        'Cmd-L': this.commands.bulletedList,
        'Ctrl-L': this.commands.bulletedList,
        'Cmd-O': this.commands.numberedList,
        'Ctrl-O': this.commands.numberedList,
        'Cmd-T': this.commands.taskList,
        'Ctrl-T': this.commands.taskList,
        'Cmd-K': this.commands.hyperlink,
        'Ctrl-K': this.commands.hyperlink,
        'Cmd-Alt-1': this.commands.header1,
        'Ctrl-Alt-1': this.commands.header1,
        'Cmd-Alt-2': this.commands.header2,
        'Ctrl-Alt-2': this.commands.header2,
        'Cmd-Alt-3': this.commands.header3,
        'Ctrl-Alt-3': this.commands.header3,
        'Cmd-Alt-4': this.commands.header4,
        'Ctrl-Alt-4': this.commands.header4,
        'Shift-Cmd-=': this.commands.superscript,
        'Shift-Ctrl-=': this.commands.superscript,
        'Cmd-=': this.commands.subscript,
        'Ctrl-=': this.commands.subscript,
        'Shift-Cmd-;': this.commands.codeBlock,
        'Shift-Ctrl-;': this.commands.codeBlock,
        'Cmd-;': this.commands.code,
        'Ctrl-;': this.commands.code,
        'Cmd-N': this.commands.clear,
        'Ctrl-N': this.commands.clear,
        'Cmd-Enter': this.commands.insertMarkdown
    };

    private isNewLine = (change) => {
        return change.length === 2 && !change[0].length && !change[1].length;
    };

    private isList = (lastLine) => {
        lastLine = lastLine.trim();
        return !lastLine.startsWith('**') && lastLine !== '---' && Main.isListRegex.test(lastLine);
    };

    private shouldContinueList = (lastLine) => {
        return !Main.isOnlyListRegex.test(lastLine.trim());
    };

    private getNextListToken = (lastLine) => {
        const [fullMatch, token] = lastLine.match(/^\s*(\d+\.|(?:[-*>])?(?:\s*\[\^?[^\]]+]:?)?)\s*/);

        if (/^(\[\^?\d+]:?|\d+\.)$/.test(token)) {
            const number = parseInt(token.match(/\d+/)[0]);
            return fullMatch.replace(number.toString(), (number + 1).toString());
        } else {
            if (/\[.+]:?$/.test(token)) {
                const [completeMatch, content] = token.match(/\[([^\]]+)]/);
                if (content !== 'x' && content !== ' ')
                    return fullMatch.replace(completeMatch, '');
            }
            return fullMatch;
        }
    };

    private isTab = (change) => {
        return change.length === 1 && change[0] === '\t';
    };

    private isHeading = (line) => {
        return line.startsWith('#');
    };

    private headerLevel = (text) => {
        return text.match(/^#+/)[0].length;
    };

    private isFencedCodeBlock = (text) => {
        return /^[`~]{3}/.test(text);
    };

    private getLanguage = (codeFenceStart) => {
        return (codeFenceStart.match(/[^`~\s.]+/) || [undefined])[0];
    };

    onChange = (cm, change) => {
        const lastLine = cm.getLine(change.from.line);
        if (this.isFencedCodeBlock(lastLine)) {
            let language = this.getLanguage(lastLine);
            if (language) {
                language = /^([cj][+#]{0,2}|java|obj(?:ective)?-?c|scala|squirrel|ceylon)$/i.test(language) ? 'clike' : language;
                import(
                    /* webpackIgnore: true */
                    `codemirror/mode/${language}/${language}`
                ).catch(console.log);
            }
        }
        localStorage.setItem(this.props.documentId, cm.getValue());
    };

    onBeforeChange = (cm, change) => {
        const lastLine = cm.getLine(change.from.line);

        if (this.isList(lastLine)) {
            if (this.isNewLine(change.text)) {
                if (this.shouldContinueList(lastLine)) {
                    change.update(change.from, change.to, ["", this.getNextListToken(lastLine)], change.origin)
                } else {
                    change.update({line: change.from.line, ch: 0}, change.to, ["", ""], change.origin);
                }
            } else if (this.isTab(change.text)) {
                const pos = {line: change.from.line, ch: 0};
                change.update(
                    pos, pos,
                    (cm.options.indentWithTabs) ? change.text : [' '.repeat(cm.options.indentUnit)],
                    change.origin
                );
            }
        } else if (this.isTab(change.text) && this.isHeading(lastLine)) {
            let currentLevel = this.headerLevel(lastLine);
            let newLevel = Math.max((currentLevel + 1) % (Main.MAX_HEADER_LEVEL + 1), 1);

            if (newLevel > currentLevel) {
                const pos = {line: change.from.line, ch: 0};
                change.update(pos, pos, ['#'], change.origin);
            } else {
                const from = {line: change.from.line, ch: 0};
                const to = {line: change.from.line, ch: Main.MAX_HEADER_LEVEL - 1};
                change.update(from, to, [''], '+delete');
            }
        }
    };

    editorDidMount = (editor) => {
        editor.on('beforeChange', this.onBeforeChange);
        editor.on('change', this.onChange);
        editor.addKeyMap(this.keyMap);
        editor.focus();

        if (this.props.editorDidMount)
            this.props.editorDidMount(editor);
    };

    render() {
        return (
            <div className='main'>
                <CodeMirror
                    editorDidMount={this.editorDidMount}
                    options={{
                        mode: {
                            name: "gfm",
                            tokenTypeOverrides: {
                                emoji: "emoji"
                            }
                        },
                        blastCode: this.props.powerMode ? {effect: 1} : undefined,
                        style: {backgroundColor: '#000'},
                        lineWrapping: true,
                        // lineNumbers: true,
                        highlightSelectionMatches: true,
                        theme: "default",
                        matchBrackets: true,
                        autoCloseBrackets: true,
                        keyMap: 'sublime',
                        fencedCodeBlockHighlighting: true,
                        allowAtxHeaderWithoutSpace: true
                    }}
                />
            </div>
        );
    }
}
