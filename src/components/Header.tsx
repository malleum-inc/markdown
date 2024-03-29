/**
 * @license MIT
 * @author Nadeem Douba <ndouba@redcanari.com>
 * @copyright Red Canari, Inc. 2018
 */

import * as React from 'react';

import {CommandBar, ICommandBarItemProps} from 'office-ui-fabric-react/lib/CommandBar';
import {EditorCommands} from "../lib/commands";

export interface HeaderProps {
    onPowerMode: any;
    onOpenWindow: any;
    onWindowEvent: any;
    onWindowMessage: any;
    documentId: string;
}

export interface HeaderState {
    powerMode: boolean;
}

export default class Header extends React.Component<HeaderProps, HeaderState> {

    state = {
        powerMode: false
    };

    editorCommands = EditorCommands.getInstance();

    public render(): JSX.Element {
        return (
            <div className='header'>
                <CommandBar
                    items={this.getItems()}
                    farItems={this.getFarItems()}
                    overflowItems={this.getOverflowItems()}
                    // isSearchBoxVisible
                    // searchPlaceholderText='Find Text...'
                    // onInput={console.log}
                />
            </div>
        );
    }

    private getOverflowItems = (): ICommandBarItemProps[] => {
        return [];
    };

    private processWindowMessage = (message) => {
        console.log(JSON.stringify(message));
        this.props.onWindowMessage(message);
    };

    private processWindowEvent = (event) => {
        console.log(JSON.stringify(event));
        this.props.onWindowEvent(event)
    };

    private getFarItems = (): ICommandBarItemProps[] => {
        return [
            {
                key: 'powerMode',
                name: 'Power Mode',
                cacheKey: `powerMode-${this.state.powerMode}`,
                iconProps: {
                    iconName: 'PowerButton'
                },
                canCheck: true,
                iconOnly: true,
                checked: this.state.powerMode,
                onClick: () => {
                    this.setState({powerMode: !this.state.powerMode});
                    this.props.onPowerMode();
                    this.editorCommands.focus();
                }
            },
            {
                key: 'openCloseEditorWindow',
                name: (Office.context.mailbox)?'Open in Separate Window':'Close Window',
                iconProps: {
                    iconName: (Office.context.mailbox)?'OpenInNewWindow':'ChromeClose'
                },
                iconOnly: true,
                onClick: () => {
                    if (Office.context.mailbox) {
                        Office.context.ui.displayDialogAsync(
                            `/?documentId=${this.props.documentId}`,
                            (result) => {
                                if (result.status == Office.AsyncResultStatus.Succeeded) {
                                    let dialog = result.value;
                                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, this.processWindowMessage);
                                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, this.processWindowEvent);
                                    this.props.onOpenWindow(dialog);
                                }
                            }
                        );
                    } else {
                        Office.context.ui.messageParent('close');
                    }
                }
            },
            {
                key: 'help',
                name: 'Help',
                iconProps: {
                    iconName: 'Info'
                },
                iconOnly: true,
                href: 'https://help.github.com/articles/basic-writing-and-formatting-syntax/',
                ['data-automation-id']: 'helpButton'
            }
        ];
    };

    // Data for CommandBar
    private getItems = (): ICommandBarItemProps[] => {
        return [
            {
                key: 'renderMarkdown',
                name: 'Insert Markdown',
                onClick: () => this.editorCommands.insertMarkdown(),
                iconProps: {
                    iconName: 'MarkDownLanguage'
                }
            },
            {
                key: 'formatMenu',
                name: 'Format',
                cacheKey: 'formatMenu',
                subMenuProps: {
                    items: [
                        {
                            key: 'bold',
                            name: 'Bold',
                            shortCut: 'Cmd-B',
                            iconProps: {
                                iconName: 'Bold'
                            },
                            onClick: () => this.editorCommands.bold()
                        },
                        {
                            key: 'italic',
                            name: 'Italic',
                            iconProps: {
                                iconName: 'Italic'
                            },
                            onClick: () => this.editorCommands.italic()
                        },
                        {
                            key: 'underline',
                            name: 'Underline',
                            iconProps: {
                                iconName: 'Underline'
                            },
                            onClick: () => this.editorCommands.underline()
                        },
                        {
                            key: 'strikethrough',
                            name: 'Strikethrough',
                            iconProps: {
                                iconName: 'Strikethrough'
                            },
                            onClick: () => this.editorCommands.strike()
                        },
                        {
                            name: 'Superscript',
                            key: 'superscript',
                            iconProps: {
                                iconName: 'Superscript'
                            },
                            onClick: () => this.editorCommands.superscript()
                        },
                        {
                            key: 'subscript',
                            name: 'Subscript',
                            iconProps: {
                                iconName: 'Subscript'
                            },
                            onClick: () => this.editorCommands.subscript()
                        },
                        {
                            key: 'separator-0',
                            name: '-'
                        },
                        {
                            key: 'heading1',
                            name: 'Header 1',
                            iconProps: {
                                iconName: 'Header1'
                            },
                            onClick: () => this.editorCommands.header1()
                        },
                        {
                            key: 'heading2',
                            name: 'Header 2',
                            iconProps: {
                                iconName: 'Header2'
                            },
                            onClick: () => this.editorCommands.header2()
                        },
                        {
                            key: 'heading3',
                            name: 'Header 3',
                            iconProps: {
                                iconName: 'Header3'
                            },
                            onClick: () => this.editorCommands.header3()
                        },
                        {
                            key: 'heading4',
                            name: 'Header 4',
                            iconProps: {
                                iconName: 'Header4'
                            },
                            onClick: () => this.editorCommands.header4()
                        },
                        {
                            key: 'separator-1',
                            name: '-'
                        },
                        {
                            key: 'quote',
                            name: 'Quote',
                            iconProps: {
                                iconName: 'RightDoubleQuote'
                            },
                            onClick: () => this.editorCommands.quote()
                        },
                        {
                            key: 'code',
                            name: 'Inline Code',
                            iconProps: {
                                iconName: 'CodeEdit'
                            },
                            onClick: () => this.editorCommands.code()
                        }
                    ]
                }
            },
            {
                key: 'insertItem',
                name: 'Insert',
                cacheKey: 'insertMenu',
                ariaLabel: 'Format selected text',
                subMenuProps: {
                    items: [
                        {
                            key: 'bulletedList',
                            name: 'Bulleted List',
                            iconProps: {
                                iconName: 'BulletedList'
                            },
                            onClick: () => this.editorCommands.bulletedList()
                        },
                        {
                            key: 'numberedList',
                            name: 'Numbered List',
                            iconProps: {
                                iconName: 'NumberedList'
                            },
                            onClick: () => this.editorCommands.numberedList()
                        },
                        {
                            key: 'taskList',
                            name: 'Task List',
                            iconProps: {
                                iconName: 'TaskLogo'
                            },
                            onClick: () => this.editorCommands.taskList()
                        },
                        {
                            key: 'separator-3',
                            name: '-',
                        },
                        // {
                        //     key: 'insertTable',
                        //     name: 'Insert Table',
                        //     iconProps: {
                        //         iconName: 'Table'
                        //     }
                        // },
                        // {
                        //     key: 'separator-4',
                        //     name: '-',
                        // },
                        {
                            key: 'insertHyperlink',
                            name: 'Hyperlink',
                            iconProps: {
                                iconName: 'Link'
                            },
                            onClick: () => this.editorCommands.hyperlink()
                        },
                        {
                            key: 'separator-5',
                            name: '-',
                        },
                        {
                            key: 'insertPicture',
                            name: 'Insert Picture',
                            iconProps: {
                                iconName: 'Picture'
                            },
                            onClick: () => this.editorCommands.picture()
                        },
                        {
                            key: 'separator-6',
                            name: '-',
                        },
                        {
                            key: 'horizontalLine',
                            name: 'Horizontal Line',
                            iconProps: {
                                iconName: 'CalculatorSubtract'
                            },
                            onClick: () => this.editorCommands.horizontalLine()
                        },
                        {
                            key: 'separator-7',
                            name: '-',
                        },
                        {
                            key: 'codeBlock',
                            name: 'Code Block',
                            iconProps: {
                                iconName: 'Code'
                            },
                            onClick: () => this.editorCommands.codeBlock()
                        },
                    ]
                }
            }
        ];
    };

}
