import {IInstance} from "react-codemirror2";
import Main from "../components/Main";

const juice = require('juice/client');

const extraStyles = [
    require('highlightjs/styles/github.css?inline'),
    require('font-awesome/css/font-awesome.css?inline').replace(/url\(\/assets/mg, `url(${document.location.origin}/assets`),
    require('../style/extra.css?inline')
].join('');

const md = require('markdown-it')({
    html: true,
    linkify: true,
    typographer: true
})
    .use(require('markdown-it-footnote'))
    .use(require('markdown-it-emoji'))
    .use(require('markdown-it-deflist'))
    .use(require('markdown-it-sub'))
    .use(require('markdown-it-sup'))
    .use(require('markdown-it-abbr'))
    .use(require('markdown-it-highlightjs'))
    .use(require('markdown-it-task-lists'))
    .use(require('markdown-it-imsize'))
    .use(require('../lib/markdown-it-fontawesome'));


interface ICursor {
    line: number
    ch: number
}

const CHECKED = '&#9989;';
const UNCHECKED = '&#10060;';

interface ICodeMirrorEditor extends IInstance {
    getSelection(): string;

    replaceSelection(value: string, where?: string);

    getCursor(): ICursor;

    setSelection(anchor: ICursor, head: ICursor);

    getLine(lineNumber: number): string;

    focus();

    getValue();

    setValue(value: string): void;
}

export interface IEditorCommands {
    bold(cm?: ICodeMirrorEditor);

    italic(cm?: ICodeMirrorEditor);

    code(cm?: ICodeMirrorEditor);

    strike(cm?: ICodeMirrorEditor);

    underline(cm?: ICodeMirrorEditor);

    horizontalLine(cm?: ICodeMirrorEditor);

    header1(cm?: ICodeMirrorEditor);

    header2(cm?: ICodeMirrorEditor);

    header3(cm?: ICodeMirrorEditor);

    header4(cm?: ICodeMirrorEditor);

    codeBlock(cm?: ICodeMirrorEditor);

    numberedList(cm?: ICodeMirrorEditor);

    bulletedList(cm?: ICodeMirrorEditor);

    taskList(cm?: ICodeMirrorEditor);

    picture(cm?: ICodeMirrorEditor);

    superscript(cm?: ICodeMirrorEditor);

    subscript(cm?: ICodeMirrorEditor);

    underline(cm?: ICodeMirrorEditor);

    quote(cm?: ICodeMirrorEditor);

    horizontalLine(cm?: ICodeMirrorEditor);

    clear(cm?: ICodeMirrorEditor);

    renderMarkdown(cm?: ICodeMirrorEditor, text?: string)

    insertMarkdown(cm?: ICodeMirrorEditor)
}

export class EditorCommands implements IEditorCommands {

    private static instance = new EditorCommands();
    private static editor: ICodeMirrorEditor = null;

    private constructor() {
    }

    static setEditorInstance(editor) {
        this.editor = editor;
    }

    static getInstance() {
        return this.instance;
    }

    private isSurrounded = (text: string, token: string, endToken?: string): boolean => {
        return text.startsWith(token) && text.endsWith(endToken);
    };

    private wrapText = (text: string, startToken: string, endToken?: string): string => {
        return `${startToken}${text}${endToken}`
    };

    private unwrapText = (text: string, startToken: string, endToken?: string): string => {
        return text.slice(startToken.length, -endToken.length);
    };

    private toggleWrap = (cm: ICodeMirrorEditor, token: string, endToken?: string): void => {
        const text = cm.getSelection();

        if (!endToken)
            endToken = token;

        if (this.isSurrounded(text, token, endToken)) {
            cm.replaceSelection(this.unwrapText(text, token, endToken), 'around');
        } else {
            cm.replaceSelection(this.wrapText(text, token, endToken), 'around');

            if (!text) {
                let pos = cm.getCursor();
                let numLines = (endToken.match(/[\r\n]+/) || []).length;
                let newPos = {line: pos.line - numLines, ch: pos.ch + numLines - endToken.length};
                cm.setSelection(newPos, newPos);
            }
        }

        cm.focus();
    };

    private toggleWrapHtml = (cm: ICodeMirrorEditor, tag: string) => this.toggleWrap(cm, `<${tag}>`, `</${tag}>`);

    private lineStartsWith = (cm: ICodeMirrorEditor, token: string) => {
        return cm.getLine(cm.getCursor().line).startsWith(token);
    };

    private prependLine = (cm: ICodeMirrorEditor, token: string) => {
        let oldPos = cm.getCursor();
        cm.setCursor(cm.getCursor().line, 0);
        cm.replaceSelection(token);
        cm.setCursor({line: oldPos.line, ch: oldPos.ch + token.length});
    };

    private unprependLine = (cm: ICodeMirrorEditor, token: string) => {
        let oldPos = cm.getCursor();
        let line: number = cm.getCursor().line;
        cm.setSelection({line, ch: 0}, {line, ch: token.length});
        cm.replaceSelection('');
        cm.setCursor({line: oldPos.line, ch: oldPos.ch - token.length});
    };

    private toggleInsert = (cm: ICodeMirrorEditor, token: string, forcePrepend: boolean = false) => {
        if (this.lineStartsWith(cm, token) && !forcePrepend) {
            this.unprependLine(cm, token);
        } else {
            this.prependLine(cm, token);
        }
        cm.focus();
    };

    private toggleWrapLink(cm, isImage: boolean = false) {
        const text = cm.getSelection();
        let link: string = '';
        let caption: string[] = [];

        // If it's an hyperlink and we ask to insert an image
        if (!isImage && /^!\[[^\]]*]\([^)]*\)/.test(text)) {
            cm.replaceSelection(text.slice(1))
            // If it's an image and we ask to insert a hyperlink
        } else if (isImage && /^\[[^\]]*]\([^)]*\)/.test(text)) {
            cm.replaceSelection(`!${text}`)
        } else {

            let isWrapped = text.match(/^!?\[([^\]]*)]\(([^)]*)\)/);

            if (isWrapped) {
                cm.replaceSelection(isWrapped.slice(1, 3).join(' ').trimRight());
            } else {
                if (text) {
                    text.split(/\s+/).forEach((t) => {
                        if (/^https?:\/\//.test(t))
                            link = t;
                        else
                            caption.push(t);
                    });
                }

                link = `[${caption.join(' ')}](${link})`;

                if (isImage)
                    link = `!${link}`;

                cm.replaceSelection(link);
            }
        }
    }

    private headerLevel = (text) => {
        return text.match(/^#+/)[0].length;
    };

    private toggleHeader = (cm, level) => {
        const line = cm.getLine(cm.getCursor().line);

        if (line.startsWith('#')) {
            const currentLevel = this.headerLevel(line);
            if (currentLevel == level)
                this.toggleInsert(cm, `${'#'.repeat(level)} `);
            else
                this.toggleInsert(cm, `${'#'.repeat(Math.abs(level - currentLevel))}`, level > currentLevel)
        } else {
            this.toggleInsert(cm, `${'#'.repeat(level)} `);
        }
    };

    public bold = (cm = EditorCommands.editor) => this.toggleWrap(cm, '**');

    public italic = (cm = EditorCommands.editor) => this.toggleWrap(cm, '_');

    public code = (cm = EditorCommands.editor) => this.toggleWrap(cm, '`');

    public strike = (cm = EditorCommands.editor) => this.toggleWrap(cm, '~~');

    public header1 = (cm = EditorCommands.editor) => this.toggleHeader(cm, 1);

    public header2 = (cm = EditorCommands.editor) => this.toggleHeader(cm, 2);

    public header3 = (cm = EditorCommands.editor) => this.toggleHeader(cm, 3);

    public header4 = (cm = EditorCommands.editor) => this.toggleHeader(cm, 4);

    public quote = (cm = EditorCommands.editor) => this.toggleInsert(cm, '> ');

    public bulletedList = (cm = EditorCommands.editor) => this.toggleInsert(cm, '* ');

    public numberedList = (cm = EditorCommands.editor) => this.toggleInsert(cm, '1. ');

    public taskList = (cm = EditorCommands.editor) => this.toggleInsert(cm, '* [ ] ');

    public hyperlink = (cm = EditorCommands.editor) => this.toggleWrapLink(cm);

    public codeBlock = (cm = EditorCommands.editor) => this.toggleWrap(cm, '```\n', '\n```');

    public picture = (cm = EditorCommands.editor) => this.toggleWrapLink(cm, true);

    public underline = (cm = EditorCommands.editor) => this.toggleWrapHtml(cm, 'u');

    public superscript = (cm = EditorCommands.editor) => this.toggleWrap(cm, '~');

    public subscript = (cm = EditorCommands.editor) => this.toggleWrap(cm, '^');

    public horizontalLine = (cm = EditorCommands.editor) => cm.replaceSelection('\n\n---\n\n');

    public renderMarkdown = (_cm = EditorCommands.editor, text?: string) => {
        let content = juice(md.render(text), {extraCss: extraStyles}) + '<br/>';

        (content.match(/<input[^>]+>/gm) || []).forEach((m) => {
            content = content.replace(m, `${m.indexOf('checked') !== -1 ? CHECKED : UNCHECKED} &nbsp; `)
        });

        console.log(content);
        return content;
    };

    private setSelectedBody = (text) => {
        Office.context.mailbox.item.body.setSelectedDataAsync(
            text,
            {
                coercionType: Office.CoercionType.Html,
                asyncContext: {}
            }
        );
    };

    private setField = (field, value) => {
        console.log(field);
        Office.context.mailbox.item[field].setAsync(value);
    };

    public insertMarkdown = (cm = EditorCommands.editor) => {
        let text = cm.getValue();
        const [emailHeader] = text.match(Main.hasEmailHeader) || [undefined];

        if (emailHeader) {
            const mailProps = {
                to: [],
                cc: [],
                bcc: [],
                subject: null
            };

            emailHeader.split(/\n+/).forEach((line) => {
                let [field, value] = line.trim().split(/\s*:\s*/);

                if (!value)
                    return;

                field = field.toLowerCase();

                if (field === 'subject') {
                    mailProps[field] = value.trim();
                } else {
                    value.split(/\s*,+\s*/).forEach((v) => {
                        if (v) mailProps[field].push(v.trim());
                    });
                }

                if (mailProps[field] !== null && mailProps[field].length)
                    this.setField(field, mailProps[field]);
            });

            text = text.replace(emailHeader, '').trim();
        }

        if (text)
            this.setSelectedBody(this.renderMarkdown(cm, text));
    };

    public clear = (cm = EditorCommands.editor) => cm.setValue('');

}
