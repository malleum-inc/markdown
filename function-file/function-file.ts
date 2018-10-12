/**
 * @license MIT
 * @author Nadeem Douba <ndouba@redcanari.com>
 * @copyright Red Canari, Inc. 2018
 */

import {EditorCommands} from '../src/lib/commands';

(() => {
    // The initialize function must be run each time a new page is loaded
    Office.initialize = () => {
        (window as any).renderMarkdown = renderMarkdown;
    };

    // Add any ui-less function here
    function renderMarkdown(event) {
        EditorCommands.getSelectedBodyAsText((content) => {
            if (content.status === Office.AsyncResultStatus.Succeeded && content.value.trim()) {
                EditorCommands.getItemType((result) => {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                        event.completed();
                    } else if (result.value === "html") {
                        EditorCommands.setSelectedBodyAsHtml(
                            EditorCommands.renderMarkdown(content.value),
                            () => event.completed()
                        );
                    }
                });
            } else {
                event.completed();
            }
        });
    }

})();
