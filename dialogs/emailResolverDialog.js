// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { InputHints, MessageFactory } = require('botbuilder');
const { WaterfallDialog, TextPrompt} = require('botbuilder-dialogs');
const { CancelAndHelpDialog } = require('./cancelAndHelpDialog');


const EMAIL_PROMPT = 'emailPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

class EmailResolverDialog extends CancelAndHelpDialog {
    constructor(id) {
        super(id || 'dateResolverDialog');
        this.addDialog(new TextPrompt(EMAIL_PROMPT, this.emailPromptValidator.bind(this)))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.initialStep.bind(this),
                this.finalStep.bind(this)
            ]));

        this.initialDialogId = WATERFALL_DIALOG;
    }

    async initialStep(stepContext) {
        const email = stepContext.options.email;

        const promptMessageText = `What is the email address for ${stepContext.options.attendee}?`;
        const promptMessage = MessageFactory.text(promptMessageText, promptMessageText, InputHints.ExpectingInput);

        const repromptMessageText = "I'm sorry, I need a valid email for ${stepContext.options.attendee}.";
        const repromptMessage = MessageFactory.text(repromptMessageText, repromptMessageText, InputHints.ExpectingInput);

        if (!email) {
            // We were not given any email at all so prompt the user.
            return await stepContext.prompt(EMAIL_PROMPT,
                {
                    prompt: promptMessage,
                    retryPrompt: repromptMessage
                });
        }

            // We are good so return.
        return await stepContext.next(timex);

    }

    async finalStep(stepContext) {
        var email = stepContext.result;
        if (email[email.length-1] === "."){ // remove trailing .
            email = email.slice(0,-1);
        }
        return await stepContext.endDialog(email);
    }

    async emailPromptValidator(promptContext) {
        const mailformat = /(?!.*\.{2})^([a-z\d!#$%&'*+\-\/=?^_`{|}~\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]+(\.[a-z\d!#$%&'*+\-\/=?^_`{|}~\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]+)*|"((([ \t]*\r\n)?[ \t]+)?([\x01-\x08\x0b\x0c\x0e-\x1f\x7f\x21\x23-\x5b\x5d-\x7e\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]|\\[\x01-\x09\x0b\x0c\x0d-\x7f\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))*(([ \t]*\r\n)?[ \t]+)?")@(([a-z\d\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]|[a-z\d\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF][a-z\d\-._~\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]*[a-z\d\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])\.)+([a-z\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]|[a-z\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF][a-z\d\-._~\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]*[a-z\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])\.?$/i;
        const valid = mailformat.test(promptContext.recognized.value);
            if (valid) {
                return (true);
            }

            return (false);
        
    }
}

module.exports.EmailResolverDialog = EmailResolverDialog;
