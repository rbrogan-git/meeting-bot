// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TimexProperty } = require('@microsoft/recognizers-text-data-types-timex-expression');
const { InputHints, MessageFactory } = require('botbuilder');
const { ConfirmPrompt, TextPrompt, WaterfallDialog } = require('botbuilder-dialogs');
const { CancelAndHelpDialog } = require('./cancelAndHelpDialog');
const { DateResolverDialog } = require('./dateResolverDialog');
const { EmailResolverDialog } = require('./emailResolverDialog');

const CONFIRM_PROMPT = 'confirmPrompt';
const DATE_RESOLVER_DIALOG = 'dateResolverDialog';
const EMAIL_RESOLVER_DIALOG = 'emailResolverDialog';
const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

class MeetingDialog extends CancelAndHelpDialog {
    constructor(id) {
        super(id || 'meetingDialog');

        this.addDialog(new TextPrompt(TEXT_PROMPT))
            .addDialog(new ConfirmPrompt(CONFIRM_PROMPT))
            .addDialog(new DateResolverDialog(DATE_RESOLVER_DIALOG))
            .addDialog(new EmailResolverDialog(EMAIL_RESOLVER_DIALOG))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.subjectStep.bind(this),
                this.attendeeStep.bind(this),
                this.emailStep.bind(this),
                this.meetingDateStep.bind(this),
                this.confirmStep.bind(this),
                this.finalStep.bind(this)
            ]));

        this.initialDialogId = WATERFALL_DIALOG;
    }

    /**
     * If a subject has not been provided, prompt for one.
     */
    async subjectStep(stepContext) {
        const meetingDetails = stepContext.options;

        if (!meetingDetails.subject) {
            const messageText = 'What is the subject for your meeting?';
            const msg = MessageFactory.text(messageText, messageText, InputHints.ExpectingInput);
            return await stepContext.prompt(TEXT_PROMPT, { prompt: msg });
        }
        return await stepContext.next(meetingDetails.subject);
    }

    /**
     * If an attendee has not been provided, prompt for one.
     */
    async attendeeStep(stepContext) {
        const meetingDetails = stepContext.options;

        // Capture the response to the previous step's prompt
        meetingDetails.subject = stepContext.result;
        if (!meetingDetails.attendee) {
            const messageText = 'With Whom would you like to meet?';
            const msg = MessageFactory.text(messageText, messageText, InputHints.ExpectingInput);
            return await stepContext.prompt(TEXT_PROMPT, { prompt: msg });
        }
        return await stepContext.next(meetingDetails.attendee);
    }
    /**
     * If an attendee has not been provided, prompt for one.
     */
    async emailStep(stepContext) {
        const meetingDetails = stepContext.options;

        // Capture the response to the previous step's prompt
        meetingDetails.attendee = stepContext.result;
        if (!meetingDetails.email || this.isBadEmail(meetingDetails.email)) {
           
            return await stepContext.beginDialog(EMAIL_RESOLVER_DIALOG, { email: meetingDetails.email, attendee: meetingDetails.attendee });
        }
        return await stepContext.next(meetingDetails.email);
    }
    /**
     * If a travel date has not been provided, prompt for one.
     * This will use the DATE_RESOLVER_DIALOG.
     */
    async meetingDateStep(stepContext) {
        const meetingDetails = stepContext.options;

        // Capture the results of the previous step
        meetingDetails.email = stepContext.result;
        if (!meetingDetails.meetingDateTime || this.isAmbiguous(meetingDetails.meetingDateTime)) {
            return await stepContext.beginDialog(DATE_RESOLVER_DIALOG, { date: meetingDetails.meetingDateTime });
        }
        return await stepContext.next(meetingDetails.meetingDateTime);
    }

    /**
     * Confirm the information the user has provided.
     */
    async confirmStep(stepContext) {
        const meetingDetails = stepContext.options;

        // Capture the results of the previous step
        meetingDetails.meetingDateTime = stepContext.result;
        const timeProperty = new TimexProperty(meetingDetails.meetingDateTime);
        meetingDetails.meetingDateMsg = timeProperty.toNaturalLanguage(new Date(Date.now()));
        const messageText = `Please confirm, I will setup a ${ meetingDetails.subject } with ${ meetingDetails.attendee } on ${ meetingDetails.meetingDateMsg}. Is this correct?`;
        const msg = MessageFactory.text(messageText, messageText, InputHints.ExpectingInput);

        // Offer a YES/NO prompt.
        return await stepContext.prompt(CONFIRM_PROMPT, { prompt: msg });
    }

    /**
     * Complete the interaction and end the dialog.
     */
    async finalStep(stepContext) {
        if (stepContext.result === true) {
            const bookingDetails = stepContext.options;
            return await stepContext.endDialog(bookingDetails);
        }
        return await stepContext.endDialog();
    }

    isAmbiguous(timex) {
        const timexPropery = new TimexProperty(timex);
        const valid = timexPropery.types.has('definite') && timexPropery.types.has('datetime');
        return !valid;
    }

    isBadEmail(mail) 
    {
    const mailformat = /(?!.*\.{2})^([a-z\d!#$%&'*+\-\/=?^_`{|}~\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]+(\.[a-z\d!#$%&'*+\-\/=?^_`{|}~\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]+)*|"((([ \t]*\r\n)?[ \t]+)?([\x01-\x08\x0b\x0c\x0e-\x1f\x7f\x21\x23-\x5b\x5d-\x7e\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]|\\[\x01-\x09\x0b\x0c\x0d-\x7f\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))*(([ \t]*\r\n)?[ \t]+)?")@(([a-z\d\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]|[a-z\d\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF][a-z\d\-._~\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]*[a-z\d\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])\.)+([a-z\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]|[a-z\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF][a-z\d\-._~\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]*[a-z\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])\.?$/i;
       
    if (mailformat.test(mail))
      {
        return (false)
      }
       
        return (true)
    }
    
}

module.exports.MeetingDialog = MeetingDialog;
