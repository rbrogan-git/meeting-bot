// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { InputHints } = require('botbuilder');
const { ComponentDialog, DialogTurnStatus } = require('botbuilder-dialogs');

/**
 * This base class watches for common phrases like "help" and "cancel" and takes action on them
 * BEFORE they reach the normal bot logic.
 */
class CancelAndHelpDialog extends ComponentDialog {
    async onBeginDialog(innerDc, options) {
        const result = await this.interrupt(innerDc);
        if (result) {
            return result;
        }

        return await super.onBeginDialog(innerDc, options);
    }
    async onContinueDialog(innerDc) {
        const result = await this.interrupt(innerDc);
        if (result) {
            return result;
        }
        return await super.onContinueDialog(innerDc);
    }

    async interrupt(innerDc) {
        if (innerDc.context.activity.text) {
            const text = innerDc.context.activity.text.toLowerCase();

            switch (text) {
                case 'help':
                case '?': {
                    const helpMessageText = 'Show help here';
                    await innerDc.context.sendActivity(helpMessageText, helpMessageText, InputHints.ExpectingInput);
                    return { status: DialogTurnStatus.waiting };
                }
                case 'cancel':
                case 'quit': {
                    const cancelMessageText = 'Cancelling...';
                    await innerDc.context.sendActivity(cancelMessageText, cancelMessageText, InputHints.IgnoringInput);
                    return await innerDc.cancelAllDialogs();
                }
                case 'logout':{
                    // The bot adapter encapsulates the authentication processes.
                    const botAdapter = innerDc.context.adapter;
                    await botAdapter.signOutUser(innerDc.context, process.env.ConnectionName);
                    await innerDc.context.sendActivity('You have been signed out.');
                    return await innerDc.cancelAllDialogs();
                }
            }
        }
    }
}

module.exports.CancelAndHelpDialog = CancelAndHelpDialog;
