// Copyright (c) SmartThings Corporation. All rights reserved.
// Licensed under the MIT License.

const { InputHints, MessageFactory } = require('botbuilder');
const { ConfirmPrompt, OAuthPrompt, WaterfallDialog, ChoicePrompt } = require('botbuilder-dialogs');
const { CancelAndHelpDialog } = require('./cancelAndHelpDialog');

const INSTALLATION_CONTEXT_PROMPT = 'installationContextPrompt';
const AUTH_TYPE_CONTEXT_PROMPT = 'authTypePrompt';
const AUTHORIZE_OAUTH_PROMPT = 'authorizeOauthPrompt';
const CONFIRM_PROMPT = 'confirmPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

class AuthorizeDialog extends CancelAndHelpDialog {
    constructor(id, insightsClient) {
        super(id || 'authorizeDialog');

        // Add ApplicationInsights as telemetry client
        this.telemetryClient = insightsClient;

        this.addDialog(new ChoicePrompt(INSTALLATION_CONTEXT_PROMPT))
            .addDialog(new ChoicePrompt(AUTH_TYPE_CONTEXT_PROMPT))
            .addDialog(new OAuthPrompt(AUTHORIZE_OAUTH_PROMPT,  {
                connectionName: process.env.BotOAuthConnectionName,
                text: 'Powered by SmartThings',
                title: 'Sign In',
                timeout: 300000
            }))
            .addDialog(new ConfirmPrompt(CONFIRM_PROMPT))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.installationContextStep.bind(this),
                this.authTypeStep.bind(this),
                this.confirmStep.bind(this),
                this.authorizeStep.bind(this),
                this.finalStep.bind(this)
            ]));

        this.initialDialogId = WATERFALL_DIALOG;
    }

    /**
     * If an installation context has not been provided, prompt for one.
     */
    async installationContextStep(stepContext) {
        const authorizeDetails = stepContext.options;
        if (!authorizeDetails.installationContext) {
            return await stepContext.prompt(INSTALLATION_CONTEXT_PROMPT, {
                prompt: 'Who should be able to access this SmartThings location?',
                retryPrompt: 'Sorry, please choose from the list.',
                choices: ['Everbody in this workspace', 'Only full members', 'Just me'],
            });
        }
    }

    /**
     * If an authorization type has not been provided, prompt for one.
     */
    async authTypeStep(stepContext) {
        const authorizeDetails = stepContext.options;
        authorizeDetails.installationContext = stepContext.result;
        if (!authorizeDetails.authType) {
            return await stepContext.prompt(AUTH_TYPE_CONTEXT_PROMPT, {
                prompt: 'How do you want to authorize your account',
                retryPrompt: 'Sorry, please choose from the list.',
                choices: ['Personal Access Token (all available locations)', 'OAuth 2.0 (single location)'],
            });
        }
    }

    /**
     * Confirm the information the user has provided.
     */
    async confirmStep(stepContext) {
        const authorizeDetails = stepContext.options;

        // Capture the results of the previous step
        authorizeDetails.authType = stepContext.result;
        let installContext, authType;

        const getInstallContext = switchcase({
            0: 'everybody',
            1: 'only full members',
            2: 'just yourself'
        })('unknown');
        installContext = getInstallContext(authorizeDetails.installationContext.index);

        const getAuthType = switchcase({
            0: 'a SmartThings personal access token',
            1: 'a SmartThings OAuth 2.0 authorization'
        })('unknown');
        authType = getAuthType(authorizeDetails.authType.index);

        const messageText = `Please confirm, you want to authorize a location for ${ installContext }, using ${ authType }. Is this correct?`;
        const msg = MessageFactory.text(messageText, messageText, InputHints.ExpectingInput);

        // Offer a YES/NO prompt.
        return await stepContext.prompt(CONFIRM_PROMPT, { prompt: msg });
    }

    /**
     * Prompt to authorize with an OAuthCard.
     */
    async authorizeStep(stepContext) {
        if (stepContext.result === true) {
            const authorizeDetails = stepContext.options;   
            if (authorizeDetails.authType.index === 1) {
                console.log("Authorizing an OAuth 2.0 connection");
                return await stepContext.prompt(AUTHORIZE_OAUTH_PROMPT);
            } else {
                console.log("Authorizing a personal access token connection");
                // TODO
            }
        }
        return await stepContext.continueDialog(stepContext);
    }

    /**
     * Complete the interaction and end the dialog.
     */
    async finalStep(stepContext) {
        if (stepContext.result.token) {
            const authorizeDetails = stepContext.options;
            authorizeDetails.authorization = stepContext.result;
            return await stepContext.endDialog(authorizeDetails);
        }
        return await stepContext.endDialog();
    }
}

const switchcase = cases => defaultCase => key =>
  cases.hasOwnProperty(key) ? cases[key] : defaultCase

module.exports.AuthorizeDialog = AuthorizeDialog;
