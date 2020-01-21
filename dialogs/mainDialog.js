// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const axios = require('axios');
const { MessageFactory, InputHints } = require('botbuilder');
const { LuisRecognizer } = require('botbuilder-ai');
const { ComponentDialog, DialogSet, DialogTurnStatus, TextPrompt, WaterfallDialog } = require('botbuilder-dialogs');

const MAIN_WATERFALL_DIALOG = 'mainWaterfallDialog';

class MainDialog extends ComponentDialog {
    constructor(luisRecognizer, authorizeDialog, queryDialog, insightsClient) {
        super('MainDialog');
        
        // Add ApplicationInsights as the telemetry client
        this.telemetryClient = insightsClient;

        if (!luisRecognizer) throw new Error('[MainDialog]: Missing parameter \'luisRecognizer\' is required');
        this.luisRecognizer = luisRecognizer;

        if (!authorizeDialog) throw new Error('[MainDialog]: Missing parameter \'authorizeDialog\' is required');
        if (!queryDialog) throw new Error('[MainDialog]: Missing parameter \'queryDialog\' is required');

        // Define the main dialog and its related components.
        // This is a sample "book a flight" dialog.
        this.addDialog(new TextPrompt('TextPrompt'))
            .addDialog(authorizeDialog)
            .addDialog(queryDialog)
            .addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
                this.introStep.bind(this),
                this.actStep.bind(this),
                this.finalStep.bind(this)
            ]));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;
    }

    /**
     * The run method handles the incoming activity (in the form of a TurnContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {*} turnContext
     * @param {*} accessor
     */
    async run(turnContext, accessor) {
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);

        const dialogContext = await dialogSet.createContext(turnContext);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    /**
     * First step in the waterfall dialog. Prompts the user for a command.
     * Currently, this expects a booking request, like "book me a flight from Paris to Berlin on march 22"
     * Note that the sample LUIS model will only recognize Paris, Berlin, New York and London as airport cities.
     */
    async introStep(stepContext) {
        if (!this.luisRecognizer.isConfigured) {
            const messageText = 'NOTE: LUIS is not configured. To enable all capabilities, add `LuisAppId`, `LuisAPIKey` and `LuisAPIHostName` to the .env file.';
            await stepContext.context.sendActivity(messageText, null, InputHints.IgnoringInput);
            return await stepContext.next();
        }

        const messageText = stepContext.options.restartMsg ? stepContext.options.restartMsg : 'What can I help you with today?\nSay something like "What\'s going on in Apollo right now?"';
        const promptMessage = MessageFactory.text(messageText, messageText, InputHints.ExpectingInput);
        return await stepContext.prompt('TextPrompt', { prompt: promptMessage });
    }

    /**
     * Second step in the waterfall.  This will use LUIS to attempt to extract data.
     * Then, it hands off to the authorizeDialog child dialog to collect any remaining details.
     */
    async actStep(stepContext) {
        const authorizeDetails = {};

        if (!this.luisRecognizer.isConfigured) {
            // LUIS is not configured, we just run the AuthorizeDialog path.
            return await stepContext.beginDialog('authorizeDialog', {});
        }

        // Call LUIS and gather any potential booking details. (Note the TurnContext has the response to the prompt)
        const luisResult = await this.luisRecognizer.executeLuisQuery(stepContext.context);
        switch (LuisRecognizer.topIntent(luisResult)) {

            case 'SmartThings_Authorize': {
                return await stepContext.beginDialog('authorizeDialog', authorizeDetails);
            }

            case 'SmartThings_QueryState': {
                const queryDetails = {};
                // Extract the values for the composite entities from the LUIS result.
                const roomEntities = this.luisRecognizer.getRoom(luisResult);
                const capabilityEntities = this.luisRecognizer.getCapability(luisResult);

                queryDetails.room = roomEntities.room;
                queryDetails.capability = capabilityEntities.capability;
                console.log('LUIS extracted these details:', JSON.stringify(queryDetails));

                return await stepContext.beginDialog('queryDialog', queryDetails);
            }

            case 'SmartThings_CheckOccupancy': {
                return await stepContext.beginDialog('checkOccupancyDialog', authorizeDetails);
            }

            default: {
                // Catch all for unhandled intents
                const didntUnderstandMessageText = `Sorry, I didn't get that. Please try asking in a different way (intent was ${ LuisRecognizer.topIntent(luisResult) })`;
                await stepContext.context.sendActivity(didntUnderstandMessageText, didntUnderstandMessageText, InputHints.IgnoringInput);
            }
        }

        return await stepContext.next();
    }

    /**
     * This is the final step in the main waterfall dialog.
     * It wraps up the sample "authorize" interaction with a simple confirmation.
     */
    async finalStep(stepContext) {
        // If the child dialog ("authorizeDialog") was cancelled or the user failed to confirm, the Result here will be null.
        if (stepContext.result) {
            const result = stepContext.result;
            const {token} = result.authorization;
            // Now we have all the authorize details.
            // This is where calls to the authorize service would go.
            // If the call to the authorize service was successful tell the user.
            let { data: locs } = await axios.get('https://api.smartthings.com/v1/locations', {headers:{Authorization: `Bearer ${token}`}})
            const msg = `I connected your SmartThings location ${ locs.items[0].name }.`;

            await stepContext.context.sendActivity(msg, msg, InputHints.IgnoringInput);
        }

        // Restart the main dialog with a different message the second time around
        return await stepContext.replaceDialog(this.initialDialogId, { restartMsg: 'What else can I do for you?' });
    }
}

module.exports.MainDialog = MainDialog;
