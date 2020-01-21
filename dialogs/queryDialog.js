// Copyright (c) SmartThings Corporation. All rights reserved.
// Licensed under the MIT License.

const axios = require('axios').default;
const { InputHints, MessageFactory } = require('botbuilder');
const { ConfirmPrompt, OAuthPrompt, WaterfallDialog, ChoicePrompt } = require('botbuilder-dialogs');
const { CancelAndHelpDialog } = require('./cancelAndHelpDialog');

const WATERFALL_DIALOG = 'waterfallDialog';
const OAUTH_PROMPT = 'oauthPrompt';

class QueryDialog extends CancelAndHelpDialog {
    constructor(id, insightsClient) {
        super(id || 'queryDialog');

        // Add ApplicationInsights as telemetry client
        this.telemetryClient = insightsClient;

        this.addDialog(new OAuthPrompt(OAUTH_PROMPT, {
            connectionName: process.env.BotOAuthConnectionName,
            text: 'Powered by SmartThings',
            title: 'Sign In',
            timeout: 300000
        }))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.commandStep.bind(this),
                this.processStep.bind(this),
            ]));

        this.initialDialogId = WATERFALL_DIALOG;
    }

    async commandStep(step) {
        step.values.command = step.result;

        // Call the prompt again because we need the token. The reasons for this are:
        // 1. If the user is already logged in we do not need to store the token locally in the bot and worry
        // about refreshing it. We can always just call the prompt again to get the token.
        // 2. We never know how long it will take a user to respond. By the time the
        // user responds the token may have expired. The user would then be prompted to login again.
        //
        // There is no reason to store the token locally in the bot because we can always just call
        // the OAuth prompt to get the token or get a new token if needed.
        return await step.beginDialog(OAUTH_PROMPT);
    }

    async processStep(step) {
        const dialogResponse = {};
        if (step.result) {
            // We do not need to store the token in the bot. When we need the token we can
            // send another prompt. If the token is valid the user will not need to log back in.
            // The token will be available in the Result property of the task.
            const tokenResponse = step.result;
            const { options } = step;
            let result = [];

            // If we have the token use the user is authenticated so we may use it to make API calls.
            if (tokenResponse && tokenResponse.token) {
                dialogResponse.token = tokenResponse.token;

                let { data: locations } = await axios.get('https://api.smartthings.com/v1/locations', { headers: { Authorization: `Bearer ${tokenResponse.token}` } });
                let { data: rooms } = await axios.get(`https://api.smartthings.com/v1/locations/${locations.items[0].locationId}/rooms`, { headers: { Authorization: `Bearer ${tokenResponse.token}` } });
                let room = rooms.items.find(room => {
                    return room.name.toLowerCase().includes(options.room.toLowerCase())
                });
                let roomId = room.roomId;

                let { data: devices } = await axios.get(`https://api.smartthings.com/v1/devices?capability=${options.capability}`, { headers: { Authorization: `Bearer ${tokenResponse.token}` } });

                const getDeviceStatuses = async () => {
                    return Promise.all(devices.items.map(async device => {
                        if (device.roomId == roomId) {
                            let { data: status } = await axios.get(`https://api.smartthings.com/v1/devices/${device.deviceId}/components/main/capabilities/${options.capability}/status`, { headers: { Authorization: `Bearer ${tokenResponse.token}` } });
                            result.push({ device, status });
                        }
                    }));
                };
                await getDeviceStatuses();
                let averageTemp = Math.round(result.reduce((total, next) => total + next.status.temperature.value, 0) / result.length);
                await step.context.sendActivity(`Average temp reading in ${room.name} is ${averageTemp}${result[0].status.temperature.unit}`);
            }
        } else {
            await step.context.sendActivity('We couldn\'t log you in. Please try again later.');
        }

        return await step.endDialog();
    }
}

const switchcase = cases => defaultCase => key =>
  cases.hasOwnProperty(key) ? cases[key] : defaultCase

module.exports.QueryDialog = QueryDialog;
