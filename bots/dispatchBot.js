
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TurnContext, ActivityTypes, ActivityHandler } = require('botbuilder');
const { LuisRecognizer, QnAMaker } = require('botbuilder-ai');

class DispatchBot extends ActivityHandler {
    constructor(qnaAppInsightsClient) {
        super();

        // Add ApplicationInsights as the telemetry client
        this.telemetryClient = qnaAppInsightsClient;

        // If the includeApiResults parameter is set to true, as shown below, the full response
        // from the LUIS api will be made available in the properties  of the RecognizerResult
        const dispatchRecognizer = new LuisRecognizer({
            applicationId: process.env.DispatchLuisAppId,
            endpointKey: process.env.DispatchLuisAPIKey,
            endpoint: `https://${process.env.DispatchLuisAPIHostName}`
        }, {
            includeAllIntents: true,
            includeInstanceData: true
        }, true);
        const qnaMaker = new QnAMaker({
            knowledgeBaseId: process.env.QnAKnowledgebaseId,
            endpointKey: process.env.QnAEndpointKey,
            host: process.env.QnAEndpointHostName
        }, this.telemetryClient);

        this.dispatchRecognizer = dispatchRecognizer;
        this.qnaMaker = qnaMaker;

        this.onMessage(async (context, next) => {
            console.log('Processing Message Activity.');

            const reference = TurnContext.getConversationReference(context.activity);
            const activity = TurnContext.applyConversationReference({ type: 'typing' }, reference);
            await context.adapter.sendActivities(context, [activity]);
            // await context.sendActivity({ type: ActivityTypes.Typing });

            // First, we use the dispatch model to determine which cognitive service (LUIS or QnA) to use.
            const recognizerResult = await this.dispatchRecognizer.recognize(context);

            // Top intent tell us which cognitive service to use.
            const intent = LuisRecognizer.topIntent(recognizerResult);

            // Next, we call the dispatcher with the top intent.
            await this.dispatchToTopIntentAsync(context, intent, recognizerResult);

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

    async dispatchToTopIntentAsync(context, intent, recognizerResult) {
        switch (intent) {
        case 'l_SmartThings':
            await this.processOfficeAutomation(context, recognizerResult.luisResult);
            break;
        case 'q_SmartThings':
            await this.processChitChat(context);
            break;
        default:
            console.log(`Dispatch unrecognized intent: ${ intent }.`);
            break;
        }
    }

    async processOfficeAutomation(context, luisResult) {
        console.log('processOfficeAutomation');

        // Retrieve LUIS result for Process Automation.
        const result = luisResult.connectedServiceResult;
        const intent = result.topScoringIntent.intent;

        console.log(`OfficeAutomation top intent ${intent}.`);
        console.log(`OfficeAutomation intents detected:  ${luisResult.intents.map((intentObj) => intentObj.intent).join('\n\n')}.`);
        
        let blocks = {
            blocks:
                [
                    {
                        type: "section",
                        text: {
                            type: "mrkdwn",
                            text: "You have a new request:\n*<fakeLink.toEmployeeProfile.com|Fred Enriquez - New device request>*"
                        }
                    },
                    {
                        type: "section",
                        fields: [
                            {
                                type: "mrkdwn",
                                text: "*Type:*\nComputer (laptop)"
                            },
                            {
                                type: "mrkdwn",
                                text: "*When:*\nSubmitted Aut 10"
                            },
                            {
                                type: "mrkdwn",
                                text: "*Last Update:*\nMar 10, 2015 (3 years, 5 months)"
                            },
                            {
                                type: "mrkdwn",
                                text: "*Reason:*\nAll vowel keys aren't working."
                            },
                            {
                                type: "mrkdwn",
                                text: "*Specs:*\n\"Cheetah Pro 15\" - Fast, really fast\""
                            }
                        ]
                    },
                    {
                        type: "actions",
                        elements: [
                            {
                                type: "button",
                                text: {
                                    type: "plain_text",
                                    emoji: true,
                                    text: "Approve"
                                },
                                style: "primary",
                                value: "click_me_123"
                            },
                            {
                                type: "button",
                                text: {
                                    type: "plain_text",
                                    emoji: true,
                                    text: "Deny"
                                },
                                style: "danger",
                                value: "click_me_123"
                            }
                        ]
                    }
                ]
        };

        await context.sendActivity({ type: ActivityTypes.Message, channelData: blocks });
        // await context.sendActivity(`OfficeAutomation top intent ${intent}.`);
        // await context.sendActivity(`OfficeAutomation intents detected:  ${ luisResult.intents.map((intentObj) => intentObj.intent).join('\n\n') }.`);

        // if (luisResult.entities.length > 0) {
        //     await context.sendActivity(`OfficeAutomation entities were found in the message: ${ luisResult.entities.map((entityObj) => entityObj.entity).join('\n\n') }.`);
        // }
    }

    async processChitChat(context) {
        console.log('processSampleQnA');

        const results = await this.qnaMaker.getAnswers(context);
        if (results.length > 0) {
            await context.sendActivity(`${ results[0].answer }`);
        }
    }
}

module.exports.DispatchBot = DispatchBot;
