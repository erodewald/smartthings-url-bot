// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { LuisRecognizer } = require('botbuilder-ai');

class SmartThingsAuthorizeRecognizer {
    constructor(config) {
        const luisIsConfigured = config && config.applicationId && config.endpointKey && config.endpoint;
        if (luisIsConfigured) {
            this.recognizer = new LuisRecognizer(config, {}, true);
        }
    }

    get isConfigured() {
        return (this.recognizer !== undefined);
    }

    /**
     * Returns an object with preformatted LUIS results for the bot's dialogs to consume.
     * @param {TurnContext} context
     */
    async executeLuisQuery(context) {
        return await this.recognizer.recognize(context);
    }
    
    getRoom(result) {
        let room;
        let { entities } = result;
        let { $instance } = entities;

        if (entities.SmartThings_Entities[0][0] &&
            entities.SmartThings_Entities[0][0] === 'Room' &&
            $instance.SmartThings_Entities[0].text) {
            room = $instance.SmartThings_Entities[0].text;
        }

        return { room: room };
    }

    getCapability(result) {
        let capability, input;
        let { entities } = result;
        let { $instance } = entities;
        if ($instance.SmartThings_Capability[0]) {
            input = $instance.SmartThings_Capability[0].text
        }
        if (entities.SmartThings_Capability[0]) {
            capability = entities.SmartThings_Capability[0];
        }
        return { text: input, capability: capability };
    }
}

module.exports.SmartThingsAuthorizeRecognizer = SmartThingsAuthorizeRecognizer;
