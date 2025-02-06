// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import * as restify from 'restify';

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication,
    ConfigurationServiceClientCredentialFactory,
    MessageFactory
} from 'botbuilder';

// This bot's main dialog.
import { SampleBot } from './bot';
import { IReminder, isIReminder } from './api-parser';

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${server.name} listening to ${server.url}`);
});

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
    {},
    new ConfigurationServiceClientCredentialFactory({
        MicrosoftAppId: process.env.BOT_ID,
        MicrosoftAppPassword: process.env.BOT_PASSWORD,
        MicrosoftAppType: 'MultiTenant'
    })
);

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
adapter.onTurnError = async (context, error) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    console.error(`\n [onTurnError] unhandled error: ${error}`);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${error}`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    // Send a message to the user
    await context.sendActivity('The bot encounted an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

// Create the main dialog.
const myBot = new SampleBot();

// Listen for incoming requests.
server.post('/api/messages', (req, res, next) => {
    // Route received a request to adapter for processing
    adapter.process(req, res, async (context) => await myBot.run(context));
});

async function remindWithDelay(body: IReminder) {
    const conversationReference = myBot.conversationReferences.get(body.conversationId);
    if (!conversationReference) {
        return;
    }
    // wait for the number of seconds before sending the proactive message
    await new Promise((resolve) => setTimeout(resolve, body.delaySeconds * 1000));

    const mention = {
        mentioned: body.mention,
        text: `<at>${body.mention.name}</at>`,
        type: 'mention'
    };

    const message = MessageFactory.text(`This is a reminder for <at>${body.mention.name}</at>!`);
    message.entities = [mention];
    
    await adapter.continueConversationAsync(process.env.BOT_ID!, conversationReference, async (context) => {
        await context.sendActivity(message);
    });
}

server.post('/api/remind', async (req, res) => {
    if (!isIReminder(req.body)) {
        res.writeHead(400);
        res.end();
        return;
    }
    const conversationReference = myBot.conversationReferences.get(req.body.conversationId);
    if (!conversationReference) {
        res.writeHead(400);
        res.end();
        return;
    }

    res.setHeader('Content-Type', 'text/html');
    res.writeHead(200);
    res.write('<html><body><h1>Proactive messages have been sent.</h1></body></html>');
    res.end();

    remindWithDelay(req.body);
});

export default server;
