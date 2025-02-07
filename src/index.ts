import * as restify from 'restify';
import {
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication,
    ConfigurationServiceClientCredentialFactory,
    MessageFactory
} from 'botbuilder';

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

const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError] unhandled error: ${error}`);

    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${error}`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    await context.sendActivity('The bot encounted an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

const myBot = new SampleBot();

server.post('/api/messages', (req, res, next) => {
    adapter.process(req, res, async (context) => await myBot.run(context));
    next();
});

async function remindWithDelay(body: IReminder) {
    const conversationReference = myBot.conversationReferences.get(body.conversationId);
    if (!conversationReference) {
        return;
    }
    
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
