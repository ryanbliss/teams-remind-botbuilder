// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// Import required Bot Framework classes.
import {
    Activity,
    CardFactory,
    ConversationReference,
    InvokeResponse,
    StatusCodes,
    TeamsActivityHandler,
    TeamsInfo,
    TurnContext
} from 'botbuilder';
import server from '.';

export class SampleBot extends TeamsActivityHandler {
    public conversationReferences: Map<string, Partial<ConversationReference>> = new Map();
    constructor() {
        super();

        this.onMessage(async (context, next) => {
            this.conversationReferences.set(
                context.activity.conversation.id,
                TurnContext.getConversationReference(context.activity)
            );

            const text = context.activity.text.toLowerCase();
            if (text.includes('remind')) {
                await this.sendRemindCard(context);
            } else if (text.includes('help')) {
                await this.sendIntroCard(context);
            } else {
                await context.sendActivity(
                    `Use 'remind' to schedule a reminder, or use 'help' for more information.`
                );
            }

            await next();
        });

        // Sends welcome messages to conversation mentions when they join the conversation.
        // Messages are only sent to conversation mentions who aren't the bot.
        this.onMembersAdded(async (context, next) => {
            if (!context.activity.membersAdded) return;
            for (let idx = 0; idx < context.activity.membersAdded.length; idx++) {
                if (context.activity.membersAdded[idx].id !== context.activity.recipient.id) {
                    await this.sendIntroCard(context);
                }
            }

            await next();
        });

        this.onInstallationUpdateAdd(async (context, next) => {
            // Set conversation reference for proactive messaging
            this.conversationReferences.set(
                context.activity.conversation.id,
                TurnContext.getConversationReference(context.activity)
            );
            // Send intro card
            await this.sendIntroCard(context);

            await next();
        });
    }

    async onInvokeActivity(context: TurnContext): Promise<InvokeResponse> {
        if (isIActivityAction(context.activity)) {
            if (context.activity.value?.action.verb === 'schedule_reminder') {
                await this.scheduleReminder(context.activity);
                return {
                    status: StatusCodes.OK,
                    body: {
                        statusCode: StatusCodes.OK,
                        type: 'application/vnd.microsoft.card.adaptive',
                        value: this.buildScheduledReminderCard(context.activity)
                    }
                };
            }
        }
        return {
            status: StatusCodes.INTERNAL_SERVER_ERROR,
            body: {
                statusCode: StatusCodes.INTERNAL_SERVER_ERROR
            }
        };
    }

    private async sendIntroCard(context: TurnContext) {
        const adaptiveCard = {
            $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
            version: '1.5',
            body: [
                {
                    type: 'Image',
                    url: 'https://aka.ms/bf-welcome-card-image',
                    size: 'Stretch'
                },
                {
                    type: 'TextBlock',
                    weight: 'Bolder',
                    size: 'Medium',
                    text: `Welcome, ${context.activity.from.name}!`
                },
                {
                    type: 'TextBlock',
                    size: 'Small',
                    text: `Use the 'remind' command to schedule a reminder.`
                }
            ],
            actions: [
                {
                    type: 'Action.OpenUrl',
                    title: 'Get an overview',
                    url: 'https://docs.microsoft.com/en-us/azure/bot-service/?view=azure-bot-service-4.0'
                }
            ]
        };

        await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(adaptiveCard)]
        });
    }

    private async sendRemindCard(context: TurnContext) {
        const pagedMembers = await TeamsInfo.getPagedMembers(context);
        const adaptiveCard = {
            $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
            type: 'AdaptiveCard',
            version: '1.5',
            body: [
                {
                    type: 'TextBlock',
                    size: 'Medium',
                    weight: 'Bolder',
                    label: "Reminder time",
                    text: `Create a reminder`
                },
                {
                    type: 'Input.ChoiceSet',
                    id: 'delaySeconds',
                    style: 'compact',
                    value: '5',
                    label: "Remind after",
                    choices: [
                        { title: '5 seconds', value: '5' },
                        { title: '15 seconds', value: '15' },
                        { title: '30 seconds', value: '30' }
                    ]
                },
                {
                    type: 'Input.ChoiceSet',
                    id: 'mention',
                    style: 'compact',
                    value: undefined,
                    label: "Who to @mention",
                    placeholder: 'Select a person',
                    choices: pagedMembers.members.map((mention) => ({
                        title: mention.name,
                        value: JSON.stringify({
                            id: mention.id,
                            name: mention.name
                        })
                    }))
                }
            ],
            actions: [
                {
                    type: 'Action.Execute',
                    verb: 'schedule_reminder',
                    title: 'Schedule reminder'
                }
            ]
        };

        await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(adaptiveCard)]
        });
    }

    private buildScheduledReminderCard(activity: IActivityAction) {
        const mention = JSON.parse(activity.value.action.data?.mention);
        const adaptiveCard = {
            $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
            type: 'AdaptiveCard',
            version: '1.5',
            body: [
                {
                    type: 'TextBlock',
                    size: 'Small',
                    text: `Reminder scheduled in ${activity.value.action.data?.delaySeconds} seconds for ${mention.name}`
                }
            ]
        };

        return adaptiveCard;
    }

    private async scheduleReminder(activity: IActivityAction) {
        await fetch(server.url + '/api/remind', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                conversationId: activity.conversation.id,
                delaySeconds: parseInt(activity.value.action.data?.delaySeconds),
                mention: JSON.parse(activity.value.action.data?.mention)
            })
        });
    }
}

interface IData {
    [key: string]: any;
}

interface IActivityAction extends Activity {
    name: 'adaptiveCard/action';
    value: {
        action: {
            verb: string;
            type: string;
            title: string;
            data?: IData;
        };
    };
}
function isIActivityAction(value: any): value is IActivityAction {
    return (
        value.name === 'adaptiveCard/action' &&
        (value as IActivityAction).value !== undefined &&
        (value as IActivityAction).value.action !== undefined &&
        typeof (value as IActivityAction).value.action.verb === 'string'
    );
}
