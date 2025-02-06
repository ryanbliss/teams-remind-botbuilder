export interface IReminder {
    conversationId: string;
    delaySeconds: number;
    mention: {
        id: string;
        name: string;
    };
}

export function isIReminder(value: any): value is IReminder {
    return (
        !!value &&
        typeof value.conversationId === 'string' &&
        typeof value.delaySeconds === 'number' &&
        typeof value.mention === 'object'
    );
}
