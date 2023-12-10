import { ActivityHandler, MessageFactory } from 'botbuilder';
import { getEnv } from '@src/utils';

import { makeLogger } from '@src/logging';
const logger = makeLogger('q-teams-bot');

import { qChatSync, Attachment } from '@src/helpers/amazon-q/amazon-q-client';

export const ERROR_MSG = '***Processing error***';

export class QTeamsBot extends ActivityHandler {
  constructor() {
    super();
    // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
    this.onMessage(async (context) => {
      const env = getEnv(process.env);
      logger.debug(`Context: ${JSON.stringify(context)}`);
      const message = context.activity.text;
      logger.info(`Message received: ${message}`);
      // Send a reply to the user.
      const attachments: Attachment[] = [];
      const output = await qChatSync(env, message, attachments);
      if (output instanceof Error) {
        const replyText = `${ERROR_MSG} : *${output.message}*`;
        await context.sendActivity(MessageFactory.text(replyText, replyText));
        return;
      }
      const replyText = output.systemMessage;
      await context.sendActivity(MessageFactory.text(replyText, replyText));
    });
  }
}
