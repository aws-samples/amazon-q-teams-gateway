import { ActivityHandler, MessageFactory } from 'botbuilder';
import { getEnv } from '@src/utils';
import { makeLogger } from '@src/logging';
const logger = makeLogger('q-teams-bot');

import { qChatSync, Attachment } from '@src/helpers/amazon-q/amazon-q-client';
import {
  getChannelKey,
  getChannelMetadata,
  saveChannelMetadata,
  saveMessageMetadata
} from '@src/helpers/cache/cache';

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

      // get metadata for existing conversation, if any
      const channelKey = getChannelKey(
        context.activity.conversation.conversationType,
        context.activity.from.id,
        context.activity.channelData.tenant.id,
        context.activity.channelData.team.id,
        context.activity.channelData.channel.id,
        context.activity.timestamp
      );
      const channelMetadata = await getChannelMetadata(channelKey, env);
      logger.debug(
        `ChannelKey: ${channelKey}, Cached channel metadata: ${JSON.stringify(channelMetadata)} `
      );
      const qContext = {
        conversationId: channelMetadata?.conversationId,
        parentMessageId: channelMetadata?.systemMessageId
      };

      // get attachments, if any
      const attachments: Attachment[] = [];

      // call ChatSync API
      const output = await qChatSync(env, message, attachments, qContext);
      if (output instanceof Error) {
        const replyText = `${ERROR_MSG} : *${output.message}*`;
        await context.sendActivity(MessageFactory.text(replyText, replyText));
        return;
      }

      // save metadata from sucessful response
      await Promise.all([
        saveChannelMetadata(
          channelKey,
          output.conversationId,
          output.systemMessageId,
          env
        ),
        saveMessageMetadata(output, env),
      ]);

      // return response to user
      const replyText = output.systemMessage;
      await context.sendActivity(MessageFactory.text(replyText, replyText));
    });
  }
}
