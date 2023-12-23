import axios from 'axios';
import { ActivityHandler, MessageFactory, Attachment } from 'botbuilder';
import { getEnv } from '@src/utils';
import { makeLogger } from '@src/logging';
import { isEmpty } from '@src/utils';
const logger = makeLogger('q-teams-bot');

import { qChatSync, QAttachment } from '@src/helpers/amazon-q/amazon-q-client';
import {
  getChannelMetadata,
  saveChannelMetadata,
  saveMessageMetadata
} from '@src/helpers/cache/cache';

export const ERROR_MSG = '***Processing error***';

//const MAX_FILE_ATTACHMENTS = 5;
const SUPPORTED_FILE_TYPES = [
  'txt',
  'html',
  'xml',
  'markdown',
  'csv',
  'json',
  'xls',
  'xlsx',
  'ppt',
  'pptx',
  'doc',
  'docx',
  'rtf',
  'pdf'
];

const attachFiles = async (attachments: Attachment[]): Promise<QAttachment[]> => {
  const qAttachments: QAttachment[] = [];
  for (const a of attachments) {
    // Process downloadable attachments
    const downloadUrl = a.content?.downloadUrl;
    if (!isEmpty(downloadUrl)) {
      logger.debug(`Processing attachment: ${JSON.stringify(a)}`);
      // Check if the file type is supported
      const fileType = a.content?.fileType;
      if (!isEmpty(fileType) && SUPPORTED_FILE_TYPES.includes(fileType) && !isEmpty(a.name)) {
        qAttachments.push({
          name: a.name,
          data: await retrieveAttachment(downloadUrl)
        });
      } else {
        logger.debug(
          `Ignoring file attachment with unsupported filetype '${fileType}' - not one of '${SUPPORTED_FILE_TYPES}'`
        );
      }
    }
  }
  return qAttachments;
};

export const retrieveAttachment = async (url: string) => {
  const response = await axios.get(url, {
    responseType: 'arraybuffer' // Important for handling binary files
  });

  // log just enough of the attachment content to validate file contents when troubleshooting.
  logger.debug(
    `retrieveAttachment from ${url}:\nData: ${response.data
      .slice(0, 300)
      .toString()
      .replace(/\r?\n/g, '')}`
  );
  return response.data;
};

export class QTeamsBot extends ActivityHandler {
  constructor() {
    super();
    // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
    this.onMessage(async (context) => {
      const env = getEnv(process.env);
      logger.debug(`Context: ${JSON.stringify(context)}`);
      const activity = context.activity;
      const type = context.activity.conversation.conversationType;

      const message = context.activity.text;
      logger.info(`Message received: ${message}`);
      const qAttachments: QAttachment[] = [];

      // We cache previous Amazon Q context metadata for personal DM channel
      let channelKey = '';
      let qContext = { conversationId: '', parentMessageId: '' };
      if (type === 'personal') {
        channelKey = `${activity.from.id}:${activity.channelData.tenant.id}`;
        const channelMetadata = await getChannelMetadata(channelKey, env);
        logger.debug(
          `ChannelKey: ${channelKey}, Cached channel metadata: ${JSON.stringify(channelMetadata)} `
        );
        qContext = {
          conversationId: channelMetadata?.conversationId,
          parentMessageId: channelMetadata?.systemMessageId
        };
        // add any attachments
        if (!isEmpty(activity?.attachments)) {
          qAttachments.push(...(await attachFiles(activity.attachments)));
        }
      }

      // For at mentions in a channel, we don't cache metadata, since we need to always pull thread history
      // which includes messages betwen other participants in the channel, not previously seen by the bot.
      // TODO

      // call ChatSync API
      const output = await qChatSync(env, message, qAttachments, qContext);
      if (output instanceof Error) {
        const replyText = `${ERROR_MSG} : *${output.message}*`;
        await context.sendActivity(MessageFactory.text(replyText, replyText));
        return;
      }

      // return response to user
      const replyText = output.systemMessage;
      await context.sendActivity(MessageFactory.text(replyText, replyText));

      // save metadata from sucessful response
      if (type === 'personal') {
        await saveChannelMetadata(channelKey, output.conversationId, output.systemMessageId, env);
      }
      await saveMessageMetadata(output, env);
    });
  }
}
