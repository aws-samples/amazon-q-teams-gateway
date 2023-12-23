import axios from 'axios';
import { ActivityHandler, MessageFactory, Attachment } from 'botbuilder';
import { getEnv } from '@src/utils';
import { makeLogger } from '@src/logging';
import { isEmpty, getTeamsSecret } from '@src/utils';
const logger = makeLogger('q-teams-bot');
import { qChatSync, QAttachment } from '@src/helpers/amazon-q/amazon-q-client';
import {
  getChannelMetadata,
  saveChannelMetadata,
  deleteChannelMetadata,
  saveMessageMetadata
} from '@src/helpers/cache/cache';
export const ERROR_MSG = '***Processing error***';
let oathToken = '';

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

async function getOathToken(tenantId: string): Promise<string> {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  const teamsSecret = await getTeamsSecret();
  params.append('client_id', teamsSecret.MicrosoftAppId);
  params.append('client_secret', teamsSecret.MicrosoftAppPassword);
  params.append('scope', 'https://graph.microsoft.com/.default');
  params.append('grant_type', 'client_credentials');
  try {
    const response = await axios.post(url, params.toString(), {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    });
    return response.data.access_token;
  } catch (error) {
    console.error('Error getting OAuth token', error);
    throw error;
  }
}

async function retrieveThreadHistory(
  teamId: string,
  channelId: string,
  threadId: string,
  oathToken: string
) {
  const url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages/${threadId}/replies`;
  try {
    const response = await axios.get(url, {
      headers: { Authorization: `Bearer ${oathToken}` }
    });
    return response.data.value;
  } catch (error) {
    logger.error('Error fetching messages:', error);
    logger.debug(`curl -X GET '${url}' -H 'Accept: application/json, text/plain, */*'   -H 'Authorization: Bearer ${oathToken}'`);
    throw error;
  }
}

export class QTeamsBot extends ActivityHandler {
  constructor() {
    super();
    // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
    this.onMessage(async (context) => {
      const tenantId = context.activity.conversation.tenantId;
      if (isEmpty(oathToken) && !isEmpty(tenantId)) {
        oathToken = await getOathToken(tenantId);
      }
      const env = getEnv(process.env);
      logger.debug(`Context: ${JSON.stringify(context)}`);
      const activity = context.activity;
      const type = context.activity.conversation.conversationType;

      const message = context.activity.text;
      logger.info(`Message received: ${message}`);
      const qAttachments: QAttachment[] = [];

      // We cache previous Amazon Q context metadata for personal DM channel
      let channelKey = '';
      let qContext = undefined;

      // DM (personal) messages
      if (type === 'personal') {
        channelKey = `${activity.from.id}:${activity.channelData.tenant.id}`;

        // check if DM message is a reset (/new_conversation or /new_context) command
        if (message.startsWith('/new_con')) {
          logger.debug(
            `Slash command: /new_conversation - deleting channel metadata for '${channelKey}'`
          );
          await deleteChannelMetadata(channelKey, env);
          const replyText = `_*Starting New Conversation*_`;
          await context.sendActivity(MessageFactory.text(replyText, replyText));
          return;
        }

        // it's not a reset command, so process the message
        // get any cached context metadata
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
      // Team channel message (@mention)
      else if (type === 'channel') {
        // we don't cache context metadata for For @mentions in a channel. We alway invoke AmazonQ with
        // a new conversation context, which embeds the thread history in the input message.
        // The thread history includes messages not previously seen by the bot.

        const threadHistory = await retrieveThreadHistory(
          activity.conversation.id,
          activity.channelData.channel.id,
          activity.conversation.id,
          oathToken
        );

        logger.debug(`Thread history: ${JSON.stringify(threadHistory)}`);
      }

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
