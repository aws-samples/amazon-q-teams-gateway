import axios from 'axios';
import { htmlToText } from 'html-to-text';
import * as marked from 'marked';
import { ConfidentialClientApplication, Configuration } from '@azure/msal-node';
import {
  ActivityHandler,
  MessageFactory,
  Attachment,
  TeamsInfo,
  TeamDetails,
  TurnContext,
  ActionTypes,
  CardFactory,
  CardAction,
  TaskModuleResponse,
  InvokeResponse
} from 'botbuilder';
import { getEnv } from '@src/utils';
import { makeLogger } from '@src/logging';
import { isEmpty, getTeamsSecret } from '@src/utils';
const logger = makeLogger('q-teams-bot');
import {
  qChatSync,
  qPutFeedbackRequest,
  QAttachment,
  AmazonQResponse,
  SourceAttribution
} from '@src/helpers/amazon-q/amazon-q-client';
import {
  getChannelMetadata,
  saveChannelMetadata,
  deleteChannelMetadata,
  saveMessageMetadata,
  getMessageMetadata
} from '@src/helpers/cache/cache';
export const ERROR_MSG = '***Processing error***';
const PROCESSING_MSG = '*Processing...*';
const FEEDBACK_ACK_MSG = '*Thank you for your feedback!*';
let oathToken = '';

const SUPPORTED_FILE_TYPES = [
  'txt',
  'html',
  'xml',
  'markdown',
  '.md',
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

async function setAmazonQUserId(context: TurnContext, userId: string | undefined) {
  if (isEmpty(userId)) {
    // Use Teams user email as Q UserId
    const userEmail = await getEmailAddress(context);
    logger.debug(
      `User's email (${userEmail}) used as Amazon Q userId, since AmazonQUserId is empty.`
    );
    return userEmail;
  }
  return userId;
}

async function getEmailAddress(context: TurnContext): Promise<string | undefined> {
  try {
    const teamsUser = await TeamsInfo.getMember(context, context.activity.from.id);
    return teamsUser.email;
  } catch (error) {
    console.error('Error getting user email:', error);
    return undefined;
  }
}

const attachFiles = async (
  attachments: Attachment[],
  oathToken: string,
  teamId?: string
): Promise<QAttachment[]> => {
  const qAttachments: QAttachment[] = [];
  for (const a of attachments) {
    // Process downloadable attachments
    // DM messages have a signed downloadUrl, but retrieved channel messages have only unsigned contentUrl
    let downloadUrl = a?.content?.downloadUrl;
    if (isEmpty(downloadUrl)) {
      if (!isEmpty(a.contentUrl) && !isEmpty(teamId)) {
        logger.debug(`Get download URL for attachment`);
        try {
          downloadUrl = await getDownloadUrl(a.contentUrl, oathToken, teamId);
        } catch (error) {
          logger.error(`Error retrieving attachment downloadUrl.. skipping: ${error}`);
        }
      }
    }
    if (!isEmpty(downloadUrl)) {
      logger.debug(`Processing attachment: ${JSON.stringify(a)}`);
      // Check if the file type is supported
      // Use name suffix, since retrieved channel messages have no fileType attribute
      const fileType = a.name?.split('.').pop();
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

// contentURL is not directly usable - app lacks permissions
// get downloadurl from microsoft graph
// see: https://github.com/microsoftgraph/msgraph-sdk-javascript/issues/200
async function getDownloadUrl(contentUrl: string, oathToken: string, teamId: string) {
  const url = new URL(contentUrl);
  const path = url.pathname.split('/').slice(4).join('/');
  const attachmentUrl = `https://graph.microsoft.com/v1.0/groups/${teamId}/drive/root:/${path}`;
  const attachment = await getAPIResponse(attachmentUrl, oathToken);
  return attachment['@microsoft.graph.downloadUrl'];
}

export const retrieveAttachment = async (downloadUrl: string) => {
  logger.debug(`curl -X GET '${downloadUrl}'`);
  const response = await axios.get(downloadUrl, {
    responseType: 'arraybuffer' // Important for handling binary files
  });
  // log just enough of the attachment content to validate file contents when troubleshooting.
  logger.debug(
    `retrieveAttachment from ${downloadUrl}:\nData: ${response.data
      .slice(0, 300)
      .toString()
      .replace(/\r?\n/g, '')}`
  );
  return response.data;
};

async function getOathToken(tenantId: string): Promise<string> {
  const teamsSecret = await getTeamsSecret();
  const msalConfig: Configuration = {
    auth: {
      authority: `https://login.microsoftonline.com/${tenantId}`,
      clientId: teamsSecret.MicrosoftAppId,
      clientSecret: teamsSecret.MicrosoftAppPassword
    }
  };
  const cca = new ConfidentialClientApplication(msalConfig);
  const tokenRequest = {
    // Use .default to request the static permissions configured in the portal
    scopes: ['https://graph.microsoft.com/.default']
  };
  try {
    const response = await cca.acquireTokenByClientCredential(tokenRequest);
    return response?.accessToken || '';
  } catch (error) {
    console.error('Error acquiring token:', error);
    throw error;
  }
}

interface ThreadMessages {
  name: string;
  message: string;
  date: string;
}

async function retrieveThreadHistory(
  context: TurnContext,
  oathToken: string
): Promise<[ThreadMessages[], QAttachment[]]> {
  const teamDetails: TeamDetails = await TeamsInfo.getTeamDetails(context);
  logger.debug(`Team Details: ${JSON.stringify(teamDetails)}`);
  const teamGUID = teamDetails.aadGroupId;
  const channelId = context.activity.channelData.teamsChannelId;

  // parse messageId from conversation if (eg 'xxx;messageid=<messageid>')
  const messageId = context.activity.conversation.id.split('messageid=')[1];
  if (isEmpty(teamGUID) || isEmpty(teamGUID) || isEmpty(messageId)) {
    throw new Error(
      `TeamGUID '${teamGUID}', channelId '${channelId}', or messageId '${messageId}' is empty`
    );
  }

  logger.debug('Retrieve first post');
  const firstPostUrl = `https://graph.microsoft.com/v1.0/teams/${teamGUID}/channels/${channelId}/messages/${messageId}`;
  const firstPost = await getAPIResponse(firstPostUrl, oathToken);
  logger.debug('Retrieve replies');
  const repliesUrl = `https://graph.microsoft.com/v1.0/teams/${teamGUID}/channels/${channelId}/messages/${messageId}/replies`;
  const replies = (await getAPIResponse(repliesUrl, oathToken)).value.reverse();
  // combine first post and replies
  const fullThread = [firstPost, ...replies];
  logger.debug(`Full Thread: ${JSON.stringify(fullThread)}`);
  // build conversation history JSON
  const threadMessages = [];
  for (const m of fullThread) {
    logger.debug(`Message: ${JSON.stringify(m)}`);
    let message = htmlToText(m.body.content);
    // Check if message is empty and attachments exist
    if (isEmpty(message) && m.attachments.length > 0) {
      const attachment0 = m.attachments[0];
      logger.debug(`Attachment0: ${JSON.stringify(m.attachments[0])}`);
      // Process content in card text if available
      if (!isEmpty(attachment0.content)) {
        try {
          const content = JSON.parse(attachment0.content);
          logger.debug(`Content: ${JSON.stringify(content)}`);
          if (content.text) {
            logger.debug(`Content.text: ${content.text}`);
            message = htmlToText(content.text);
          }
        } catch (error) {
          logger.error(
            `Error parsing JSON in attachment content: ${attachment0.content}: ${error}`
          );
        }
      }
    }
    if (!isEmpty(message) && message !== FEEDBACK_ACK_MSG) {
      threadMessages.push({
        name: m.from?.user?.displayName || m.from?.application?.displayName || '',
        message: message,
        date: m.createdDateTime
      });
    }
  }
  // get attachments from thread history
  const threadAttachments: QAttachment[] = [];
  for (const m of fullThread) {
    const teamId = m.channelIdentity.teamId;
    if (!isEmpty(m?.attachments)) {
      logger.debug(`Process attachments`);
      threadAttachments.push(...(await attachFiles(m.attachments, oathToken, teamId)));
    }
  }
  return [threadMessages, threadAttachments];
}

async function getAPIResponse(url: string, oathToken: string) {
  logger.debug(
    `curl -X GET '${url}' -H 'Accept: application/json, text/plain, */*'   -H 'Authorization: Bearer ${oathToken}'`
  );
  try {
    const response = await axios.get(url, {
      headers: { Authorization: `Bearer ${oathToken}` }
    });
    return response.data;
  } catch (error) {
    logger.error(`Error invoking API: ${error}`);
    throw error;
  }
}

async function getButtonActivity(qResponse: AmazonQResponse) {
  const cardButtons: CardAction[] = [];
  // view sources
  if (!isEmpty(qResponse.sourceAttributions)) {
    cardButtons.push({
      type: 'invoke',
      title: 'View sources',
      value: {
        type: 'task/fetch',
        action: 'ViewSources',
        systemMessageId: qResponse.systemMessageId
      }
    });
  }
  // feedback buttons
  [
    { action: 'ThumbsUp', label: '👍' },
    { action: 'ThumbsDown', label: '👎' }
  ].map((feedback) => {
    cardButtons.push({
      type: 'invoke',
      title: feedback.label,
      value: {
        type: ActionTypes.ImBack,
        action: feedback.action,
        systemMessageId: qResponse.systemMessageId
      }
    });
  });
  const html = await marked.parse(qResponse.systemMessage);
  const card = CardFactory.heroCard('', html, undefined, cardButtons);
  return MessageFactory.attachment(card);
}

function sourcesMarkdown(sources: SourceAttribution[]) {
  const md = [];
  for (let i = 0; i < sources.length; i++) {
    const source = sources[i];
    if (!isEmpty(source.title)) {
      if (!isEmpty(source.url)) {
        md.push(`**${i + 1}). [${source.title.trim()}](${source.url})**`);
      } else {
        md.push(`**${i + 1}). ${source.title.trim()}**`);
      }
    }
    if (!isEmpty(source.snippet)) {
      md.push(source.snippet.trim().replaceAll('\n\n', '\n'));
    }
  }
  return md.join('\n\n---\n') || 'No Sources found';
}

async function getSourceAttributions(sources: SourceAttribution[]) {
  const card = CardFactory.adaptiveCard({
    type: 'AdaptiveCard',
    version: '1.3',
    body: [
      {
        type: 'TextBlock',
        text: sourcesMarkdown(sources),
        wrap: true
      }
    ]
  });
  logger.debug(`Sources Card: ${JSON.stringify(card)}`);
  const taskModuleResponse: TaskModuleResponse = {
    task: {
      type: 'continue',
      value: {
        title: 'Sources',
        height: 700,
        width: 500,
        card: card
      }
    }
  };
  return taskModuleResponse;
}

export class QTeamsBot extends ActivityHandler {
  constructor() {
    super();
    // See https://aka.ms/about-bot-activity-message to learn more about activity types.
    this.onMessage(async (context) => {
      const tenantId = context.activity.conversation.tenantId;
      if (isEmpty(oathToken) && !isEmpty(tenantId)) {
        oathToken = await getOathToken(tenantId);
      }
      const env = getEnv(process.env);
      logger.debug(`Message handler`);
      const activity = context.activity;
      const type = context.activity.conversation.conversationType;

      const message = context.activity.text;
      logger.info(`Message received: ${message}`);
      let qUserMessage = message;
      const qAttachments: QAttachment[] = [];

      env.AMAZON_Q_USER_ID = await setAmazonQUserId(context, env.AMAZON_Q_USER_ID);

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
          await context.sendActivity(MessageFactory.text(replyText));
          return;
        }

        // it's not a reset command, so process the message
        // get any cached context metadata
        const channelMetadata = await getChannelMetadata(channelKey, env);
        logger.debug(
          `Cached channel metadata: channelKey: ${channelKey}, metadata: ${JSON.stringify(
            channelMetadata
          )} `
        );
        qContext = {
          conversationId: channelMetadata?.conversationId,
          parentMessageId: channelMetadata?.systemMessageId
        };
        // get any attachments from this message
        if (!isEmpty(activity?.attachments)) {
          qAttachments.push(...(await attachFiles(activity.attachments, oathToken)));
        }
      }
      // Team channel message (@mention)
      else if (type === 'channel') {
        // we don't cache context metadata for For @mentions in a channel. Instead we invoke AmazonQ with
        // a new conversation context, which embeds the thread message history in the input message.
        // The thread history includes any participant messages not previously seen by the bot.
        const [threadMessages, threadAttachments] = await retrieveThreadHistory(context, oathToken);
        logger.debug(`Thread messages: ${JSON.stringify(threadMessages)}`);
        // ignore last message.. it duplicates current message
        threadMessages.pop();
        if (!isEmpty(threadMessages)) {
          qUserMessage = `Given the following conversation thread history in JSON:\n${JSON.stringify(
            threadMessages
          )}\n${message}`;
        }
        if (!isEmpty(threadAttachments)) {
          threadAttachments.reverse(); // newest first
          qAttachments.push(...threadAttachments);
        }
      } else {
        logger.error(`Unsupported conversation type: ${type}`);
        return;
      }

      // Send Processing indicator
      logger.debug('Sending Processing indicator');
      const messageResponse = await context.sendActivity(MessageFactory.text(PROCESSING_MSG));

      // call ChatSync API
      const qResponse = await qChatSync(env, qUserMessage, qAttachments, qContext);
      if (qResponse instanceof Error) {
        const replyText = `${ERROR_MSG} : *${qResponse.message}*`;
        await context.sendActivity(MessageFactory.text(replyText));
        return;
      }

      // return Amazon Q response to user by updating the first ('Processing...') message
      const updatedMessage = await getButtonActivity(qResponse);
      if (!isEmpty(messageResponse)) {
        updatedMessage.id = messageResponse.id;
      }
      logger.debug(
        `Replace 'Processing...' message with Amazon Q Response: ${JSON.stringify(updatedMessage)}`
      );
      await context.updateActivity(updatedMessage);

      // save metadata from sucessful response
      if (type === 'personal') {
        logger.debug(`Saving channel metadata for '${channelKey}'`);
        await saveChannelMetadata(
          channelKey,
          qResponse.conversationId,
          qResponse.systemMessageId,
          env
        );
      }
      logger.debug(`Saving message metadata for '${qResponse.systemMessageId}'`);
      await saveMessageMetadata(qResponse, env);
      return;
    });
  }
  protected async onInvokeActivity(context: TurnContext): Promise<InvokeResponse> {
    logger.debug(`InvokeActivity handler`);
    const data = context.activity.value.data || context.activity.value;
    const action = data.action;
    const systemMessageId = data.systemMessageId;
    logger.debug(`Action: ${action}, systemMessageId: ${systemMessageId}`);
    const env = getEnv(process.env);
    const qResponse = (await getMessageMetadata(systemMessageId, env)) as AmazonQResponse;
    if (action === 'ViewSources') {
      const sources = qResponse?.sourceAttributions || [];
      const response = await getSourceAttributions(sources);
      logger.debug(`Response: ${JSON.stringify(response)}`);
      return { status: 200, body: response };
    }
    if (action === 'ThumbsUp' || action === 'ThumbsDown') {
      env.AMAZON_Q_USER_ID = await setAmazonQUserId(context, env.AMAZON_Q_USER_ID);
      await qPutFeedbackRequest(
        env,
        {
          conversationId: qResponse.conversationId,
          messageId: qResponse.systemMessageId
        },
        action === 'ThumbsUp' ? 'USEFUL' : 'NOT_USEFUL',
        action === 'ThumbsUp' ? 'HELPFUL' : 'NOT_HELPFUL'
      );
      // Diabled message, since Teams already displays "Your response was sent to the app"
      //await context.sendActivity(MessageFactory.text(FEEDBACK_ACK_MSG));
      return {
        status: 200,
        body: {
          task: {
            type: 'continue'
          }
        }
      };
    }
    logger.error(`Action '${action}' not implemented`);
    return { status: 200, body: `Error - Action '${action}' not implemented` };
  }
}
