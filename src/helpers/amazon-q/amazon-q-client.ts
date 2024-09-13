import { Env } from '@src/utils';
import { makeLogger } from '@src/logging';
import { v4 as uuid } from 'uuid';
import {
  QBusinessClient,
  ChatSyncCommand,
  MessageUsefulnessReason,
  MessageUsefulness,
  ChatSyncCommandOutput,
  AttachmentInput,
  PutFeedbackCommand,
  PutFeedbackCommandInput,
  PutFeedbackCommandOutput
} from '@aws-sdk/client-qbusiness';
import { Credentials } from 'aws-sdk';
import { ExpiredTokenException } from '@aws-sdk/client-sts';

const logger = makeLogger('amazon-q-client');

const amazonQClientByTeamUserId: { [key: string]: QBusinessClient } = {};

export const getClient = (
  env: Env,
  teamsUserId: string,
  iamSessionCreds: Credentials
) => {
  logger.debug(`Initiating AmazonQ client with region ${env.AMAZON_Q_REGION}`);
  if (amazonQClientByTeamUserId[teamsUserId]) {
    return amazonQClientByTeamUserId[teamsUserId];
  }

  const newClient = new QBusinessClient({
    credentials: iamSessionCreds,
    region: env.AMAZON_Q_REGION
  });

  amazonQClientByTeamUserId[teamsUserId] = newClient;

  return newClient;
};

export const qChatSync = async (
  env: Env,
  teamsUserId: string,
  message: string,
  attachments: AttachmentInput[],
  iamSessionCreds: Credentials,
  context?: {
    conversationId: string;
    parentMessageId: string;
  }
): Promise<ChatSyncCommandOutput | Error> => {
  try {
    // removed message and attachment limit checks, in favor
    // of letting it fail and report expception to user if limits
    // are exceeded.
    const input = {
      applicationId: env.AMAZON_Q_APP_ID,
      clientToken: uuid(),
      userMessage: message,
      ...(attachments.length > 0 && { attachments }),
      ...context
    };
    const response = await getClient(env,teamsUserId,iamSessionCreds).send(new ChatSyncCommand(input));
    logger.debug(`AmazonQ chatSync response: ${JSON.stringify(response)}`);
    return response;
  } catch (error) {
    logger.error(`Caught Exception: ${JSON.stringify(error)}`);
    if (error instanceof Error) {
      logger.debug(error.stack);
      if (error instanceof ExpiredTokenException) {
        logger.error(`Token expired: ${error.message}`);
      }
      return new Error(error.message);
    } else {
      return new Error(`${JSON.stringify(error)}`);
    }
  }
};

export const qPutFeedbackRequest = async (
  env: Env,
  teamsUserId: string,
  iamSessionCreds: Credentials,
  context: {
    conversationId: string;
    messageId: string;
  },
  usefulness: MessageUsefulness,
  reason: MessageUsefulnessReason
): Promise<PutFeedbackCommandOutput> => {
  const input: PutFeedbackCommandInput = {
    applicationId: env.AMAZON_Q_APP_ID,
    ...context,
    messageUsefulness: {
      usefulness: usefulness,
      reason: reason,
      submittedAt: new Date()
    }
  };

  logger.debug(`putFeedbackRequest input ${JSON.stringify(input)}`);
  const response = await getClient(env,teamsUserId,iamSessionCreds).send(new PutFeedbackCommand(input));
  logger.debug(`putFeedbackRequest output ${JSON.stringify(response)}`);

  return response;
};
