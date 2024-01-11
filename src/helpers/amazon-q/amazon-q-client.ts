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

const logger = makeLogger('amazon-q-client');

let amazonQClient: QBusinessClient | null = null;
export const getClient = (env: Env) => {
  if (amazonQClient === null) {
    logger.debug(`Initiating AmazonQ client with region ${env.AMAZON_Q_REGION}`);
    amazonQClient = new QBusinessClient({
      region: env.AMAZON_Q_REGION
    });
  }
  return amazonQClient;
};

export const qChatSync = async (
  env: Env,
  message: string,
  attachments: AttachmentInput[],
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
      userId: env.AMAZON_Q_USER_ID,
      clientToken: uuid(),
      userMessage: message,
      ...(attachments.length > 0 && { attachments }),
      ...context
    };
    logger.debug(`AmazonQ chatSync input: ${JSON.stringify(input)}`);
    const response = await getClient(env).send(new ChatSyncCommand(input));
    logger.debug(`AmazonQ chatSync response: ${JSON.stringify(response)}`);
    return response;
  } catch (error) {
    logger.error(`Caught Exception: ${JSON.stringify(error)}`);
    if (error instanceof Error) {
      return new Error(error.message);
    } else {
      return new Error(`${JSON.stringify(error)}`);
    }
  }
};

export const qPutFeedbackRequest = async (
  env: Env,
  context: {
    conversationId: string;
    messageId: string;
  },
  usefulness: MessageUsefulness,
  reason: MessageUsefulnessReason
): Promise<PutFeedbackCommandOutput> => {
  const input: PutFeedbackCommandInput = {
    applicationId: env.AMAZON_Q_APP_ID,
    userId: env.AMAZON_Q_USER_ID,
    ...context,
    messageUsefulness: {
      usefulness: usefulness,
      reason: reason,
      submittedAt: new Date()
    }
  };

  logger.debug(`putFeedbackRequest input ${JSON.stringify(input)}`);
  const response = await getClient(env).send(new PutFeedbackCommand(input));
  logger.debug(`putFeedbackRequest output ${JSON.stringify(response)}`);

  return response;
};
