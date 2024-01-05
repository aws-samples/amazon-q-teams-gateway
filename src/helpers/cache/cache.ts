import * as AWS from 'aws-sdk';
import { Env } from '@src/utils';
import { AmazonQResponse } from '@src/helpers/amazon-q/amazon-q-client';

export const client = new AWS.DynamoDB.DocumentClient({
  region: process.env.AWS_REGION,
  paramValidation: false, // Avoid extra latency
  convertResponseTypes: false // Avoid extra latency
});

export const deleteItem = async (args: AWS.DynamoDB.DocumentClient.DeleteItemInput) =>
  await client.delete(args).promise();

export const putItem = async (args: AWS.DynamoDB.DocumentClient.PutItemInput) =>
  await client.put(args).promise();

export const getItem = async (args: AWS.DynamoDB.DocumentClient.GetItemInput) =>
  await client.get(args).promise();

const expireAt = (env: Env) => {
  const contextTTL = Number(env.CONTEXT_DAYS_TO_LIVE) * 24 * 60 * 60 * 1000; // milliseconds
  return Math.floor((Date.now() + contextTTL) / 1000); // Unix time (seconds);
};

export const saveChannelMetadata = async (
  channel: string,
  conversationId: string,
  systemMessageId: string,
  env: Env
) => {
  await putItem({
    TableName: env.CACHE_TABLE_NAME,
    Item: {
      channel,
      conversationId,
      systemMessageId,
      latestTs: Date.now(),
      expireAt: expireAt(env)
    }
  });
};

export const getChannelMetadata = async (channel: string, env: Env) =>
  (
    await getItem({
      TableName: env.CACHE_TABLE_NAME,
      Key: {
        channel: channel
      }
    })
  ).Item;

export const deleteChannelMetadata = async (channel: string, env: Env) =>
  await deleteItem({
    TableName: env.CACHE_TABLE_NAME,
    Key: {
      channel
    }
  });

export const saveMessageMetadata = async (amazonQResponse: AmazonQResponse, env: Env) => {
  await putItem({
    TableName: env.MESSAGE_METADATA_TABLE_NAME,
    Item: {
      messageId: amazonQResponse.systemMessageId,
      conversationId: amazonQResponse.conversationId,
      systemMessage: amazonQResponse.systemMessage,
      sourceAttributions: amazonQResponse.sourceAttributions,
      systemMessageId: amazonQResponse.systemMessageId,
      userMessageId: amazonQResponse.userMessageId,
      ts: Date.now(),
      expireAt: expireAt(env)
    }
  });
};

export const getMessageMetadata = async (systemMessageId: string, env: Env) =>
  (
    await getItem({
      TableName: env.MESSAGE_METADATA_TABLE_NAME,
      Key: {
        messageId: systemMessageId
      }
    })
  ).Item;
