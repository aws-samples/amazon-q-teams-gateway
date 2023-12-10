#!/usr/bin/env node
import 'source-map-support/register';
import * as cdk from 'aws-cdk-lib';
import { MyAmazonQTeamsBotStack } from '../lib/my-amazon-q-teams-bot-stack';
import { readFileSync } from 'fs';

export interface StackEnvironment {
  StackName: string;
  AmazonQAppId: string;
  AmazonQUserId: string;
  AmazonQRegion: string;
  AmazonQEndpoint?: string;
  ContextDaysToLive: string;
}

const app = new cdk.App();
const inputEnvFile = app.node.tryGetContext('environment');
if (inputEnvFile === undefined) {
  throw new Error('An input environment file is required');
}

const environment = JSON.parse(readFileSync(inputEnvFile).toString()) as StackEnvironment;
if (environment.StackName === undefined) {
  throw new Error('StackName is required');
}
if (environment.AmazonQAppId === undefined) {
  throw new Error('AmazonQAppId is required');
}
if (environment.AmazonQRegion === undefined) {
  throw new Error('AmazonQRegion is required');
}
if (environment.AmazonQUserId === undefined) {
  throw new Error('AmazonQUserId is required');
}
if (environment.ContextDaysToLive === undefined) {
  throw new Error('ContextDaysToLive is required');
}

new MyAmazonQTeamsBotStack(
  app,
  'MyAmazonQTeamsBotStack',
  {
    stackName: environment.StackName
  },
  environment
);
