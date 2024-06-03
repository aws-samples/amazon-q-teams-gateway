/* eslint @typescript-eslint/no-explicit-any: "off" */

import { SecretsManager } from 'aws-sdk';
import { makeLogger } from '@src/logging';
const logger = makeLogger('teams-event-handler');

export const isEmpty = <T>(value: T | undefined, checkAttributes = false): value is undefined => {
  if (value === undefined || value === null) {
    return true;
  }

  if (typeof value === 'number' || typeof value === 'boolean') {
    return false;
  }

  if (value instanceof Date) {
    return false;
  }

  if (Array.isArray(value)) {
    return value.length === 0 || value.every((item) => isEmpty(item));
  }

  if (value instanceof Object) {
    if (Object.keys(value).length === 0) {
      return true;
    }

    if (checkAttributes) {
      return Object.values(value).every((item) => isEmpty(item));
    }
  }

  return <any>value === '';
};

export const getOrThrowIfEmpty = <T>(value: T | undefined, name = 'element') => {
  if (isEmpty(value)) {
    throw new Error(`InvalidArgumentException: ${name} can't be empty`);
  }

  return value;
};

export const getEnv = (env: NodeJS.ProcessEnv) => ({
  AMAZON_Q_ENDPOINT: env.AMAZON_Q_ENDPOINT,
  AMAZON_Q_APP_ID: getOrThrowIfEmpty(env.AMAZON_Q_APP_ID),
  AMAZON_Q_REGION: getOrThrowIfEmpty(env.AMAZON_Q_REGION),
  CONTEXT_DAYS_TO_LIVE: getOrThrowIfEmpty(env.CONTEXT_DAYS_TO_LIVE),
  CACHE_TABLE_NAME: getOrThrowIfEmpty(env.CACHE_TABLE_NAME),
  MESSAGE_METADATA_TABLE_NAME: getOrThrowIfEmpty(env.MESSAGE_METADATA_TABLE_NAME),
  OIDC_STATE_TABLE_NAME: getOrThrowIfEmpty(env.OIDC_STATE_TABLE_NAME),
  IAM_SESSION_TABLE_NAME: getOrThrowIfEmpty(env.IAM_SESSION_CREDENTIALS_TABLE_NAME),
  OIDC_IDP_NAME: getOrThrowIfEmpty(env.OIDC_IDP_NAME),
  OIDC_ISSUER_URL: getOrThrowIfEmpty(env.OIDC_ISSUER_URL),
  OIDC_CLIENT_ID: getOrThrowIfEmpty(env.OIDC_CLIENT_ID),
  OIDC_CLIENT_SECRET_NAME: getOrThrowIfEmpty(env.OIDC_CLIENT_SECRET_NAME),
  OIDC_REDIRECT_URL: getOrThrowIfEmpty(env.OIDC_REDIRECT_URL),
  KMS_KEY_ARN: getOrThrowIfEmpty(env.KEY_ARN),
  Q_USER_API_ROLE_ARN: getOrThrowIfEmpty(env.Q_USER_API_ROLE_ARN),
  GATEWAY_IDC_APP_ARN: getOrThrowIfEmpty(env.GATEWAY_IDC_APP_ARN),
  MICROSOFT_APP_ID: getOrThrowIfEmpty(env.MICROSOFT_APP_ID),
});

export const getTeamsSecret = async () => {
  const secretName = getOrThrowIfEmpty(process.env.TEAMS_SECRET_NAME);
  logger.debug(`Getting secret value for SecretId ${secretName}`);
  const secretManagerClient = new SecretsManager();
  const secret = await secretManagerClient
    .getSecretValue({
      SecretId: secretName
    })
    .promise();
  if (secret.SecretString === undefined) {
    throw new Error('Missing SecretString');
  }
  return JSON.parse(secret.SecretString);
};

export type Env = ReturnType<typeof getEnv>;
