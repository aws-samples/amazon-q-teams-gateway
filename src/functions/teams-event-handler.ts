import 'source-map-support/register';
import express from 'express';
import serverless from 'serverless-http';
import { Context } from 'aws-lambda';
import { isEmpty, getTeamsSecret } from '@src/utils';
import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  ConfigurationBotFrameworkAuthenticationOptions
} from 'botbuilder';
import { QTeamsBot, ERROR_MSG } from '../helpers/teams/q-teams-bot';

import { makeLogger } from '@src/logging';
const logger = makeLogger('teams-event-handler');

// We use serverless-http to emulate an http app server in Lambda
// By making 'apphandler' a global, we can avoid initializing it every handler function invocation
let apphandler: serverless.Handler | undefined;
const init_apphandler = async () => {
  // bot framework authenticates api requests, using our provisioned secrets
  const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
    (await getTeamsSecret()) as ConfigurationBotFrameworkAuthenticationOptions
  );
  const adapter = new CloudAdapter(botFrameworkAuthentication);
  adapter.onTurnError = async (context, error) => {
    const errorMessage = `${ERROR_MSG} *[onTurnError]: ${error}*`;
    logger.error(errorMessage);
    await context.sendActivity(errorMessage);
  };
  const app: express.Application = express();
  app.use(express.json());
  const qBot = new QTeamsBot();
  // handle POST requests (new messages)
  app.post('/', (req: express.Request, res: express.Response) => {
    logger.debug(
      `HTTP Request - URL: ${req.url}, Method: ${req.method}, Headers: ${JSON.stringify(
        req.headers
      )}`
    );
    logger.debug(`HTTP Request - Body: ${JSON.stringify(req.body)}`);
    adapter.process(req, res, async (context) => {
      logger.debug(`Bot Context: ${JSON.stringify(context)}`);
      await qBot.run(context);
    });
  });
  return serverless(app);
};

export const handler = async (
  event: {
    body: string;
    headers: { [key: string]: string | undefined };
  },
  context: Context
) => {
  logger.debug(`EVENT request received: ${JSON.stringify(event)}`);
  if (isEmpty(apphandler)) {
    logger.debug(`Initializing serverless apphandler`);
    apphandler = await init_apphandler();
  }
  const response = await apphandler(event, context);
  logger.debug(`EVENT response: ${JSON.stringify(response)}`);
  return response;
};
