# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.2.0] - 2024-06-03
### Added
Add support for Q Business Apps integrated with IdC.
- The gateway registers the Amazon Q Business Microsoft Teams Gateway as an OpenID Connect (OIDC) app with Okta (or other OIDC compliant Identity Providers).
- This registration allows the gateway to invoke the Q Business ChatSync API on behalf of the end-user.
- The gateway provisions an OIDC callback handler for the Identity Provider (IdP) to return an authorization code after the end-user authenticates using the authorization grant flow.
- The callback handler exchanges the authorization code for IAM session credentials through a series of interactions with the IdP, IdC, and AWS Security Token Service (STS).
- The IAM session credentials, which are short-lived (15-minute duration), are encrypted and stored in a DynamoDB table along with the refresh token from the IdP.
- The IAM session credentials are then used to invoke the Q Business ChatSync and PutFeedback APIs.
- If the IAM session credentials expire, the refresh token from the IdP is used to obtain new IAM session credentials without requiring the end-user to sign in again.


## [0.1.1] - 2024-01-11
### Added
Initial release
- In DMs it responds to all messages
- In channels it responds only to @mentions, and always replies in thread
- Renders answers containing markdown - e.g. headings, lists, bold, italics, tables, etc. 
- Provides thumbs up / down buttons to track user sentiment and help improve performance over time
- Provides Source Attribution - see references to sources used by Amazon Q
- Aware of conversation context - it tracks the conversation and applies context
- Aware of multiple users - when it's tagged in a thread, it knows who said what, and when - so it can contribute in context and accurately summarize the thread when asked.  
- Process up to 5 attached files for document question answering, summaries, etc.
- Reset and start new conversation in DM channel by using `/new_conversation`

[Unreleased]: https://github.com/aws-samples/amazon-q-teams-gateway/compare/v0.1.1...develop
[0.1.1]: https://github.com/aws-samples/amazon-q-teams-gateway/releases/tag/v0.1.1
[0.1.0]: https://github.com/aws-samples/amazon-q-teams-gateway/releases/tag/v0.1.0
