import { buildClient, CommitmentPolicy, KmsKeyringNode } from '@aws-crypto/client-node';

type EncryptionContext = {
  [key: string]: string;
};

const { encrypt, decrypt } = buildClient(CommitmentPolicy.REQUIRE_ENCRYPT_REQUIRE_DECRYPT);

// Encrypts the given data using the provided keyId and teamsUserId
export const encryptData = async (data: string, keyId: string, teamsUserId: string) => {
  if (!data || !keyId || !teamsUserId) {
    throw new Error('Invalid arguments');
  }
  const keyring = new KmsKeyringNode({ generatorKeyId: keyId });
  const contextData = buildContextData(teamsUserId);
  const encryptionContext = buildEncryptionContext(contextData);

  const { result } = await encrypt(keyring, data, { encryptionContext: encryptionContext });

  return result.toString('base64');
};

// Decrypts the given cipherText using the provided keyId and teamsUserId
export const decryptData = async (cipherText: string, keyId: string, teamsUserId: string) => {
  // decode base64
  const data = Buffer.from(cipherText, 'base64');
  const keyring = new KmsKeyringNode({ generatorKeyId: keyId });

  const { plaintext, messageHeader } = await decrypt(keyring, data);

  // validate encryption context
  const { encryptionContext } = messageHeader;
  if (encryptionContext.teamsUserId !== teamsUserId) {
    throw new Error('Invalid encryption context - teamsUserId mismatch');
  }

  return plaintext.toString();
};

// Builds an encryption context from the given contextData
const buildEncryptionContext = (contextData: Record<string, string>): EncryptionContext => {
  const encryptionContext: EncryptionContext = {};

  for (const key in contextData) {
    if (Object.prototype.hasOwnProperty.call(contextData, key)) {
      encryptionContext[key] = contextData[key];
    }
  }

  return encryptionContext;
};

// Builds context data for the given teamsUserId
const buildContextData = (teamsUserId: string): Record<string, string> => {
  return {
    teamsUserId: teamsUserId
  };
};
