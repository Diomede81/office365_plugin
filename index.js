/**
 * @openclaw/office365
 * Office 365 Channel Plugin for OpenClaw
 * 
 * Provides persistent Teams chat sessions with full conversation context
 */

import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import fs from 'fs';

export default {
  id: 'office365',
  name: 'Office 365',
  
  /**
   * Plugin registration
   */
  async register(api) {
    const config = api.config;
    
    console.log('[Office365] Registering plugin...');
    
    // Initialize Graph client
    const graphClient = await initializeGraphClient(config);
    
    // Register webhook HTTP route
    api.registerHttpRoute({
      path: '/webhooks/office365/teams',
      auth: 'plugin',
      match: 'exact',
      handler: async (req, res) => {
        // Validation request
        if (req.query?.validationToken) {
          res.statusCode = 200;
          res.end(req.query.validationToken);
          return true;
        }

        const notifications = req.body?.value || [];
        
        for (const notification of notifications) {
          try {
            await processNotification(notification, graphClient, api);
          } catch (error) {
            console.error('[Office365] Error processing notification:', error);
          }
        }

        res.statusCode = 202;
        res.end();
        return true;
      }
    });
    
    // Set up webhook subscriptions
    await setupSubscriptions(graphClient, config);
    
    console.log('[Office365] Plugin registered successfully');
  }
};

/**
 * Initialize Microsoft Graph client
 */
async function initializeGraphClient(config) {
  const tokenFile = config.tokenFile;
  
  if (!fs.existsSync(tokenFile)) {
    throw new Error(`Token file not found: ${tokenFile}`);
  }

  const authProvider = {
    getAccessToken: async () => {
      const tokens = JSON.parse(fs.readFileSync(tokenFile, 'utf8'));
      const expiresAt = tokens.obtained_at + (tokens.expires_in * 1000);
      
      // Refresh if expired or expiring soon (5 min buffer)
      if (Date.now() > expiresAt - 300000) {
        console.log('[Office365] Refreshing access token...');
        
        const response = await fetch(
          `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`,
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
              client_id: config.clientId,
              refresh_token: tokens.refresh_token,
              grant_type: 'refresh_token',
              scope: 'https://graph.microsoft.com/.default'
            })
          }
        );
        
        const newTokens = await response.json();
        if (newTokens.error) {
          throw new Error(`Token refresh failed: ${newTokens.error_description}`);
        }
        
        newTokens.obtained_at = Date.now();
        fs.writeFileSync(tokenFile, JSON.stringify(newTokens, null, 2));
        return newTokens.access_token;
      }
      
      return tokens.access_token;
    }
  };

  const client = Client.initWithMiddleware({ authProvider });
  console.log('[Office365] Graph client initialized');
  return client;
}

/**
 * Set up webhook subscriptions
 */
async function setupSubscriptions(client, config) {
  const webhookUrl = config.webhookUrl;
  
  try {
    const subscription = await client
      .api('/subscriptions')
      .post({
        changeType: 'created',
        notificationUrl: webhookUrl,
        resource: '/me/chats/getAllMessages',
        expirationDateTime: new Date(Date.now() + 60 * 60 * 1000).toISOString(),
        clientState: 'openclaw-teams'
      });
    
    console.log(`[Office365] Teams subscription created: ${subscription.id}`);
  } catch (error) {
    console.error('[Office365] Failed to create subscription:', error.message);
  }
}

/**
 * Process a single notification
 */
async function processNotification(notification, client, api) {
  const resourceUrl = notification.resource;
  
  // Fetch the actual message
  const message = await client.api(resourceUrl).get();
  
  // Extract details
  const chatId = message.chatId;
  const messageId = message.id;
  const content = message.body?.content?.replace(/<[^>]*>/g, '').trim() || '';
  const from = message.from?.user?.displayName || 'Unknown';
  const fromId = message.from?.user?.id;
  
  // Skip our own messages
  const me = await client.api('/me').get();
  if (fromId === me.id) {
    console.log('[Office365] Skipping own message');
    return;
  }

  // Determine chat type
  const chat = await client.api(`/me/chats/${chatId}`).get();
  const isGroup = chat.chatType === 'group';
  
  // Create persistent session key (matches WhatsApp format)
  const agentId = 'max'; // TODO: Get from api
  const sessionKey = `agent:${agentId}:teams:${isGroup ? 'group' : 'direct'}:${chatId}`;
  
  console.log(`[Office365] Delivering message from ${from} with sessionKey: ${sessionKey}`);
  
  // TODO: Figure out correct api.deliver() call
  // For now just log it
  console.log(`[Office365] Would deliver: channel=teams, sessionKey=${sessionKey}, from=${from}, text=${content}`);
}
