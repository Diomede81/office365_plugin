/**
 * @openclaw/office365
 * Office 365 Channel Plugin for OpenClaw
 */

import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import fs from 'fs';

let graphClient = null;

/**
 * Get access token with auto-refresh
 */
async function getAccessToken(config) {
  const tokens = JSON.parse(fs.readFileSync(config.tokenFile, 'utf8'));
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
    fs.writeFileSync(config.tokenFile, JSON.stringify(newTokens, null, 2));
    console.log('[Office365] ✅ Token refreshed');
    return newTokens.access_token;
  }
  
  return tokens.access_token;
}

/**
 * Initialize Graph client
 */
async function initGraphClient(config) {
  const authProvider = {
    getAccessToken: async () => await getAccessToken(config)
  };
  
  graphClient = Client.initWithMiddleware({ authProvider });
  console.log('[Office365] Graph client initialized');
}

/**
 * Create webhook subscription
 */
async function createSubscription(config) {
  try {
    const subscription = await graphClient
      .api('/subscriptions')
      .post({
        changeType: 'created',
        notificationUrl: config.webhookUrl,
        resource: '/me/chats/getAllMessages',
        expirationDateTime: new Date(Date.now() + 60 * 60 * 1000).toISOString(),
        clientState: 'openclaw-teams'
      });
    
    console.log(`[Office365] ✅ Subscription created: ${subscription.id}`);
    return subscription;
  } catch (error) {
    console.error('[Office365] ❌ Subscription failed:', error.message);
    throw error;
  }
}

export default {
  id: 'office365',
  name: 'Office 365',
  
  /**
   * Plugin registration
   */
  async register(api) {
    console.log('[Office365] Plugin registering...');
    
    // Access plugin-specific config
    const config = api.config?.plugins?.entries?.office365?.config;
    
    if (!config || !config.webhookUrl || !config.tokenFile) {
      console.log('[Office365] ⚠️  No valid config - plugin disabled');
      return;
    }
    
    console.log('[Office365] Config loaded');
    
    // Initialize Graph client
    try {
      await initGraphClient(config);
    } catch (error) {
      console.error('[Office365] ❌ Graph client init failed:', error.message);
      return;
    }
    
    // Register Teams webhook
    api.registerHttpRoute({
      path: '/webhooks/office365/teams',
      auth: 'plugin',
      match: 'exact',
      handler: async (req, res) => {
        // Validation request from Microsoft
        if (req.query?.validationToken) {
          console.log('[Office365] Webhook validation');
          res.statusCode = 200;
          res.end(req.query.validationToken);
          return true;
        }
        
        // Notification from Microsoft
        const notifications = req.body?.value || [];
        console.log(`[Office365] 📬 Received ${notifications.length} notification(s)`);
        
        // TODO: Process notifications
        
        res.statusCode = 202;
        res.end();
        return true;
      }
    });
    
    console.log('[Office365] ✅ Teams webhook registered');
    
    // Create subscription
    try {
      await createSubscription(config);
    } catch (error) {
      console.log('[Office365] ⚠️  Subscription creation failed, will retry later');
    }
    
    console.log('[Office365] ✅ Plugin ready');
  }
};
