/**
 * @openclaw/office365
 * Office 365 Channel Plugin for OpenClaw
 * 
 * Provides persistent Teams chat sessions with full conversation context
 */

import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import fs from 'fs';
import path from 'path';

export default class Office365Plugin {
  constructor(context) {
    this.context = context;
    this.config = context.config;
    this.log = context.log;
    this.client = null;
    
    // Track active subscriptions
    this.subscriptions = new Map();
    
    this.log.info('[Office365] Plugin initialized');
  }

  /**
   * Plugin metadata
   */
  static get meta() {
    return {
      id: 'office365',
      name: 'Office 365',
      version: '1.0.0',
      description: 'Microsoft Teams, Outlook, and Calendar integration',
      channels: ['teams', 'outlook'],
      capabilities: {
        send: true,
        receive: true,
        attachments: true,
        reactions: false,
        presence: true
      }
    };
  }

  /**
   * Initialize the plugin
   */
  async init() {
    this.log.info('[Office365] Initializing plugin...');
    
    // Initialize Graph API client
    await this.initializeGraphClient();
    
    // Set up webhook subscriptions
    await this.setupSubscriptions();
    
    this.log.info('[Office365] Plugin ready');
  }

  /**
   * Initialize Microsoft Graph client
   */
  async initializeGraphClient() {
    const tokenFile = this.config.tokenFile;
    
    if (!fs.existsSync(tokenFile)) {
      throw new Error(`Token file not found: ${tokenFile}`);
    }

    const authProvider = {
      getAccessToken: async () => {
        return await this.getAccessToken();
      }
    };

    this.client = Client.initWithMiddleware({ authProvider });
    this.log.info('[Office365] Graph client initialized');
  }

  /**
   * Get access token with auto-refresh
   */
  async getAccessToken() {
    const tokens = JSON.parse(fs.readFileSync(this.config.tokenFile, 'utf8'));
    const expiresAt = tokens.obtained_at + (tokens.expires_in * 1000);
    
    // Refresh if expired or expiring soon (5 min buffer)
    if (Date.now() > expiresAt - 300000) {
      this.log.info('[Office365] Refreshing access token...');
      
      const response = await fetch(
        `https://login.microsoftonline.com/${this.config.tenantId}/oauth2/v2.0/token`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: new URLSearchParams({
            client_id: this.config.clientId,
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
      fs.writeFileSync(this.config.tokenFile, JSON.stringify(newTokens, null, 2));
      return newTokens.access_token;
    }
    
    return tokens.access_token;
  }

  /**
   * Set up webhook subscriptions for Teams messages
   */
  async setupSubscriptions() {
    const webhookUrl = this.config.webhookUrl || `${this.context.gateway.publicUrl}/webhooks/office365/teams`;
    
    try {
      // Subscribe to Teams chat messages
      const subscription = await this.client
        .api('/subscriptions')
        .post({
          changeType: 'created',
          notificationUrl: webhookUrl,
          resource: '/me/chats/getAllMessages',
          expirationDateTime: new Date(Date.now() + 60 * 60 * 1000).toISOString(), // 1 hour
          clientState: 'openclaw-teams'
        });
      
      this.subscriptions.set('teams', subscription.id);
      this.log.info(`[Office365] Teams subscription created: ${subscription.id}`);
    } catch (error) {
      this.log.error('[Office365] Failed to create subscription:', error.message);
    }
  }

  /**
   * Handle incoming webhook from Microsoft Graph
   */
  async handleWebhook(req, res) {
    // Validation request
    if (req.query.validationToken) {
      return res.status(200).send(req.query.validationToken);
    }

    const notifications = req.body.value || [];
    
    for (const notification of notifications) {
      try {
        await this.processNotification(notification);
      } catch (error) {
        this.log.error('[Office365] Error processing notification:', error);
      }
    }

    res.status(202).send();
  }

  /**
   * Process a single notification
   */
  async processNotification(notification) {
    const resourceUrl = notification.resource;
    
    // Fetch the actual message
    const message = await this.client.api(resourceUrl).get();
    
    // Extract chat and message details
    const chatId = message.chatId;
    const messageId = message.id;
    const content = message.body?.content?.replace(/<[^>]*>/g, '').trim() || '';
    const from = message.from?.user?.displayName || 'Unknown';
    
    // Skip our own messages
    const me = await this.client.api('/me').get();
    if (message.from?.user?.id === me.id) {
      this.log.debug('[Office365] Skipping own message');
      return;
    }

    // Determine chat type
    const chat = await this.client.api(`/me/chats/${chatId}`).get();
    const isGroup = chat.chatType === 'group';
    
    // Create persistent session key
    const sessionKey = `agent:${this.context.agentId}:teams:${isGroup ? 'group' : 'direct'}:${chatId}`;
    
    // Deliver to OpenClaw
    await this.context.deliver({
      channel: 'teams',
      sessionKey: sessionKey,
      from: {
        id: message.from?.user?.id,
        name: from
      },
      message: {
        id: messageId,
        text: content,
        timestamp: new Date(message.createdDateTime).getTime()
      },
      chat: {
        id: chatId,
        type: isGroup ? 'group' : 'direct',
        name: chat.topic || from
      }
    });
  }

  /**
   * Send a message
   */
  async send(params) {
    const { to, text, html } = params;
    
    const payload = {
      body: {
        contentType: html ? 'html' : 'text',
        content: html || text
      }
    };

    const response = await this.client
      .api(`/me/chats/${to}/messages`)
      .post(payload);

    return {
      id: response.id,
      timestamp: new Date(response.createdDateTime).getTime()
    };
  }

  /**
   * Clean up on shutdown
   */
  async shutdown() {
    // Delete subscriptions
    for (const [name, id] of this.subscriptions) {
      try {
        await this.client.api(`/subscriptions/${id}`).delete();
        this.log.info(`[Office365] Deleted subscription: ${name}`);
      } catch (error) {
        this.log.warn(`[Office365] Failed to delete subscription ${name}:`, error.message);
      }
    }
    
    this.log.info('[Office365] Plugin shut down');
  }
}
