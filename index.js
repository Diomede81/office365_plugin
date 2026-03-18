/**
 * @openclaw/office365
 * Office 365 Channel Plugin for OpenClaw
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
  register(api) {
    console.log('[Office365] Plugin registering...');
    
    // Access plugin-specific config
    const config = api.config?.plugins?.entries?.office365?.config;
    
    if (config) {
      console.log('[Office365] Config loaded successfully');
      console.log('[Office365] - clientId:', config.clientId?.substring(0, 8) + '...');
      console.log('[Office365] - tokenFile:', config.tokenFile);
      console.log('[Office365] - webhookUrl:', config.webhookUrl);
    } else {
      console.log('[Office365] No config found');
    }
    
    // Test webhook - always register
    api.registerHttpRoute({
      path: '/webhooks/office365/test',
      auth: 'plugin',
      match: 'exact',
      handler: async (req, res) => {
        res.statusCode = 200;
        res.end(JSON.stringify({
          status: 'ok',
          plugin: 'office365',
          configLoaded: config ? 'yes' : 'no'
        }));
        return true;
      }
    });
    
    // Set up Teams webhook if config is available
    if (config && config.webhookUrl && config.tokenFile) {
      console.log('[Office365] Setting up Teams webhook...');
      
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
          console.log(`[Office365] Received ${notifications.length} notification(s)`);
          
          // TODO: Process notifications
          
          res.statusCode = 202;
          res.end();
          return true;
        }
      });
      
      console.log('[Office365] ✅ Teams webhook ready at /webhooks/office365/teams');
    } else {
      console.log('[Office365] ⚠️  No valid config - Teams webhook disabled');
    }
    
    console.log('[Office365] Plugin registered successfully');
  }
};
