# Office 365 Plugin for OpenClaw

Official Microsoft Office 365 integration plugin for OpenClaw, providing persistent Teams chat sessions with full conversation context.

## Features

✅ **Persistent Teams Sessions** - Full conversation history maintained across messages  
✅ **Group Chat Support** - Separate sessions for each Teams chat  
✅ **Auto Token Refresh** - Automatic token management with 5-minute buffer  
✅ **Webhook Subscriptions** - Real-time message delivery via Microsoft Graph  
✅ **Message Filtering** - Skip own messages, duplicates, and old messages  
✅ **Attachment Support** - Send and receive files (planned)  

## Installation

```bash
npm install @openclaw/office365
```

Or install from GitHub:

```bash
openclaw plugins install github:Diomede81/office365_plugin
```

## Configuration

Add to your `openclaw.json`:

```json
{
  "plugins": {
    "entries": {
      "office365": {
        "enabled": true,
        "clientId": "your-azure-app-client-id",
        "tenantId": "your-azure-tenant-id",
        "tokenFile": "/path/to/tokens.json",
        "webhookUrl": "https://your-domain.com/webhooks/office365/teams"
      }
    }
  }
}
```

## Azure App Setup

1. **Register App** in Azure Portal → App Registrations
2. **API Permissions:**
   - `Chat.ReadWrite`
   - `ChatMessage.Read`
   - `ChatMessage.Send`
   - `User.Read`
3. **Authentication:**
   - Redirect URI: `http://localhost` (for device code flow)
   - Enable public client flows
4. **Copy:**
   - Application (client) ID
   - Directory (tenant) ID

## Token Generation

Use device code flow to generate initial tokens:

```javascript
const fetch = require('isomorphic-fetch');
const fs = require('fs');

const CLIENT_ID = 'your-client-id';
const TENANT_ID = 'your-tenant-id';
const SCOPES = 'https://graph.microsoft.com/.default offline_access';

async function getTokens() {
  // 1. Request device code
  const deviceResp = await fetch(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/devicecode`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: CLIENT_ID,
        scope: SCOPES
      })
    }
  );
  
  const device = await deviceResp.json();
  console.log(device.message); // Show user instructions
  
  // 2. Poll for token
  const interval = device.interval * 1000;
  const expiresAt = Date.now() + (device.expires_in * 1000);
  
  while (Date.now() < expiresAt) {
    await new Promise(resolve => setTimeout(resolve, interval));
    
    const tokenResp = await fetch(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
          client_id: CLIENT_ID,
          device_code: device.device_code
        })
      }
    );
    
    const tokens = await tokenResp.json();
    
    if (tokens.access_token) {
      tokens.obtained_at = Date.now();
      fs.writeFileSync('tokens.json', JSON.stringify(tokens, null, 2));
      console.log('✅ Tokens saved to tokens.json');
      return;
    }
    
    if (tokens.error !== 'authorization_pending') {
      throw new Error(tokens.error_description);
    }
  }
  
  throw new Error('Device code expired');
}

getTokens().catch(console.error);
```

## Session Format

Teams sessions use the format:

```
agent:{agentId}:teams:{chatType}:{chatId}
```

Examples:
- `agent:max:teams:direct:19:15f6df21...` (1:1 chat)
- `agent:max:teams:group:19:a26774df...` (group chat)

This matches OpenClaw's WhatsApp session format for consistency.

## Sending Messages

Via OpenClaw `message` tool:

```javascript
await message({
  action: 'send',
  channel: 'teams',
  to: 'chat-id',
  message: 'Hello from OpenClaw!'
});
```

Via plugin API:

```javascript
await office365.send({
  to: 'chat-id',
  text: 'Plain text message',
  html: '<p>HTML message</p>'
});
```

## Webhook Endpoint

The plugin registers a webhook at:

```
POST /webhooks/office365/teams
```

Ensure this endpoint is:
1. Publicly accessible (use Cloudflare Tunnel, ngrok, etc.)
2. Uses HTTPS (Microsoft requires it)
3. Configured in plugin config as `webhookUrl`

## Architecture

```
Microsoft Graph API
       ↓
  Webhook Subscription (1-hour TTL)
       ↓
  Plugin Webhook Handler
       ↓
  Message Processing & Filtering
       ↓
  OpenClaw Session Delivery
       ↓
  Persistent Teams Session
```

## Session Persistence

Unlike hooks, this plugin creates proper persistent sessions:

- ✅ Each chat has its own session
- ✅ Conversation history maintained
- ✅ Context carried across messages
- ✅ Automatic session compression
- ✅ Same behavior as WhatsApp channel

## Subscription Management

Subscriptions auto-renew every 5 minutes with 1-hour expiry. On shutdown, all subscriptions are cleaned up.

## Troubleshooting

**No messages arriving:**
1. Check webhook URL is publicly accessible
2. Verify subscription is active: `GET /subscriptions`
3. Check token hasn't expired
4. Ensure Azure app has correct permissions

**Token refresh failing:**
1. Verify `refresh_token` in tokens.json
2. Check refresh token hasn't expired (90 days inactivity)
3. Re-run token generation script

**Session not persisting:**
1. Verify plugin is loaded (not hook-based integration)
2. Check session key format in logs
3. Ensure OpenClaw version >= 2026.3.0

## Development

```bash
git clone https://github.com/Diomede81/office365_plugin.git
cd office365_plugin
npm install
npm link
```

Then in OpenClaw config:

```json
{
  "plugins": {
    "entries": {
      "office365": {
        "enabled": true,
        "source": "link"
      }
    }
  }
}
```

## License

MIT

## Author

Luca Licata

## Repository

https://github.com/Diomede81/office365_plugin
