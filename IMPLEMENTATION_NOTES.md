# Implementation Notes

## Why a Plugin Instead of Hooks?

### Problem with Hooks
OpenClaw's `/hooks/agent` endpoint creates **temporary one-off sessions** regardless of the `sessionKey` parameter. Each Teams message was creating a new session ID, meaning:

- ❌ No conversation history maintained
- ❌ Context lost between messages  
- ❌ Each message started fresh
- ❌ sessionKey parameter accepted but ignored for persistence

Evidence from testing:
- Session 70393abc (15:31): First 2 messages
- Session 4b31fa86 (15:33): Third message  
- Session eb6770f7 (15:39): Fourth message
- Same sessionKey sent, different sessions created

### Why Plugins Work

Channel plugins like WhatsApp create **proper persistent "other" sessions** that:

- ✅ Maintain conversation history
- ✅ Carry context across messages
- ✅ Use consistent session keys
- ✅ Support session compression
- ✅ Show in `sessions_list`

### Plugin Architecture

```javascript
class Office365Plugin {
  // Called by OpenClaw on initialization
  async init() {
    // Set up Graph client
    // Create webhook subscriptions
  }

  // Handle incoming webhooks from Microsoft
  async handleWebhook(req, res) {
    // Process notifications
    // Filter messages  
    // Deliver to OpenClaw with sessionKey
  }

  // Deliver message to OpenClaw
  async context.deliver({
    channel: 'teams',
    sessionKey: 'agent:max:teams:direct:chatId',
    from: { id, name },
    message: { id, text, timestamp },
    chat: { id, type, name }
  })

  // Send message
  async send(params) {
    // Post to Graph API
  }
}
```

## Session Key Format

Matches WhatsApp convention:

```
agent:{agentId}:teams:{chatType}:{chatId}
```

Examples:
- `agent:max:teams:direct:19:15f6df21...` (1:1 with Luca)
- `agent:max:teams:group:19:a26774df...` (NHS Assessments group)
- `agent:sophia:teams:direct:19:22e69c79...` (1:1 with Sophia)

## Integration Flow

```
Microsoft Teams Message
  ↓
Microsoft Graph API
  ↓  
Webhook to https://your-domain.com/webhooks/office365/teams
  ↓
Plugin handleWebhook()
  ↓
Message processing & filtering
  ↓
context.deliver() with sessionKey
  ↓
OpenClaw creates/reuses persistent session
  ↓
Agent processes with full conversation history
  ↓
Agent replies via message tool
  ↓
Plugin send() → Graph API → Teams
```

## Migration from Middleware

The existing middleware (`microsoft-middleware`) can be deprecated in favor of this plugin. Benefits:

| Feature | Middleware (Hooks) | Plugin (Channel) |
|---------|-------------------|------------------|
| Session persistence | ❌ Creates new sessions | ✅ Persistent sessions |
| Conversation context | ❌ Lost between messages | ✅ Full history |
| sessions_list visibility | ❌ Not shown | ✅ Listed as "other" |
| Token management | ✅ Working | ✅ Working |
| Message filtering | ✅ Working | ✅ Working |
| Subscriptions | ✅ Working | ✅ Working |
| Attachments | ✅ Send/receive | ✅ (To be implemented) |

## Next Steps

1. **Create GitHub repo**: `office365_plugin`
2. **Publish to npm**: `@openclaw/office365`
3. **Test installation**: `openclaw plugins install @openclaw/office365`
4. **Verify sessions persist**: Check with `sessions_list`
5. **Deprecate middleware**: Once plugin is stable

## Configuration Example

```json
{
  "plugins": {
    "entries": {
      "office365": {
        "enabled": true,
        "clientId": "79b3f60a-ddfe-4029-8af4-1c95a37c6aa7",
        "tenantId": "982780f8-0424-4e57-9cc0-bee3d6acc797",
        "tokenFile": "/home/lucalicata/clawd/max-microsoft-tokens.json",
        "webhookUrl": "https://microsoft.acuity.expert/webhooks/office365/teams"
      }
    }
  }
}
```

## Testing Plan

1. Install plugin
2. Send Teams message from Luca
3. Check `sessions_list` for Teams session
4. Send second message from Luca  
5. Verify same session ID used
6. Check agent remembers context
7. Test group chat with @mentions
8. Test attachments (future)

---

**Created**: 2026-03-18  
**Author**: Max (with Luca's guidance)  
**Status**: Initial implementation ready for testing
