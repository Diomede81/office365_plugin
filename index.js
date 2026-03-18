/**
 * @openclaw/office365
 * Office 365 Channel Plugin for OpenClaw
 * Minimal working version - step by step
 */

export default {
  id: 'office365',
  name: 'Office 365',
  
  /**
   * Plugin registration - minimal version that won't crash
   */
  register(api) {
    console.log('[Office365] Plugin registering - minimal version');
    
    // Just register a simple HTTP route to prove it works
    api.registerHttpRoute({
      path: '/webhooks/office365/test',
      auth: 'plugin',
      match: 'exact',
      handler: async (req, res) => {
        console.log('[Office365] Test webhook received');
        res.statusCode = 200;
        res.end('Office365 plugin is working');
        return true;
      }
    });
    
    console.log('[Office365] Plugin registered successfully');
  }
};
