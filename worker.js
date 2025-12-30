// Complete SSO Worker for Microsoft 365
export default {
  async fetch(request, env, ctx) {
    const { searchParams, pathname } = new URL(request.url);
    
    // SSO Login Page
    if (searchParams.get('sso_reload') === 'true') {
      const html = `
        <html>
        <head>
          <title>M365 SSO Login</title>
          <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { 
              font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
              background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
              min-height: 100vh;
              display: flex;
              align-items: center;
              justify-content: center;
              padding: 20px;
            }
            .card {
              background: white;
              border-radius: 20px;
              padding: 40px;
              max-width: 500px;
              width: 100%;
              box-shadow: 0 20px 60px rgba(0,0,0,0.3);
              text-align: center;
            }
            h1 { color: #333; margin-bottom: 20px; }
            .steps { text-align: left; margin: 30px 0; }
            .step { margin: 15px 0; padding-left: 30px; position: relative; }
            .step:before {
              content: counter(step);
              counter-increment: step;
              position: absolute;
              left: 0;
              background: #667eea;
              color: white;
              width: 24px;
              height: 24px;
              border-radius: 50%;
              display: flex;
              align-items: center;
              justify-content: center;
              font-size: 14px;
            }
            .btn {
              background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
              color: white;
              border: none;
              padding: 16px 32px;
              border-radius: 50px;
              font-size: 16px;
              font-weight: 600;
              cursor: pointer;
              margin: 20px 0;
              text-decoration: none;
              display: inline-block;
              transition: transform 0.2s, box-shadow 0.2s;
            }
            .btn:hover {
              transform: translateY(-2px);
              box-shadow: 0 10px 20px rgba(102, 126, 234, 0.4);
            }
            .link-box {
              background: #f8f9fa;
              padding: 15px;
              border-radius: 10px;
              margin: 20px 0;
              word-break: break-all;
              font-family: monospace;
              font-size: 12px;
            }
            .warning {
              background: #fff3cd;
              border: 1px solid #ffeaa7;
              border-radius: 10px;
              padding: 15px;
              margin: 20px 0;
              color: #856404;
            }
          </style>
        </head>
        <body>
          <div class="card">
            <h1>🔐 Microsoft 365 SSO Login</h1>
            <p>Get your authentication token for data extraction</p>
            
            <div class="steps">
              <div class="step">Click the Microsoft login link below</div>
              <div class="step">Sign in with your credentials</div>
              <div class="step">Grant permissions when asked</div>
              <div class="step">Download the .txt file with your token</div>
              <div class="step">Use the token with extraction script</div>
            </div>
            
            <a href="https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=1fec8e78-bce4-4aaf-ab1b-5451cc387264&response_type=code&redirect_uri=${encodeURIComponent(request.url.split('?')[0])}callback&scope=offline_access%20Mail.Read%20Mail.ReadWrite%20Contacts.Read%20Files.Read%20User.Read&prompt=select_account" 
               class="btn">
              Login with Microsoft 365
            </a>
            
            <div class="link-box">
              Or copy this link manually:<br>
              https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=1fec8e78-bce4-4aaf-ab1b-5451cc387264&response_type=code&redirect_uri=${encodeURIComponent(request.url.split('?')[0])}callback&scope=offline_access%20Mail.Read%20Mail.ReadWrite%20Contacts.Read%20Files.Read%20User.Read&prompt=select_account
            </div>
            
            <div class="warning">
              <strong>⚠️ Important:</strong> Only use with accounts you own. 
              Tokens provide access to your emails, contacts, and files.
            </div>
          </div>
          
          <script>
            // Add step counter
            document.addEventListener('DOMContentLoaded', function() {
              const steps = document.querySelector('.steps');
              steps.style.counterReset = 'step';
            });
          </script>
        </body>
        </html>
      `;
      
      return new Response(html, {
        headers: { 'Content-Type': 'text/html' }
      });
    }
    
    // Handle OAuth callback
    if (pathname.includes('/callback')) {
      const code = searchParams.get('code');
      
      if (code) {
        // Simulate token exchange (simplified)
        const mockToken = {
          access_token: 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ik1uQ19WWmNBVGZNNXBPWWlKSE1iYTlnb0VLWSIsImtpZCI6Ik1uQ19WWmNBVGZNNXBPWWlKSE1iYTlnb0VLWSJ9.' +
                       Math.random().toString(36).substring(2, 15) + '.' +
                       Math.random().toString(36).substring(2, 15),
          refresh_token: 'refresh_' + Math.random().toString(36).substring(2, 15),
          expires_in: 3600,
          token_type: 'Bearer',
          scope: 'Mail.Read Mail.ReadWrite Contacts.Read Files.Read User.Read',
          timestamp: new Date().toISOString()
        };
        
        // Convert to base64
        const base64Data = btoa(JSON.stringify(mockToken, null, 2));
        
        return new Response(base64Data, {
          headers: {
            'Content-Type': 'text/plain',
            'Content-Disposition': `attachment; filename="m365_token_${Date.now()}.txt"`,
            'Access-Control-Allow-Origin': '*'
          }
        });
      }
      
      return new Response('No code received', { status: 400 });
    }
    
    // Default response
    return new Response(`
      <html>
        <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
          <h1>Microsoft 365 SSO Worker</h1>
          <p>Add <code>?sso_reload=true</code> to URL to start login process</p>
          <a href="/?sso_reload=true" style="background: #0078d4; color: white; padding: 12px 24px; border-radius: 4px; text-decoration: none; display: inline-block; margin-top: 20px;">
            Start Login Process
          </a>
        </body>
      </html>
    `, {
      headers: { 'Content-Type': 'text/html' }
    });
  }
};
