// Microsoft 365 Token Generator - Working Cloudflare Worker
// No OAuth issues - Uses manual Graph Explorer token method

addEventListener('fetch', event => {
  event.respondWith(handleRequest(event.request))
})

async function handleRequest(request) {
  const url = new URL(request.url)
  const path = url.pathname
  
  // Main page with sso_reload
  if (url.searchParams.get('sso_reload') === 'true') {
    return generateTokenPage(request)
  }
  
  // API endpoint to create token file
  if (path === '/generate' || path === '/api/generate') {
    return handleTokenGeneration(request)
  }
  
  // Health check
  if (path === '/health') {
    return new Response(JSON.stringify({
      status: 'ok',
      service: 'm365-token-generator',
      timestamp: new Date().toISOString()
    }), {
      headers: { 'Content-Type': 'application/json' }
    })
  }
  
  // Default page
  return renderMainPage(request)
}

function renderMainPage(request) {
  const html = `
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Microsoft 365 Token Generator</title>
      <style>
        * {
          margin: 0;
          padding: 0;
          box-sizing: border-box;
        }
        
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          min-height: 100vh;
          display: flex;
          align-items: center;
          justify-content: center;
          padding: 20px;
        }
        
        .container {
          background: rgba(255, 255, 255, 0.95);
          backdrop-filter: blur(10px);
          border-radius: 20px;
          padding: 40px;
          max-width: 800px;
          width: 100%;
          box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
        }
        
        h1 {
          color: #2d3748;
          margin-bottom: 20px;
          text-align: center;
          font-size: 2.5em;
        }
        
        .subtitle {
          color: #4a5568;
          text-align: center;
          margin-bottom: 30px;
          font-size: 1.1em;
        }
        
        .features {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
          gap: 20px;
          margin: 30px 0;
        }
        
        .feature {
          background: white;
          padding: 20px;
          border-radius: 10px;
          text-align: center;
          box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
          transition: transform 0.3s;
        }
        
        .feature:hover {
          transform: translateY(-5px);
        }
        
        .feature-icon {
          font-size: 2em;
          margin-bottom: 10px;
        }
        
        .feature-title {
          color: #2d3748;
          margin-bottom: 10px;
          font-weight: 600;
        }
        
        .feature-desc {
          color: #718096;
          font-size: 0.9em;
        }
        
        .btn {
          display: inline-block;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
          padding: 16px 32px;
          border-radius: 50px;
          text-decoration: none;
          font-weight: 600;
          font-size: 1.1em;
          text-align: center;
          cursor: pointer;
          border: none;
          transition: all 0.3s;
          margin: 10px;
          width: 100%;
          max-width: 300px;
        }
        
        .btn:hover {
          transform: translateY(-2px);
          box-shadow: 0 10px 20px rgba(102, 126, 234, 0.4);
        }
        
        .btn-container {
          text-align: center;
          margin-top: 30px;
        }
        
        .info-box {
          background: #e8f4fd;
          border-left: 4px solid #2196f3;
          padding: 20px;
          border-radius: 0 10px 10px 0;
          margin: 20px 0;
        }
        
        .steps {
          counter-reset: step;
          margin: 30px 0;
        }
        
        .step {
          background: white;
          padding: 20px;
          border-radius: 10px;
          margin-bottom: 15px;
          position: relative;
          padding-left: 70px;
          box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }
        
        .step:before {
          counter-increment: step;
          content: counter(step);
          position: absolute;
          left: 20px;
          top: 50%;
          transform: translateY(-50%);
          width: 36px;
          height: 36px;
          background: #667eea;
          color: white;
          border-radius: 50%;
          display: flex;
          align-items: center;
          justify-content: center;
          font-weight: bold;
          font-size: 1.2em;
        }
        
        .warning {
          background: #fff5f5;
          border: 2px solid #fed7d7;
          border-radius: 10px;
          padding: 20px;
          margin: 20px 0;
          color: #c53030;
        }
        
        .code {
          background: #2d3748;
          color: #e2e8f0;
          padding: 15px;
          border-radius: 8px;
          font-family: 'Courier New', monospace;
          margin: 10px 0;
          overflow-x: auto;
          font-size: 0.9em;
        }
        
        .token-input {
          width: 100%;
          padding: 15px;
          border: 2px solid #e2e8f0;
          border-radius: 10px;
          font-family: monospace;
          font-size: 0.9em;
          resize: vertical;
          min-height: 100px;
          margin: 15px 0;
        }
        
        .token-input:focus {
          outline: none;
          border-color: #667eea;
        }
        
        .loading {
          display: none;
          text-align: center;
          margin: 20px 0;
        }
        
        .spinner {
          border: 4px solid #f3f3f3;
          border-top: 4px solid #667eea;
          border-radius: 50%;
          width: 40px;
          height: 40px;
          animation: spin 1s linear infinite;
          margin: 0 auto 10px;
        }
        
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        
        .success {
          background: #f0fff4;
          border: 2px solid #9ae6b4;
          border-radius: 10px;
          padding: 20px;
          margin: 20px 0;
          color: #276749;
          display: none;
        }
        
        @media (max-width: 768px) {
          .container {
            padding: 20px;
          }
          
          h1 {
            font-size: 2em;
          }
          
          .features {
            grid-template-columns: 1fr;
          }
        }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>🔐 Microsoft 365 Token Generator</h1>
        <p class="subtitle">Get your access token for complete email scanning</p>
        
        <div class="features">
          <div class="feature">
            <div class="feature-icon">📂</div>
            <div class="feature-title">All Folders</div>
            <div class="feature-desc">Scan Inbox, Archive, Sent, and custom folders</div>
          </div>
          
          <div class="feature">
            <div class="feature-icon">🛡️</div>
            <div class="feature-title">No Errors</div>
            <div class="feature-desc">Avoids AADSTS7000112 and other OAuth issues</div>
          </div>
          
          <div class="feature">
            <div class="feature-icon">⚡</div>
            <div class="feature-title">Fast & Reliable</div>
            <div class="feature-desc">Direct Graph Explorer method - always works</div>
          </div>
        </div>
        
        <div class="btn-container">
          <a href="/?sso_reload=true" class="btn">🚀 Generate Token File</a>
        </div>
        
        <div class="info-box">
          <strong>ℹ️ How it works:</strong>
          <p>This tool helps you generate a token file that works with the M365 Data Extractor script. 
          No OAuth registration needed - uses Microsoft's own Graph Explorer service.</p>
        </div>
        
        <div class="warning">
          <strong>⚠️ Important Security Notice:</strong>
          <p>• Only use with accounts you own</p>
          <p>• Tokens expire after 1 hour</p>
          <p>• Keep token files secure</p>
          <p>• Delete token files after use</p>
        </div>
        
        <div class="steps">
          <div class="step">
            <strong>Open Microsoft Graph Explorer</strong>
            <p>Visit: <a href="https://developer.microsoft.com/en-us/graph/graph-explorer" target="_blank">https://developer.microsoft.com/en-us/graph/graph-explorer</a></p>
          </div>
          
          <div class="step">
            <strong>Sign In with Microsoft</strong>
            <p>Click the "Sign In" button and login with your Microsoft 365 account</p>
          </div>
          
          <div class="step">
            <strong>Get Your Token</strong>
            <p>After signing in, press F12 → Network tab → Run any query → Find "Authorization: Bearer" header</p>
          </div>
          
          <div class="step">
            <strong>Generate Token File</strong>
            <p>Return here, paste your token, and download the .txt file</p>
          </div>
        </div>
      </div>
    </body>
    </html>
  `
  
  return new Response(html, {
    headers: {
      'Content-Type': 'text/html',
      'Cache-Control': 'no-cache'
    }
  })
}

function generateTokenPage(request) {
  const html = `
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Generate Token File</title>
      <style>
        * {
          margin: 0;
          padding: 0;
          box-sizing: border-box;
        }
        
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          min-height: 100vh;
          display: flex;
          align-items: center;
          justify-content: center;
          padding: 20px;
        }
        
        .container {
          background: rgba(255, 255, 255, 0.95);
          backdrop-filter: blur(10px);
          border-radius: 20px;
          padding: 40px;
          max-width: 700px;
          width: 100%;
          box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
        }
        
        h1 {
          color: #2d3748;
          margin-bottom: 10px;
          text-align: center;
        }
        
        .subtitle {
          color: #4a5568;
          text-align: center;
          margin-bottom: 30px;
        }
        
        .token-input {
          width: 100%;
          padding: 15px;
          border: 2px solid #e2e8f0;
          border-radius: 10px;
          font-family: monospace;
          font-size: 0.9em;
          resize: vertical;
          min-height: 150px;
          margin: 20px 0;
        }
        
        .token-input:focus {
          outline: none;
          border-color: #667eea;
        }
        
        .btn {
          display: block;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
          padding: 16px 32px;
          border-radius: 50px;
          text-decoration: none;
          font-weight: 600;
          font-size: 1.1em;
          text-align: center;
          cursor: pointer;
          border: none;
          transition: all 0.3s;
          width: 100%;
          margin: 20px 0;
        }
        
        .btn:hover {
          transform: translateY(-2px);
          box-shadow: 0 10px 20px rgba(102, 126, 234, 0.4);
        }
        
        .btn-secondary {
          background: #718096;
        }
        
        .loading {
          display: none;
          text-align: center;
          margin: 20px 0;
        }
        
        .spinner {
          border: 4px solid #f3f3f3;
          border-top: 4px solid #667eea;
          border-radius: 50%;
          width: 40px;
          height: 40px;
          animation: spin 1s linear infinite;
          margin: 0 auto 10px;
        }
        
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        
        .success {
          background: #f0fff4;
          border: 2px solid #9ae6b4;
          border-radius: 10px;
          padding: 20px;
          margin: 20px 0;
          color: #276749;
          display: none;
        }
        
        .error {
          background: #fff5f5;
          border: 2px solid #fed7d7;
          border-radius: 10px;
          padding: 20px;
          margin: 20px 0;
          color: #c53030;
          display: none;
        }
        
        .instructions {
          background: #e8f4fd;
          border-radius: 10px;
          padding: 20px;
          margin: 20px 0;
        }
        
        .instruction-step {
          display: flex;
          align-items: center;
          margin: 10px 0;
        }
        
        .step-number {
          width: 30px;
          height: 30px;
          background: #667eea;
          color: white;
          border-radius: 50%;
          display: flex;
          align-items: center;
          justify-content: center;
          margin-right: 15px;
          font-weight: bold;
        }
        
        .clipboard-btn {
          background: #48bb78;
          color: white;
          border: none;
          padding: 8px 16px;
          border-radius: 5px;
          cursor: pointer;
          font-size: 0.9em;
          margin-left: 10px;
        }
        
        .clipboard-btn:hover {
          background: #38a169;
        }
        
        .code-block {
          background: #2d3748;
          color: #e2e8f0;
          padding: 15px;
          border-radius: 8px;
          font-family: monospace;
          margin: 10px 0;
          overflow-x: auto;
          font-size: 0.9em;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <h1>📝 Generate Your Token File</h1>
        <p class="subtitle">Paste your Graph Explorer token below</p>
        
        <div class="instructions">
          <div class="instruction-step">
            <div class="step-number">1</div>
            <div>
              <strong>Get your token from Graph Explorer</strong>
              <p>Make sure you're signed in and have copied the token from the Authorization header</p>
            </div>
          </div>
          
          <div class="instruction-step">
            <div class="step-number">2</div>
            <div>
              <strong>Paste token here</strong>
              <p>The token should start with "eyJ" and be very long (500+ characters)</p>
            </div>
          </div>
          
          <div class="instruction-step">
            <div class="step-number">3</div>
            <div>
              <strong>Download token file</strong>
              <p>Click generate to create your downloadable .txt file</p>
            </div>
          </div>
        </div>
        
        <div class="code-block" id="sampleToken">
          eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ik1uQ19WWmNBVGZNNXBPWWlKSE1iYTlnb0VLWSIsImtpZCI6Ik1uQ19WWmNBVGZNNXBPWWlKSE1iYTlnb0VLWSJ9.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC8...
          <button class="clipboard-btn" onclick="copySample()">📋 Copy Sample</button>
        </div>
        
        <textarea 
          id="tokenInput" 
          class="token-input" 
          placeholder="Paste your Graph Explorer token here (the long string starting with eyJ...)"
        ></textarea>
        
        <div class="loading" id="loading">
          <div class="spinner"></div>
          <p>Generating token file...</p>
        </div>
        
        <div class="success" id="success">
          <h3>✅ Token File Generated Successfully!</h3>
          <p>Your token file has been generated and will download automatically.</p>
          <p>Use this file with the M365 Data Extractor script.</p>
          <button class="btn btn-secondary" onclick="resetForm()">Generate Another</button>
        </div>
        
        <div class="error" id="error">
          <h3>❌ Error Generating Token File</h3>
          <p id="errorMessage"></p>
          <button class="btn btn-secondary" onclick="resetForm()">Try Again</button>
        </div>
        
        <button class="btn" onclick="generateTokenFile()">🚀 Generate Token File</button>
        <a href="/" class="btn btn-secondary">← Back to Home</a>
      </div>
      
      <script>
        // Try to read from clipboard on page load
        window.addEventListener('DOMContentLoaded', async () => {
          try {
            const clipboardText = await navigator.clipboard.readText();
            if (clipboardText.includes('eyJ') && clipboardText.includes('.')) {
              document.getElementById('tokenInput').value = clipboardText;
            }
          } catch (err) {
            // Can't read clipboard, that's fine
          }
        });
        
        function copySample() {
          const sample = document.getElementById('sampleToken').textContent.split('...')[0];
          navigator.clipboard.writeText(sample).then(() => {
            alert('Sample token copied to clipboard! Replace with your actual token.');
          });
        }
        
        function generateTokenFile() {
          const tokenInput = document.getElementById('tokenInput');
          const token = tokenInput.value.trim();
          
          if (!token) {
            alert('Please paste your token first!');
            return;
          }
          
          // Validate token format
          if (!token.startsWith('eyJ') || token.length < 100) {
            if (!confirm('This doesn\'t look like a valid JWT token. Continue anyway?')) {
              return;
            }
          }
          
          // Show loading
          document.getElementById('loading').style.display = 'block';
          document.getElementById('success').style.display = 'none';
          document.getElementById('error').style.display = 'none';
          
          // Clean token (remove "Bearer " if present)
          let cleanToken = token;
          if (cleanToken.startsWith('Bearer ')) {
            cleanToken = cleanToken.substring(7);
          }
          
          // Prepare data
          const tokenData = {
            access_token: cleanToken,
            token_type: 'Bearer',
            source: 'graph_explorer',
            generated_at: new Date().toISOString(),
            expires_in: 3600,
            scope: 'Mail.Read Mail.ReadWrite Contacts.Read Files.Read User.Read',
            instructions: 'Use with M365 Complete Folder Scanner script'
          };
          
          // Convert to base64
          const jsonStr = JSON.stringify(tokenData, null, 2);
          const base64Data = btoa(unescape(encodeURIComponent(jsonStr)));
          
          // Create download link
          const blob = new Blob([base64Data], { type: 'text/plain' });
          const url = URL.createObjectURL(blob);
          
          const a = document.createElement('a');
          a.href = url;
          a.download = \`m365_token_\${Date.now()}.txt\`;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url);
          
          // Hide loading, show success
          setTimeout(() => {
            document.getElementById('loading').style.display = 'none';
            document.getElementById('success').style.display = 'block';
            tokenInput.value = '';
          }, 1000);
        }
        
        function resetForm() {
          document.getElementById('tokenInput').value = '';
          document.getElementById('loading').style.display = 'none';
          document.getElementById('success').style.display = 'none';
          document.getElementById('error').style.display = 'none';
        }
        
        // Handle paste events
        document.getElementById('tokenInput').addEventListener('paste', (e) => {
          setTimeout(() => {
            const text = e.target.value;
            if (text.startsWith('Bearer ')) {
              e.target.value = text.substring(7);
            }
          }, 10);
        });
      </script>
    </body>
    </html>
  `
  
  return new Response(html, {
    headers: {
      'Content-Type': 'text/html',
      'Cache-Control': 'no-cache'
    }
  })
}

async function handleTokenGeneration(request) {
  // API endpoint for programmatic token generation
  if (request.method !== 'POST') {
    return new Response(JSON.stringify({ error: 'Method not allowed' }), {
      status: 405,
      headers: { 'Content-Type': 'application/json' }
    })
  }
  
  try {
    const data = await request.json()
    const { token } = data
    
    if (!token) {
      return new Response(JSON.stringify({ error: 'No token provided' }), {
        status: 400,
        headers: { 'Content-Type': 'application/json' }
      })
    }
    
    // Clean token
    let cleanToken = token.trim()
    if (cleanToken.startsWith('Bearer ')) {
      cleanToken = cleanToken.substring(7)
    }
    
    // Create token data
    const tokenData = {
      access_token: cleanToken,
      token_type: 'Bearer',
      source: 'api_request',
      generated_at: new Date().toISOString(),
      expires_in: 3600
    }
    
    // Convert to base64
    const jsonStr = JSON.stringify(tokenData, null, 2)
    const base64Data = btoa(unescape(encodeURIComponent(jsonStr)))
    
    return new Response(base64Data, {
      headers: {
        'Content-Type': 'text/plain',
        'Content-Disposition': \`attachment; filename="m365_token_\${Date.now()}.txt"\`,
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type'
      }
    })
    
  } catch (error) {
    return new Response(JSON.stringify({ error: error.message }), {
      status: 500,
      headers: { 'Content-Type': 'application/json' }
    })
  }
}
