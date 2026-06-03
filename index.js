import express from 'express';
import fetch from 'node-fetch';
import * as msal from '@azure/msal-node';

const app = express();

const msalConfig = {
  auth: {
    clientId: process.env.VITE_CLIENT_ID,
    clientSecret: process.env.VITE_CLIENT_SECRET,
    authority: `https://login.microsoftonline.com/${process.env.VITE_TENANT_ID}`
  }
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

async function getToken() {
  const result = await cca.acquireTokenByClientCredential({
    scopes: ['https://mmmalufconsultoria.sharepoint.com/.default']
  });

  return result.accessToken;
}

app.get('/video', async (req, res) => {
  try {
    const token = await getToken();

    const fileUrl = `https://mmmalufconsultoria.sharepoint.com${req.query.path}/$value`;

    const headers = {
      Authorization: `Bearer ${token}`
    };

    if (req.headers.range) {
      headers.Range = req.headers.range;
    }

    const spResp = await fetch(fileUrl, { headers });

    if (!spResp.ok) {
      const errText = await spResp.text();
      console.error('SharePoint error:', spResp.status, errText);
      return res.status(spResp.status).send(errText);
    }

    res.set('Content-Type', 'video/mp4');
    res.set('Accept-Ranges', 'bytes');
    res.set('Access-Control-Allow-Origin', '*');

    const contentRange = spResp.headers.get('content-range');
    const contentLength = spResp.headers.get('content-length');

    if (contentRange) {
      res.set('Content-Range', contentRange);
    }

    if (contentLength) {
      res.set('Content-Length', contentLength);
    }

    res.status(spResp.status);

    spResp.body.pipe(res);
  } catch (err) {
    console.error(err);
    res.status(500).json({
      error: err.message
    });
  }
});

app.listen(process.env.PORT || 3000, () => {
  console.log('Proxy rodando');
});
