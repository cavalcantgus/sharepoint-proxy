import express from 'express';
import fetch from 'node-fetch';
import * as msal from '@azure/msal-node';

const app = express();

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`
  }
};

const cca = new msal.ConfidentialClientApplication(msalConfig);

async function getToken() {
  const result = await cca.acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default'] // <-- troca pra Graph
  });
  return result.accessToken;
}

const DRIVE_ID = 'b!vQlJtps_j0yPYJqkz3nghnAaRIU42LtLudzr-7_dcjpwXzayodOfT5TQMhcGvJa4';

app.get('/video', async (req, res) => {
  try {
    const token = await getToken();
    
    // O path chega como /sites/ServidorGeraoBancria/Teste/Servidor Geração Bancária/...
    // /Teste é a raiz do drive, então removemos até ele
    const filePath = req.query.path.replace('/sites/ServidorGeraoBancria/Teste', '');
    const encodedPath = filePath.split('/').map(encodeURIComponent).join('/');
    
    const graphUrl = `https://graph.microsoft.com/v1.0/drives/${DRIVE_ID}/root:${encodedPath}:/content`;
    console.log('Graph URL:', graphUrl);
    const headers = { Authorization: `Bearer ${token}` };
    if (req.headers.range) headers.Range = req.headers.range;

    const spResp = await fetch(graphUrl, { headers });

    if (!spResp.ok) {
      const errText = await spResp.text();
      console.error('Graph error:', spResp.status, errText);
      return res.status(spResp.status).send(errText);
    }

    res.set('Content-Type', 'video/mp4');
    res.set('Accept-Ranges', 'bytes');
    res.set('Access-Control-Allow-Origin', '*');
    
    const contentRange = spResp.headers.get('content-range');
    const contentLength = spResp.headers.get('content-length');
    if (contentRange) res.set('Content-Range', contentRange);
    if (contentLength) res.set('Content-Length', contentLength);

    res.status(spResp.status);
    spResp.body.pipe(res);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(process.env.PORT || 3000, () => {
  console.log('Proxy rodando');
});
