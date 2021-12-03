const config = require('./config');
const https = require('https');

async function getToken(){
  return new Promise((resolve, reject) => {
    //https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
    let body = `${encodeURI('client_id')}=${encodeURI(config.clientId)}&${encodeURI('scope')}=${encodeURI(`https://${config.graph_endpoint}/.default`)}&${encodeURI('client_secret')}=${encodeURI(config.clientSecret)}&${encodeURI('grant_type')}=${encodeURI('client_credentials')}`;

    const req = https.request({
      hostname: config.ad_endpoint,
      path: `/${config.tenant}.onmicrosoft.com/oauth2/v2.0/token`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body)
      }
    }, (res) => {
      let data = '';

      res.on('data', (chunk) => {
        data += chunk;
      });

      res.on('end', () => {          
        if(res.statusCode !== 200) reject(res.statusMessage);

        resolve(JSON.parse(data));
      });
    }
    );

    req.on('error', (e) => {
      reject(`problem with request: ${e.message}`);
    });

    req.write(body);

    req.end();
  });
}

//https://docs.microsoft.com/en-us/graph/api/resources/teams-api-overview?view=graph-rest-1.0
async function listUsers(query) {
  let jwt = undefined;
  await this.getToken().then(tokenRes => { jwt = tokenRes.access_token });

  return new Promise((resolve, reject) => {
    const req = https.get({
      hostname: config.graph_endpoint,
      path: `/v1.0/users${query}`,
      headers: { 'Authorization': `Bearer ${jwt}`}
    }, (res) => {
      let data = '';

      res.on('data', (chunk) => {
        data += chunk;
      });

      res.on('end', () => {
        if(res.statusCode !== 200) reject(JSON.parse(data));
        
        resolve(JSON.parse(data));
      });
    });
    
    req.on('error', (e) => {
      reject(`problem with request: ${e.message}`);
    });

    req.end();
  });
}

async function listTeams(query) {
  let jwt = undefined;
  await this.getToken().then(tokenRes => { jwt = tokenRes.access_token });

  return new Promise((resolve, reject) => {
    const req = https.get({
      hostname: config.graph_endpoint,
      path: `/v1.0/groups${query}`,
      headers: { 'Authorization': `Bearer ${jwt}`}
    }, (res) => {
      let data = '';

      res.on('data', (chunk) => {
        data += chunk;
      });

      res.on('end', () => {
        if(res.statusCode !== 200) reject(JSON.parse(data));
        
        resolve(JSON.parse(data));
      });
    });
    
    req.on('error', (e) => {
      reject(`problem with request: ${e.message}`);
    });

    req.end();
  });
}

async function create(displayName, description) {
  let jwt = undefined;
  await this.getToken().then(tokenRes => jwt = tokenRes.access_token);

  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: config.graph_endpoint,
      path: '/v1.0/teams',
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${jwt}`,
        'Content-Type': 'application/json'
      } 
    }, (res) => {
      res.on('end', () => {
        console.log('Creation complete.');

        if(res.statusCode !== 202) reject(JSON.stringify(data));

        resolve();
      });
    });
    
    req.on('error', (e) => {
      reject(`problem with request: ${e.message}`);
    });

    req.write(JSON.stringify({
      'template@odata.bind': 'https://graph.microsoft.com/v1.0/teamsTemplates(\'standard\')',
      'displayName': displayName,
      'description': description,
      'members': [
        {
            '@odata.type':'#microsoft.graph.aadUserConversationMember',
            'roles': [
              'owner'
            ],
            'user@odata.bind':'https://graph.microsoft.com/v1.0/users(\'cad677c7-4c0b-40d2-a3f5-960db36c665e\')'
        }
      ]
     }));
    
    req.end();
  });
}

module.exports = {
  team: {
    getToken: getToken,
    listUsers: listUsers,
    listTeams: listTeams,
    create: create
  }
}