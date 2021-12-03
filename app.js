const teamsService = require('./integrations/teamsService');

(async () => {
  await teamsService.team.getToken().then(res => console.log(res)).catch(err => console.log(`ERROR: ${err}`));
  
  await teamsService.team.listUsers('?$filter=startswith(userPrincipalName,\'admin\')').then(res => console.log(res)).catch(err => console.log(`ERROR: ${err}`));
  
  await teamsService.team.listTeams('?$select=displayName,id&$filter=startswith(displayName,\'POC\')').then(res => console.log(res)).catch(err => console.log(`ERROR: ${err}`));
  
  await teamsService.team.create(`POC-${Date.now()}`).catch(err => console.log(`ERROR: ${err}`));
})();