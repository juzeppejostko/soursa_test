// Create an authentication provider
const authProvider = {
    getAccessToken: async () => {
        // Call getToken in auth.js
        return await getToken();
    }
};
// Initialize the Graph client
const graphClient = MicrosoftGraph.Client.initWithMiddleware({ authProvider });
//Get user info from Graph
async function getUser() {
    ensureScope('user.read');
    return await graphClient
        .api('/me')
        .select('id,displayName')
        .get();
}

async function getMails() {
    ensureScope('user.read');
    return await graphClient
        .api('/me/messages')
        .select('sender,subject')
        .get();
}

async function listInboxAsync() {
    try {
      const messagePage = await graphHelper.getInboxAsync();
      const messages = messagePage.value;
  
      // Output each message's details
      for (const message of messages) {
        console.log(`Message: ${message.subject ?? 'NO SUBJECT'}`);
        console.log(`  From: ${message.from?.emailAddress?.name ?? 'UNKNOWN'}`);
        console.log(`  Status: ${message.isRead ? 'Read' : 'Unread'}`);
        console.log(`  Received: ${message.receivedDateTime}`);
      }
  
      // If @odata.nextLink is not undefined, there are more messages
      // available on the server
      const moreAvailable = messagePage['@odata.nextLink'] != undefined;
      console.log(`\nMore messages available? ${moreAvailable}`);
    } catch (err) {
      console.log(`Error getting user's inbox: ${err}`);
    }
  }