const options =
  new MicrosoftGraph.MSALAuthenticationProviderOptions([
    'user.read',
    'onlineMeetings.ReadWrite'
  ]);
const authProvider =
  new MicrosoftGraph.ImplicitMSALAuthenticationProvider(msalClient, options);
const graphClient = MicrosoftGraph.Client.initWithMiddleware({authProvider});

var mcount = 1;

async function createMeeting() {

    var d = new Date();
    var di = d.toISOString();

    d.setMinutes(d.getMinutes() + 30);
    var df = d.toISOString();

    console.log(di);
    console.log(df);

    var onlineMeeting = {
        startDateTime: di,
        endDateTime: df,
        subject:"Teams Meeting " + mcount
      };

      mcount++;
    try {
        let res = await graphClient
        .api('/me/onlineMeetings')
        .post(onlineMeeting);
        
        updatePage(msalClient.getAccount(), Views.meeting, res);
    } catch (error) {
        updatePage(msalClient.getAccount(), Views.error, {
            message: 'Error creating the meeting',
            debug: error 
        });
    }
}
async function getEvents() {
    try {
      let events = await graphClient
          .api('/me/events')
          .select('subject,organizer,start,end')
          .orderby('createdDateTime DESC')
          .get();
  
      updatePage(msalClient.getAccount(), Views.calendar, events);
    } catch (error) {
      updatePage(msalClient.getAccount(), Views.error, {
        message: 'Error getting events',
        debug: error
      });
    }
  }