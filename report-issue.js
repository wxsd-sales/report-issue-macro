/*********************************************************
 * 
 * Author:              William Mills
 *                    	Technical Solutions Specialist 
 *                    	wimills@cisco.com
 *                    	Cisco Systems
 * 
 * Version: 1-0-0
 * Released: 02/09/23
 * 
 * This is a Webex Device macro which lets a user select issue categories 
 * and enter issue details which are sent to a Webhook service
 * 
 * 
 * Full Readme, source code and license details are available here:
 * https://github.com/wxsd-sales/report-issue-macro
 * 
 ********************************************************/
 
 import xapi from 'xapi';

/*********************************************************
 * Configure the settings below
**********************************************************/

const config = {
  name: 'Report Issue',          // Name of the Button and Panel
  submitText: 'Submit Issue',       // Text displays on the submit button
  waitingText: 'Sending Feedback',
  showAlert: true,                  // Show success and error alerts while true. One waiting alert is shown while false
  serviceUrl: 'https://<Your Webhook URL>',
  allowInsecureHTTPs: true,         // Allow insecure HTTPS connections to the instant connect broker for testing
  panelId: 'feedback',
  start: {
    options: [
      'Technical Issue with Incoming Audio/Video',
      'Technical Issue with Outgoing Audio/Video',
      'Can\'t connect to my meeting',
      'Request for a technician',
      'Issue with sharing content'
    ]
  },
  form: {
    category: {
      type: {
        Text: {
          prefix: '',
          options: 'size=2;fontSize=normal;align=left'
        },
        Button: {
          name: ['Select Category', 'Change Category'],
          options: 'size=2'
        }
      },
      action: 'Options',
      placeholder: 'eg. Please select category',
      promptText: 'Please enter the problem description',         // Text input message
      inputType: 'SingleLine',   // Type of input field. SingleLine = alphanum, other options (Numeric, Password, PIN)
      showPlaceholder: true,
      visiable: true,     // True = field is visable | False = field is removed from UI
      modifiable: true   // If false, placeholder will be used always
    },
    name: {
      requires: ['category'],
      type: {
        Text: {
          prefix: 'Name:',
          options: 'size=2;fontSize=normal;align=left'
        },
        Button: {
          name: ['Enter Name', 'Change Name'],
          options: 'size=2'
        }
      },
      action: 'TextInput',
      placeholder: 'eg. John Smith (optional)',
      promptText: 'Please enter your name',
      inputType: 'SingleLine',
      showPlaceholder: true,
      visiable: true,
      modifiable: true
    },
    submit: {
      requires: ['category'],
      visiable: true,
      modifiable: true,
      action: 'Submit',
      value: 'active',
      type: {
        Button: {
          name: ['Submit Issue'],
          options: 'size=2'
        }
      }
    }
  }
}


/*********************************************************
 * Main function to setup and add event listeners
**********************************************************/

/// Marco variables
let inputs = {};
let identification = {};

function main() {

  // Enable the HTTP Client
  xapi.Config.HttpClient.Mode.set('On');
  xapi.Config.HttpClient.AllowHTTP.set(config.allowInsecureHTTPs ? 'True' : 'False');

  // Get Device Details
  xapi.Status.SystemUnit.Software.DisplayName.get()
  .then(result => {identification.software = result})
  .catch(e=>console.log('Could not get DisplayName: ' + e.message))

  xapi.Status.SystemUnit.Hardware.Module.SerialNumber.get()
  .then(result => {identification.SerialNumber = result})
  .catch(e=>console.log('Could not get SerialNumber: ' + e.message))

  xapi.Status.SystemUnit.ProductId.get()
  .then(result => {identification.ProductId = result})
  .catch(e=>console.log('Could not get ProductId: ' + e.message))

  xapi.Status.Webex.DeveloperId.get()
  .then(result => {identification.deviceId = result})
  .catch(e=>console.log('Could not get Device Id: ' + e.message))

  xapi.Status.UserInterface.ContactInfo.ContactMethod[1].Number.get()
  .then(result => {identification.contactNumber = result})
  .catch(e=>console.log('Could not get Contact Number: ' + e.message))

  // Create the UI
  createPanel();

  // Monitor for Text Input Responses and Widget clicks
  xapi.Event.UserInterface.Message.TextInput.Response.on(processInput);
  xapi.Event.UserInterface.Extensions.Widget.Action.on(processWidget);

  // Reset the previous inputs when the panel is opened
  xapi.Event.UserInterface.Extensions.Panel.Clicked.on(event => {
    if (event.PanelId != config.panelId) return
    inputs = {};
    createPanel('start');
  });

}

setTimeout(main, 1000);

/*********************************************************
 * Additional functions which this macros uses
**********************************************************/


// Listen for clicks on the buttons
function processWidget(event) {
  // console.log(event);
  if (event.Type !== 'clicked') return

  console.log(event.WidgetId + ' Clicked');
  
  if(config.form.hasOwnProperty(event.WidgetId)) {
    switch (config.form[event.WidgetId].action) {
      case 'TextInput':
        createInput(event.WidgetId);
        break;
      case 'Options':
        createPanel('start');
        break;
      case 'Submit':
        xapi.Command.UserInterface.Extensions.Panel.Close();
        sendInformation();
    }
  } else if(event.WidgetId.startsWith('option')){

    const option = parseInt(event.WidgetId.slice(-1))

    console.log(`Option [${option}] selected, category [${config.start.options[option]}]`);
    inputs.category = config.start.options[option];
    createPanel();

  }
}

function createInput(type) {
  console.log('Opening Text Input for: ' + type);
  const field = config.form[type];
  let paramters = {
    FeedbackId: type,
    InputType: field.inputType,
    Placeholder: field.placeholder,
    Text: field.promptText,
    Title: config.name
  }
  if (inputs.hasOwnProperty(type)) {
    paramters.InputText = inputs[type];
  }
  xapi.Command.UserInterface.Message.TextInput.Display(paramters);
}

function processInput(event) {
  if (config.form.hasOwnProperty(event.FeedbackId)) {
    if (config.form.hasOwnProperty('regex')) {
      // Check input
    } else {
      inputs[event.FeedbackId] = event.Text;
    }
  }
  createPanel();
}

function alert(title, message, duration) {
  console.log(title + ': ' + message);
  if (!config.showAlert && !duration)
    return
  xapi.Command.UserInterface.Message.Alert.Display({
    Duration: duration ? duration : 3
    , Text: message
    , Title: title
  });
}

function parseJSON(inputString) {
  if (inputString) {
    try {
      return JSON.parse(inputString);
    } catch (e) {
      return false;
    }
  }
}

// The function will post the current inputs objects to a configured service URL
async function sendInformation() {
  alert('Sending', config.waitingText, 10);
  inputs.identification = identification;
  inputs.bookingId = await getBookingId();
  inputs.callDetails = await getCallDetails();
  inputs.conferenceDetails = await getConferenceDetails();

  console.log(JSON.stringify(inputs))
  xapi.Command.HttpClient.Post(
    {
      AllowInsecureHTTPS: true,
      Header: ["Content-Type: application/json"],
      ResultBody: "PlainText",
      Url: config.serviceUrl
    },
    JSON.stringify(inputs)
  ).then(result => {
    // Check the response from the server display the correct message
    const body = parseJSON(result.Body);
    alert('Success', 'Feedback sent, please wait for an agent to process', 10);
  })
    .catch(err => {
      alert('Error', JSON.stringify(err))
    });
}

function getCallDetails(){
  return xapi.Status.Call.get()
    .then(result => {
      console.log('Current CallId:', ( result.length > 0 ) ? result[0].id : null)
      return ( result.length > 0 ) ? result[0] : null
    });
}

function getConferenceDetails(){
  return xapi.Status.Conference.Call.get()
    .then(result => {
      return (result.length > 0) ? result : null;
    });
}

function getBookingId(){
  return xapi.Status.Bookings.Current.Id.get()
  .then(result => {
    console.log('Current Booking Id:', (result == '') ? null : result)
    return (result == '') ? null : result;
  });
}

/*********************************************************
 * This function creates/updates the UI
**********************************************************/

function arrayContains(array, contains) {
  return contains.every(element => {
    return array.indexOf(element) !== -1;
  });
}

function createPanel(state) {

  function createWidget(key, type, name, options) {
    return `<Widget>
              <WidgetId>${key}</WidgetId>
              <Name>${name}</Name>
              <Type>${type}</Type>
              <Options>${options}</Options>
            </Widget>`
  }

  let fields = '';
  let active = {};

  if (state == 'start') {
    const prompt = createWidget('category-text', 'Text', 'Please select a category below:', 'size=3;fontSize=normal;align=left')
    fields = fields.concat(`<Row>${prompt}</Row>`);
    config.start.options.forEach( (option, i) => {
      //console.log(i + ':' + option);
      const widget = createWidget('option'+i, 'Button', option, 'size=4')
      fields = fields.concat(`<Row>${widget}</Row>`);
    })
  } else {

    for (const [key, field] of Object.entries(config.form)) {

      // If not modifiable, use the placeholder as value
      if (!field.modifiable)
        inputs[key] = field.placeholder;

      // If not visiable, or no types present don't display it.
      if (!field.visiable || !field.hasOwnProperty('type'))
        continue;

      console.log(`Field [${key}] requires: [${field.requires}] | Current Inputs:${Object.keys(inputs)}`);
      // If it has requirements, check they have been met
      if (field.hasOwnProperty('requires'))
        if (!arrayContains(Object.keys(inputs), field.requires))
          continue;

      // Store any active/inactive states for setting later
      if (field.hasOwnProperty('value'))
        active[key] = field.value;


      // Create the widgets
      let widgets = '';
      for (const [type, widget] of Object.entries(field.type)) {
        // console.log(type);
        // console.log(widget);
        if (type === 'Button') {
          widgets = widgets.concat(createWidget(key, type, inputs.hasOwnProperty(key) ? widget.name[1] : widget.name[0], widget.options));
        } else if (type === 'Text' && (inputs.hasOwnProperty(key) || field.showPlaceholder)) {
          widgets = widgets.concat(createWidget(key+'-text', type, inputs.hasOwnProperty(key) ? widget.prefix + ' ' + inputs[key] : field.placeholder, widget.options));
        }
      }

      fields = fields.concat(`<Row>${widgets}</Row>`);
    }
  }

  const panel = `
  <Extensions>
    <Panel>
      <Location>HomeScreenAndCallControls</Location>
      <Type>Statusbar</Type>
      <Icon>Helpdesk</Icon>
      <Name>${config.name}</Name>
      <Color>#0067ac</Color>
      <ActivityType>Custom</ActivityType>
      <Page>
        <Name>${config.name}</Name>
        ${fields}
        <Options>hideRowNames=1</Options>
      </Page>
    </Panel>
  </Extensions>`;

  xapi.Command.UserInterface.Extensions.Panel.Save(
    { PanelId: config.panelId },
    panel
  )

  // Set active/inactive widget values
  for (const [key, value] of Object.entries(active)) {
    console.log(`Key: ${key} | Value: ${value}`);
    xapi.Command.UserInterface.Extensions.Widget.SetValue(
      { Value: value, WidgetId: key });
  }
}
