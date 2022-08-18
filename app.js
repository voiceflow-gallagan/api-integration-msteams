require('dotenv').config()
const process = require('node:process')
express = require('express')
axios = require('axios').default
botbuilder = require('botbuilder')
const { MessageFactory, CardFactory, TurnContext } = require('botbuilder')
localtunnel = require('localtunnel')
tunnel = null

const DMconfig = {
  tts: false,
  stripSSML: false,
}

// Create HTTP server.
const app = express()
const server = app.listen(process.env.PORT || 3978, async function () {
  const { port } = server.address()
  console.log(
    '\nServer listening on port %d in %s mode',
    port,
    app.settings.env
  )
  // Setup the tunnel for testing
  if (app.settings.env == 'development') {
    tunnel = await localtunnel({
      port: port,
      subdomain: process.env.TUNNEL_SUBDOMAIN,
    })
    console.log(`\nEndpoint: ${tunnel.url}/api/messages`)
    console.log(
      `\nGet Bot Framework Emulator: https://aka.ms/botframework-emulator`
    )
    tunnel.on('close', () => {
      // tunnels are closed
      console.log('\n\nClosing tunnel')
    })
  }
  console.log('\n')

  var i = 0 // dots counter
  setInterval(function () {
    process.stdout.clearLine() // clear current text
    process.stdout.cursorTo(0) // move cursor to beginning of line
    i = (i + 1) % 4
    var dots = new Array(i + 1).join('.')
    process.stdout.write('Listening' + dots) // write text
  }, 300)
})

// Create bot adapter, which defines how the bot sends and receives messages.
const adapter = new botbuilder.BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId,
  appPassword: process.env.MicrosoftAppPassword,
})

// Catch-all for errors.
const onTurnErrorHandler = async (context, error) => {
  // This check writes out errors to console log
  console.error(`\n [onTurnError] unhandled error: ${error}`)

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    'OnTurnError Trace',
    `${error}`,
    'https://www.botframework.com/schemas/error',
    'TurnError'
  )

  // Send a message to the user
  await context.sendActivity('The bot encountered an error or bug.')
  await context.sendActivity(
    'To continue to run this bot, please fix the bot source code.'
  )
}
// Set the onTurnError for the singleton BotFrameworkAdapter.
adapter.onTurnError = onTurnErrorHandler

// Listen for incoming requests at /api/messages.
app.post('/api/messages', async (req, res) => {
  // Use the adapter to process the incoming web request into a TurnContext object.
  adapter.processActivity(req, res, async (turnContext) => {
    // Do something with this incoming activity!
    if (turnContext.activity.type === 'message') {
      const user_id = turnContext.activity.from.id
      // Get the user's text
      const utterance = turnContext.activity.text
      let response = await interact(
        user_id,
        {
          type: 'text',
          payload: utterance,
        },
        turnContext
      )
      if (response.length > 0) {
        await sendMessage(response, turnContext)
      }
    }
  })
})

async function interact(user_id, request, turnContext) {
  // Update {user_id} variable with DM API
  await axios({
    method: 'PATCH',
    url: `${process.env.VOICEFLOW_RUNTIME_ENDPOINT}/state/user/${encodeURI(
      user_id
    )}/variables`,
    headers: {
      Authorization: process.env.VOICEFLOW_API_KEY,
      'Content-Type': 'application/json',
    },
    data: {
      user_id: user_id,
    },
  })

  // Interact with DM API
  let response = await axios({
    method: 'POST',
    url: `${process.env.VOICEFLOW_RUNTIME_ENDPOINT}/state/user/${encodeURI(
      user_id
    )}/interact`,
    headers: {
      Authorization: process.env.VOICEFLOW_API_KEY,
      'Content-Type': 'application/json',
      versionID: process.env.VOICEFLOW_VERSION,
    },
    data: {
      action: request,
      config: DMconfig,
    },
  })

  let responses = []
  for (let i = 0; i < response.data.length; i++) {
    if (response.data[i].type == 'text') {
      let tmpspeech = ''
      for (let j = 0; j < response.data[i].payload.slate.content.length; j++) {
        for (
          let k = 0;
          k < response.data[i].payload.slate.content[j].children.length;
          k++
        ) {
          if (response.data[i].payload.slate.content[j].children[k].type) {
            if (
              response.data[i].payload.slate.content[j].children[k].type ==
              'link'
            ) {
              tmpspeech +=
                response.data[i].payload.slate.content[j].children[k].url
            }
          } else if (
            response.data[i].payload.slate.content[j].children[k].text != '' &&
            response.data[i].payload.slate.content[j].children[k].fontWeight
          ) {
            tmpspeech +=
              '**' +
              response.data[i].payload.slate.content[j].children[k].text +
              '**'
          } else if (
            response.data[i].payload.slate.content[j].children[k].text != '' &&
            response.data[i].payload.slate.content[j].children[k].italic
          ) {
            tmpspeech +=
              '_' +
              response.data[i].payload.slate.content[j].children[k].text +
              '_'
          } else if (
            response.data[i].payload.slate.content[j].children[k].text != '' &&
            response.data[i].payload.slate.content[j].children[k].underline
          ) {
            tmpspeech +=
              // ignore underline
              response.data[i].payload.slate.content[j].children[k].text
          } else if (
            response.data[i].payload.slate.content[j].children[k].text != '' &&
            response.data[i].payload.slate.content[j].children[k].strikeThrough
          ) {
            tmpspeech +=
              '~' +
              response.data[i].payload.slate.content[j].children[k].text +
              '~'
          } else if (
            response.data[i].payload.slate.content[j].children[k].text != ''
          ) {
            tmpspeech +=
              response.data[i].payload.slate.content[j].children[k].text
          }
        }
        tmpspeech += '\n'
      }

      responses.push({
        type: 'text',
        value: tmpspeech,
      })
    } else if (response.data[i].type == 'visual') {
      responses.push({
        type: 'image',
        value: response.data[i].payload.image,
      })
    } else if (response.data[i].type == 'choice') {
      let buttons = []
      for (let b = 0; b < response.data[i].payload.buttons.length; b++) {
        buttons.push({
          label: response.data[i].payload.buttons[b].request.payload.label,
        })
      }
      responses.push({
        type: 'buttons',
        buttons: buttons,
      })
    }
  }

  let isEnding = response.data.filter(({ type }) => type === 'end')
  if (isEnding.length > 0) {
    console.log('\nConvo has ended')
  }

  return responses
}

async function sendMessage(messages, turnContext) {
  let responses = []
  for (let j = 0; j < messages.length; j++) {
    let message = null
    if (messages[j].type == 'image') {
      // Image
      const card = CardFactory.heroCard(null, [messages[j].value])
      message = MessageFactory.attachment(card)
    } else if (messages[j].type == 'buttons') {
      // Buttons
      let actions = []
      for (let k = 0; k < messages[j].buttons.length; k++) {
        actions.push(messages[j]?.buttons[k]?.label)
      }
      const card = CardFactory.heroCard(null, null, actions)
      message = MessageFactory.attachment(card)
    } else if (messages[j].type == 'text') {
      // Text
      message = messages[j].value
    } else {
      return
    }
    await turnContext.sendActivity(message)
  }
}

process.on('SIGINT', function () {
  process.exit()
})

process.on('exit', () => {
  if (process.env.NODE_ENV == 'development') {
    tunnel.close()
  }
  return console.log(`Bye!\n\n`)
})
