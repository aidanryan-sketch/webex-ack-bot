require('dotenv').config();
const express = require('express');
const Webex = require('webex');

const webex = Webex.init({
  credentials: { access_token: process.env.WEBEX_BOT_TOKEN }
});

const app = express();
app.use(express.json());

const activeAlerts = new Map();
const ESCALATION_TIMEOUT_MS = 30000; // 30 seconds for quick testing

function buildAlertCard(alertId, text, isEscalated = false) {
  const headerText = isEscalated 
    ? "🚨 ESCALATED CLINICAL ALERT (No response from Person B)"
    : "🚨 URGENT CLINICAL ALERT";

  const actionTitle = isEscalated 
    ? "✅ Acknowledge Escalated Alert" 
    : "✅ Acknowledge & Accept";

  return {
    contentType: "application/vnd.microsoft.card.adaptive",
    content: {
      \$schema: "http://adaptivecards.io/schemas/adaptive-card.json",
      type: "AdaptiveCard",
      version: "1.3",
      body: [
        {
          type: "TextBlock",
          text: headerText,
          weight: "Bolder",
          size: "Large",
          color: "Attention"
        },
        {
          type: "TextBlock",
          text: text,
          wrap: true
        }
      ],
      actions: [
        {
          type: "Action.Submit",
          title: actionTitle,
          data: { action: "acknowledge", alertId: alertId }
        }
      ]
    }
  };
}

async function triggerEscalation(alertId) {
  const alertData = activeAlerts.get(alertId);
  if (!alertData) return;

  console.log(`\n⏰ [TIMEOUT] Alert ${alertId} was not acknowledged by Person B. Escalating to Person C...`);

  try {
    // 1. Update Person B's card to show timeout
    await webex.messages.update(alertData.messageId, {
      toPersonEmail: process.env.PERSON_B_EMAIL,
      attachments: [{
        contentType: "application/vnd.microsoft.card.adaptive",
        content: {
          \$schema: "http://adaptivecards.io/schemas/adaptive-card.json",
          type: "AdaptiveCard",
          version: "1.3",
          body: [
            {
              type: "TextBlock",
              text: "⚠️ UNACKNOWLEDGED — ESCALATED TO PERSON C",
              weight: "Bolder",
              color: "Warning",
              size: "Medium"
            },
            {
              type: "TextBlock",
              text: `Alert ID: ${alertId} | Timed out at: ${new Date().toLocaleTimeString()}`,
              isSubtle: true
            }
          ]
        }
      }]
    });

    // 2. Send card to Person C
    const escalatedMsg = await webex.messages.create({
      toPersonEmail: process.env.PERSON_C_EMAIL,
      markdown: `🚨 **ESCALATED ALERT**: ${alertData.text}`,
      attachments: [buildAlertCard(alertId, alertData.text, true)]
    });

    console.log(`[ESCALATED SENT] Delivered to Person C (${process.env.PERSON_C_EMAIL})`);

    // 3. Notify Person A
    if (process.env.PERSON_A_EMAIL) {
      await webex.messages.create({
        toPersonEmail: process.env.PERSON_A_EMAIL,
        markdown: `ℹ️ **Alert Status**: Alert \`\${alertId}\` timed out for Person B and was escalated to Person C.`
      });
    }

    activeAlerts.delete(alertId);
  } catch (err) {
    console.error('❌ Error during escalation:', err.message || err);
  }
}

// Endpoint to trigger test alert
app.get('/send-alert', async (req, res) => {
  const alertId = `ALERT_${Date.now()}`;
  const text = "Bed 4: Immediate clinical review required.";

  try {
    const msg = await webex.messages.create({
      toPersonEmail: process.env.PERSON_B_EMAIL,
      markdown: `🚨 URGENT ALERT: ${text}`,
      attachments: [buildAlertCard(alertId, text)]
    });

    const timer = setTimeout(() => {
      triggerEscalation(alertId);
    }, ESCALATION_TIMEOUT_MS);

    activeAlerts.set(alertId, {
      messageId: msg.id,
      text: text,
      timer: timer
    });

    console.log(`\n[ALERT SENT] ID: ${alertId} | Sent to Person B | Timer: ${ESCALATION_TIMEOUT_MS / 1000}s`);
    res.json({ success: true, alertId, messageId: msg.id });
  } catch (err) {
    console.error('Failed to send alert:', err);
    res.status(500).json({ error: err.message });
  }
});

// Webhook endpoint to catch button clicks
app.post('/webhook', async (req, res) => {
  res.status(200).send('OK');

  const { resource, event, data } = req.body;

  if (resource === 'attachmentActions' && event === 'created') {
    try {
      const action = await webex.attachmentActions.get(data.id);
      const { action: userAction, alertId } = action.inputs;
      const messageId = action.messageId;
      const personId = action.personId;

      const person = await webex.people.get(personId);

      if (userAction === 'acknowledge') {
        console.log(`\n[ACKNOWLEDGED] Alert ${alertId} acknowledged by ${person.displayName}`);

        if (activeAlerts.has(alertId)) {
          clearTimeout(activeAlerts.get(alertId).timer);
          activeAlerts.delete(alertId);
          console.log(`[TIMER CLEARED] Escalation canceled for ${alertId}`);
        }

        await webex.messages.update(messageId, {
          toPersonEmail: person.emails[0],
          attachments: [{
            contentType: "application/vnd.microsoft.card.adaptive",
            content: {
              \$schema: "http://adaptivecards.io/schemas/adaptive-card.json",
              type: "AdaptiveCard",
              version: "1.3",
              body: [
                {
                  type: "TextBlock",
                  text: `✅ ACKNOWLEDGED by ${person.displayName}`,
                  weight: "Bolder",
                  color: "Good",
                  size: "Medium"
                },
                {
                  type: "TextBlock",
                  text: `Alert ID: ${alertId} | Time: ${new Date().toLocaleTimeString()}`,
                  isSubtle: true
                }
              ]
            }
          }]
        });

        if (process.env.PERSON_A_EMAIL && person.emails[0] !== process.env.PERSON_A_EMAIL) {
          await webex.messages.create({
            toPersonEmail: process.env.PERSON_A_EMAIL,
            markdown: `✅ **Alert Acknowledged**: Alert \`\${alertId}\` accepted by ${person.displayName}.`
          });
        }
      }
    } catch (err) {
      console.error('Error handling webhook:', err);
    }
  }
});

app.listen(process.env.PORT || 3000, () => {
  console.log(`\n🚀 Server listening on port ${process.env.PORT || 3000}`);
});