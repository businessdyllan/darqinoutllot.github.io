Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("analyze-button").onclick = analyzeEmail;
  }
});

async function analyzeEmail() {
  try {
    const item = Office.context.mailbox.item;
    const subject = item.subject;
    const body = await getEmailBody();
    const sender = item.from.displayName;

    const analysis = await callOpenAI(subject, body, sender);
    document.getElementById("analysis-result").innerHTML = analysis;
  } catch (error) {
    console.error("Error:", error);
    document.getElementById("analysis-result").innerHTML = "An error occurred. Please try again.";
  }
}

async function getEmailBody() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Failed to get email body"));
      }
    });
  });
}

async function analyzeWithChatGPT(subject, body, sender) {
    const API_KEY = 'sk-ilDQ40_bxSyU8_Gdb_oT1ovwx6_0uK5vaiEgJ6QoUBT3BlbkFJDxuKB-SEMXjFqdVTk51OT1u_LDGfYIfK9A6E3-vWcA';
    const API_ENDPOINT = 'https://api.openai.com/v1/chat/completions';

    const prompt = `Analyze the following email:

From: ${sender}
Subject: ${subject}

Body:
${body}

Provide a brief summary, sentiment analysis, and priority level.`;

    const response = await axios.post(API_ENDPOINT, {
        model: "gpt-3.5-turbo",
        messages: [
            { role: "system", content: "Tu es un assistant commercial qui doit analyser le mail, donner son avis sur l'état psychologique des interlocuteurs et générer une réponse" },
            { role: "user", content: prompt }
        ]
    }, {
        headers: {
            'Authorization': `Bearer ${API_KEY}`,
            'Content-Type': 'application/json'
        }
    });
    if (!response.ok) {
    throw new Error(`HTTP error! status: ${response.status}`);
  }

  const data = await response.json();
    return response.data.choices[0].message.content;
}
