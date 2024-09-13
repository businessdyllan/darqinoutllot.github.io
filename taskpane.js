Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("analyze-button").onclick = analyzeEmail;
    }
});

async function analyzeEmail() {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const body = result.value;
            const subject = Office.context.mailbox.item.subject;
            const sender = Office.context.mailbox.item.from.displayName;
            
            try {
                const analysis = await analyzeWithChatGPT(subject, body, sender);
                document.getElementById("analysis-result").innerHTML = analysis;
            } catch (error) {
                document.getElementById("analysis-result").innerHTML = `Erreur: ${error.message}`;
            }
        }
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

    return response.data.choices[0].message.content;
}
