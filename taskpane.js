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
    const API_KEY = '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';
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
            { role: "system", content: "You are an AI assistant that analyzes emails." },
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