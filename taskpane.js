Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("generateBtn").onclick = handleGenerate;
        document.getElementById("saveKeyBtn").onclick = saveApiKey;

        // 保存済みキーがあれば反映
        const savedKey = localStorage.getItem("GEMINI_API_KEY");
        if (savedKey) {
            document.getElementById("apiKeyInput").value = savedKey;
        }
    }
});

function saveApiKey() {
    const key = document.getElementById("apiKeyInput").value.trim();
    if (key) {
        localStorage.setItem("GEMINI_API_KEY", key);
        document.getElementById("status").innerText = "APIキーを端末に安全に保存しました！";
    }
}

async function handleGenerate() {
    const status = document.getElementById("status");
    const points = document.getElementById("userPoints").value.trim();
    const apiKey = localStorage.getItem("GEMINI_API_KEY") || document.getElementById("apiKeyInput").value.trim();
    
    if (!apiKey) {
        status.innerText = "下の「⚙ APIキー設定」からAPIキーを保存してください。";
        return;
    }

    if (!points) {
        status.innerText = "盛り込む要点を入力してください。";
        return;
    }

    status.innerText = "メール本文を読み取り、返信文を生成中...";

    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            status.innerText = "メール本文の取得に失敗しました。";
            return;
        }

        const emailBody = result.value;

        const prompt = `あなたはビジネスメール作成のエキスパートです。
以下の【受信メール】に対する返信メールの本文を作成してください。
【盛り込みたい要点】を過不足なく反映し、丁寧で自然なビジネス敬語を用いてください。
件名や署名は含めず、本文（相手への挨拶・名乗りから結びの言葉まで）のみを出力してください。

【盛り込みたい要点】
${points}

【受信メール】
${emailBody}`;

        try {
            const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=${apiKey}`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: prompt }] }]
                })
            });

            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.error ? errData.error.message : "APIリクエスト失敗");
            }

            const data = await response.json();
            const replyText = data.candidates[0].content.parts[0].text;

            Office.context.mailbox.item.body.setSelectedDataAsync(
                replyText,
                { coercionType: Office.CoercionType.Text },
                (asyncResult) => {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        status.innerText = "返信文をメール本文に挿入しました！";
                    } else {
                        status.innerText = "本文への挿入に失敗しました。";
                    }
                }
            );
        } catch (error) {
            status.innerText = "エラー: " + error.message;
        }
    });
}
