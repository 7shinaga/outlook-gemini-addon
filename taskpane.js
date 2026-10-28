// 1行目のクォーテーションの中に、先ほど取得したAPIキーを貼り付けてください
const GEMINI_API_KEY = "AQ.Ab8RN6IpSjka4sqf_gwFyDPzpjK3gGpfy1a974Fua6xhTcKUzg";

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("generateBtn").onclick = handleGenerate;
    }
});

async function handleGenerate() {
    const status = document.getElementById("status");
    const points = document.getElementById("userPoints").value.trim();
    
    if (!points) {
        status.innerText = "盛り込む要点を入力してください。";
        return;
    }

    status.innerText = "メール本文を読み取り、返信文を生成中...";

    // 返信元（受信メール）の本文を取得
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            status.innerText = "メール本文の取得に失敗しました。";
            return;
        }

        const emailBody = result.value;

        // Gemini APIへ送るプロンプト
        const prompt = `あなたはビジネスメール作成のエキスパートです。
以下の【受信メール】に対する返信メールの本文を作成してください。
【盛り込みたい要点】を過不足なく反映し、丁寧で自然なビジネス敬語を用いてください。
件名や署名は含めず、本文（相手への挨拶・名乗りから結びの言葉まで）のみを出力してください。

【盛り込みたい要点】
${points}

【受信メール】
${emailBody}`;

        try {
            const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`, {
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

            // 作成中のメール本文に返信文を挿入
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