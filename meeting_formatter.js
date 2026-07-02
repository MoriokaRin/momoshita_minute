/**
 * Google Meetの議事録メールを監視し、Gemini APIを使って
 * 指定された「ももしたOS用議事録」テンプレートに自動成形してスレッドに返信するスクリプト
 */

// 設定項目
const SEARCH_QUERY = 'from:gemini-notes@google.com subject:"「" -label:momoshita-os-processed';
const PROCESSED_LABEL = 'momoshita-os-processed';

/**
 * メイン処理：未処理の議事録メールを検索し、フォーマットして返信する
 */
function processMeetMinutes() {
  Logger.log('処理を開始します。');
  
  // 1. 重複防止用のラベルを取得または作成
  let label = GmailApp.getUserLabelByName(PROCESSED_LABEL);
  if (!label) {
    label = GmailApp.createLabel(PROCESSED_LABEL);
    Logger.log('ラベルを作成しました: ' + PROCESSED_LABEL);
  }
  
  // 2. 未処理のメールスレッドを検索
  const threads = GmailApp.search(SEARCH_QUERY);
  Logger.log('未処理のスレッド数: ' + threads.length);
  
  if (threads.length === 0) {
    Logger.log('処理対象のメールはありません。');
    return;
  }
  
  // 各スレッドを処理
  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    const messages = thread.getMessages();
    const latestMessage = messages[messages.length - 1]; // 最新のメッセージを取得
    const bodyText = latestMessage.getPlainBody();
    const mailDate = latestMessage.getDate();
    
    Logger.log('スレッドの処理中: ' + thread.getFirstMessageSubject());
    
    try {
      // 3. GoogleドキュメントのURLをメール本文から探す
      let minutesContent = '';
      const docUrlMatch = bodyText.match(/https:\/\/docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/);
      
      if (docUrlMatch && docUrlMatch[1]) {
        const docId = docUrlMatch[1];
        Logger.log('Googleドキュメントを検出しました。ID: ' + docId);
        try {
          const doc = DocumentApp.openById(docId);
          minutesContent = doc.getBody().getText();
          Logger.log('Googleドキュメントからテキストを抽出しました。');
        } catch (e) {
          Logger.log('ドキュメントの読み込みに失敗しました（権限不足などの可能性）。メール本文を代わりに使用します。エラー: ' + e.toString());
          minutesContent = bodyText;
        }
      } else {
        Logger.log('GoogleドキュメントのURLが見つかりませんでした。メール本文を処理対象にします。');
        minutesContent = bodyText;
      }
      
      // 4. Gemini APIを呼び出して「ももしたOS用議事録」を生成
      const formattedMinutes = generateMomoshitaOSMinutes(minutesContent, mailDate);
      
      // 5. 生成した議事録を自分宛てに送信し、元のスレッド内にぶら下げる（バウンス防止）
      const replyBody = "自動生成された「ももしたOS用議事録」です。\n\n" + formattedMinutes;
      const myEmail = Session.getActiveUser().getEmail();
      GmailApp.sendEmail(myEmail, "Re: " + thread.getFirstMessageSubject(), replyBody, {
        threadId: thread.getId()
      });
      Logger.log('自分宛てにスレッド紐付けメールを送信しました。');
      
      // 6. 処理済みラベルを付与
      thread.addLabel(label);
      Logger.log('処理済みラベルを付与しました。');
      
    } catch (error) {
      Logger.log('エラーが発生しました（スレッドID: ' + thread.getId() + '）: ' + error.toString());
    }
  }
  
  Logger.log('すべての処理が完了しました。');
}

/**
 * Gemini APIを呼び出して、テンプレートに沿った議事録を生成する
 * @param {string} contextText 会議の生テキスト
 * @param {Date} mailDate メールの受信日時
 * @return {string} 生成された議事録テキスト
 */
function generateMomoshitaOSMinutes(contextText, mailDate) {
  // スクリプトプロパティからAPIキーを取得
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error("スクリプトプロパティ 'GEMINI_API_KEY' が設定されていません。マニュアルに従って設定してください。");
  }
  
  // メール受信日から日付文字列（YYYY/MM/DD）を作成
  const formattedDate = Utilities.formatDate(mailDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey;
  
  // Geminiへの指示プロンプト
  const prompt = `あなたは非常に優秀なプロジェクトマネージャー、およびAIアシスタントです。
提供された「会議データ」（文字起こし、または会議メモ）を深く理解し、指定されたテンプレートに従って、論理的かつ極めて読みやすい「ももしたOS用議事録」を作成してください。

--- 議事録の品質に関する重要指示 ---
1. 【役割の付与】: 会議に参加していないメンバーがこの議事録を読むだけで、「何が」「どのような文脈で議論され」「どう決まり」「次に誰が何をすればよいか」が完璧に把握できるように記述してください。
2. 【アジェンダ】: 会議で話し合われた主要なアジェンダ（目的や議題のタイトル）を漏れなく箇条書きで抽出してください。
3. 【会議中メモ】: 単に雑音や無駄な雑談を載せるのではなく、会話から読み取れる「重要な背景情報」「参考にした外部リンクやツール」「会議中の補足事項」「印象的な参加者の具体的な発言」などを構造化して記述してください。
4. 【議論内容】: 単なる発言の要約ではなく、「【議題X】（テーマ名）」の下に「何が論点となり」「どのような意見や懸念が出され」「なぜその結論に至ったのか」という意思決定プロセス・議論 of 経緯がわかるように、論理的かつ具体的に記述してください。
5. 【決定事項】: 会議内で最終的に決定・合意された事項（決定した方針やアクションなど）を明確に箇条書きで抽出してください。
6. 【アクションアイテム】: 「担当者」「やること」「期限」を会議データから徹底的に抽出してください。データ内に明記されていない場合も、「誰の担当になりそうか」「期限は次回MTGまでか」などを文脈から合理的に推測・仮設定し、その旨を記載してください。
7. 【次回予定】: 次回の会議の予定日時や、持ち越しとなった課題を抽出してください。
8. 【表現と文体】: 敬体（〜です、〜ます）または常体（〜である、〜だ）の文体を統一し、ビジネスライクかつ簡潔にまとめてください。口語表現、繰り返し、相槌（「えーと」「はい」など）は完全に除外してください。
9. 【ハルシネーションの防止】: 会議データに含まれない全くの虚偽や無関係な事実を作り上げないでください。

--- 制約事項 ---
- 出力はテンプレートの項目のみとし、余計な前置き、挨拶、解説は一切含めないでください。
- テンプレートの「📅 開催日」には、会議データの情報から推測される日付を記入してください。情報がない場合は、メール受信日である「${formattedDate}」を記入してください。
- テンプレートのフォーマット（罫線や記号など）は一字一句変更せずにそのまま使用してください。
- 各項目で該当する情報がない場合は、空欄にするか「なし」と記載してください。

--- テンプレート ---
議事録

📅 開催日: ${formattedDate}
🕐 時間:
📍 場所:
👥 参加者:

────────────────

■ アジェンダ

1.
2.
3.

────────────────

■ 会議中メモ
（議論しながら気づいたこと、参考リンク、雑談で出た発言など、自由に書いてOK）

────────────────

■ 議論内容

【議題1】
・
・

【議題2】
・

────────────────

■ 決定事項

・
・

────────────────

■ アクションアイテム

担当:
やること:
期限:

────────────────

■ 次回予定

・

--- 会議データ ---
${contextText}
`;

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }]
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();
  
  if (responseCode !== 200) {
    throw new Error("Gemini APIエラー (ステータスコード: " + responseCode + "): " + responseBody);
  }
  
  const json = JSON.parse(responseBody);
  
  if (!json.candidates || json.candidates.length === 0 || !json.candidates[0].content || !json.candidates[0].content.parts || json.candidates[0].content.parts.length === 0) {
    throw new Error("Gemini APIからの応答が空、または解析できませんでした。");
  }
  
  return json.candidates[0].content.parts[0].text;
}
