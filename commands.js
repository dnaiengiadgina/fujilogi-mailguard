// 社内ドメイン設定
const INTERNAL_DOMAIN = "fujilogi.co.jp";

Office.initialize = function () {};

// 送信前イベントハンドラ
function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  // 宛先（To / CC / BCC）を取得
  Promise.all([
    getRecipientsAsync(item.to),
    getRecipientsAsync(item.cc),
    getRecipientsAsync(item.bcc),
    getAttachmentsAsync(item),
    getSubjectAsync(item),
    getBodyAsync(item)
  ]).then(function (results) {
    const toList      = results[0];
    const ccList      = results[1];
    const bccList     = results[2];
    const attachments = results[3];
    const subject     = results[4];
    const body        = results[5];

    // 外部ドメイン判定
    const allRecipients = [...toList, ...ccList, ...bccList];
    const hasExternal = allRecipients.some(function (r) {
      return r.emailAddress &&
        !r.emailAddress.toLowerCase().endsWith("@" + INTERNAL_DOMAIN.toLowerCase());
    });
    const hasAttachment = attachments.length > 0;

    // トリガー条件：外部宛先 OR 添付ファイルあり
    if (!hasExternal && !hasAttachment) {
      event.completed({ allowEvent: true });
      return;
    }

    // ダイアログに渡すデータを作成
    const dialogData = {
      to:          toList,
      cc:          ccList,
      bcc:         bccList,
      attachments: attachments,
      subject:     subject,
      body:        body.substring(0, 300),
      hasExternal: hasExternal,
      internalDomain: INTERNAL_DOMAIN
    };

    // ダイアログを開く
    const dialogUrl = Office.context.mailbox.ewsUrl
      ? window.location.origin + "/dialog.html"
      : "https://YOUR_GITHUB_USERNAME.github.io/fujilogi-mailguard/dialog.html";

    Office.context.ui.displayDialogAsync(
      "https://YOUR_GITHUB_USERNAME.github.io/fujilogi-mailguard/dialog.html",
      { height: 70, width: 50, displayInIframe: false },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          event.completed({ allowEvent: true });
          return;
        }

        const dialog = asyncResult.value;

        // ダイアログが開いたらデータを送信
        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          function (arg) {
            // ダイアログが閉じられた（キャンセル）
            event.completed({ allowEvent: false });
            dialog.close();
          }
        );

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          function (arg) {
            const message = arg.message;
            if (message === "send") {
              event.completed({ allowEvent: true });
            } else {
              event.completed({ allowEvent: false });
            }
            dialog.close();
          }
        );

        // データをダイアログに送信
        dialog.messageChild(JSON.stringify(dialogData));
      }
    );
  }).catch(function (err) {
    console.error("Error:", err);
    event.completed({ allowEvent: true });
  });
}

// 宛先取得ヘルパー
function getRecipientsAsync(recipientField) {
  return new Promise(function (resolve) {
    if (!recipientField) { resolve([]); return; }
    recipientField.getAsync(function (result) {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : []);
    });
  });
}

// 添付ファイル取得ヘルパー
function getAttachmentsAsync(item) {
  return new Promise(function (resolve) {
    if (!item.attachments) { resolve([]); return; }
    resolve(item.attachments);
  });
}

// 件名取得ヘルパー
function getSubjectAsync(item) {
  return new Promise(function (resolve) {
    item.subject.getAsync(function (result) {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : "");
    });
  });
}

// 本文取得ヘルパー
function getBodyAsync(item) {
  return new Promise(function (resolve) {
    item.body.getAsync(Office.CoercionType.Text, function (result) {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : "");
    });
  });
}

// Office.js へ関数を登録
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
