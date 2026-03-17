// 社内ドメイン設定
const INTERNAL_DOMAIN = "fujilogi.co.jp";

// 送信前イベントハンドラ（クラシックOutlook on Windows用）
function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

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

    const allRecipients = [...toList, ...ccList, ...bccList];
    const hasExternal = allRecipients.some(function (r) {
      return r.emailAddress &&
        !r.emailAddress.toLowerCase().endsWith("@" + INTERNAL_DOMAIN.toLowerCase());
    });
    const hasAttachment = attachments.length > 0;

    if (!hasExternal && !hasAttachment) {
      event.completed({ allowEvent: true });
      return;
    }

    const dialogData = {
      to: toList, cc: ccList, bcc: bccList,
      attachments: attachments,
      subject: subject,
      body: body.substring(0, 300),
      hasExternal: hasExternal,
      internalDomain: INTERNAL_DOMAIN
    };

    Office.context.ui.displayDialogAsync(
      "https://dnaiengiadgina.github.io/fujilogi-mailguard/dialog.html",
      { height: 70, width: 50, displayInIframe: false },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          event.completed({ allowEvent: true });
          return;
        }
        const dialog = asyncResult.value;

        dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
          event.completed({ allowEvent: false });
          dialog.close();
        });

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          if (arg.message === "send") {
            event.completed({ allowEvent: true });
          } else {
            event.completed({ allowEvent: false });
          }
          dialog.close();
        });

        dialog.messageChild(JSON.stringify(dialogData));
      }
    );
  }).catch(function () {
    event.completed({ allowEvent: true });
  });
}

function getRecipientsAsync(field) {
  return new Promise(function (resolve) {
    if (!field) { resolve([]); return; }
    field.getAsync(function (r) {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []);
    });
  });
}

function getAttachmentsAsync(item) {
  return new Promise(function (resolve) {
    resolve(item.attachments || []);
  });
}

function getSubjectAsync(item) {
  return new Promise(function (resolve) {
    item.subject.getAsync(function (r) {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : "");
    });
  });
}

function getBodyAsync(item) {
  return new Promise(function (resolve) {
    item.body.getAsync(Office.CoercionType.Text, function (r) {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : "");
    });
  });
}

// イベントハンドラを登録
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
