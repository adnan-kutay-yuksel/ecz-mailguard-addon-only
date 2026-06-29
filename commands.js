// MailGuard — On-Send Handler
const API_BASE = "https://unworthy-dreamily-calculus.ngrok-free.dev";
Office.onReady(function () {
  Office.actions.associate("onItemSend", onItemSend);
});
Office.initialize = function () {};

function offlineOnayDialogGoster(event) {
  var mesaj = "MailGuard şu anda devrede değil.";
  var soru  = "Mailinizi göndermeden önce kontrol ediniz.";
  var dialogUrl = "https://adnan-kutay-yuksel.github.io/ecz-mailguard-addon-only/confirm.html"
    + "?mesaj=" + encodeURIComponent(mesaj)
    + "&soru="  + encodeURIComponent(soru);
  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 30, width: 40, promptBeforeOpen: false },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        Office.context.mailbox.item.notificationMessages.replaceAsync("mailguard_offline", {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: mesaj + " (Onay penceresi açılamadı.)"
        });
        event.completed({ allowEvent: true });
        return;
      }
      var dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        dialog.close();
        if (arg.message === "GONDER") {
          event.completed({ allowEvent: true });
        } else {
          event.completed({ allowEvent: false });
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
        event.completed({ allowEvent: false });
      });
    }
  );
}

function onItemSend(event) {
  var item     = Office.context.mailbox.item;
  var gonderen = Office.context.mailbox.userProfile.emailAddress || "bilinmiyor";
  item.subject.getAsync(function (subjectResult) {
    var konu = subjectResult.value || "";
    item.body.getAsync(Office.CoercionType.Text, function (bodyResult) {
      var icerik = bodyResult.value || "";
      fetch(API_BASE + "/token", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ kullanici: gonderen })
      })
      .then(function (res) { return res.json(); })
      .then(function (tokenData) {
        var jwt = tokenData.token;
        analiz({ konu: konu, icerik: icerik, gonderen: gonderen }, jwt, event);
      })
      .catch(function (err) {
        console.error("MailGuard token hatasi:", err);
        offlineOnayDialogGoster(event);
      });
    });
  });
}

function analiz(payload, jwt, event) {
  fetch(API_BASE + "/analiz", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + jwt
    },
    body: JSON.stringify(payload)
  })
  .then(function (res) { return res.json(); })
  .then(function (data) {
    var aktifMod = data.aktif_mod || "ZORUNLU";
    var karar;
    try {
      karar = typeof data.karar === "string" ? JSON.parse(data.karar) : data.karar;
    } catch (e) {
      event.completed({ allowEvent: true });
      return;
    }
    var karantina = karar.karar === "KARANTİNA";
    if (!karantina) {
      event.completed({ allowEvent: true });
      return;
    }
    var sabitMesaj = "MailGuard yazışmanızı kontrol etmenizi öneriyor.";
    if (aktifMod === "SERBEST") {
      event.completed({ allowEvent: true });
    } else if (aktifMod === "KONTROLLU") {
      var dialogUrl = "https://adnan-kutay-yuksel.github.io/ecz-mailguard-addon-only/confirm.html"
        + "?mesaj=" + encodeURIComponent(sabitMesaj)
        + "&soru="  + encodeURIComponent("Yine de göndermek istiyor musunuz?");
      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 30, width: 40, promptBeforeOpen: false },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            Office.context.mailbox.item.notificationMessages.replaceAsync("mailguard_1", {
              type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
              message: sabitMesaj + " | Onay için Outlook uygulamasını kullanın."
            });
            event.completed({ allowEvent: false });
            return;
          }
          var dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
            dialog.close();
            if (arg.message === "GONDER") {
              event.completed({ allowEvent: true });
            } else {
              event.completed({ allowEvent: false });
            }
          });
          dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
            event.completed({ allowEvent: false });
          });
        }
      );
    } else {
      Office.context.mailbox.item.notificationMessages.replaceAsync("mailguard_1", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: sabitMesaj
      });
      event.completed({ allowEvent: false });
    }
  })
  .catch(function (err) {
    console.error("MailGuard API hatasi:", err);
    offlineOnayDialogGoster(event);
  });
}