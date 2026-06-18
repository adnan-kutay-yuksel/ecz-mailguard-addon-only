// MailGuard — On-Send Handler
const API_BASE = "https://unworthy-dreamily-calculus.ngrok-free.dev";

Office.onReady(function () {
  Office.actions.associate("onItemSend", onItemSend);
});

Office.initialize = function () {};

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
        Office.context.mailbox.item.notificationMessages.replaceAsync("mailguard_offline", {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: "MailGuard erisilemuyor. DLP politikalari gecerlidir. Sorumlusunuz."
        }, function () {
          event.completed({ allowEvent: true });
        });
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

    var aciklama = (karar.aciklama || "").substring(0, 60);
    var mesaj = "Risk Skoru: " + (karar.skor || "?") + "/10 | " + aciklama;

    if (aktifMod === "SERBEST") {
      event.completed({ allowEvent: true });

    } else if (aktifMod === "KONTROLLU") {
      Office.context.mailbox.item.notificationMessages.replaceAsync("mailguard_1", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: mesaj,
        icon: "Icon.16x16",
        persistent: true
      });
      event.completed({ allowEvent: true });

    } else {
      Office.context.mailbox.item.notificationMessages.replaceAsync("mailguard_1", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: mesaj,
        icon: "Icon.16x16",
        persistent: true
      });
      event.completed({ allowEvent: false });
    }
  })
  .catch(function (err) {
    console.error("MailGuard API hatasi:", err);
    Office.context.mailbox.item.notificationMessages.replaceAsync("mailguard_offline", {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: "MailGuard erisilemuyor. DLP politikalari gecerlidir. Sorumlusunuz."
    }, function () {
      event.completed({ allowEvent: true });
    });
  });
}
