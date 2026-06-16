// MailGuard — On-Send Handler

const API_BASE  = "https://unworthy-dreamily-calculus.ngrok-free.dev";
const JWT_TOKEN = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJrdWxsYW5pY2kiOiJ0ZXN0In0.nzZfgN5MVkyL3Sdr_0hJRVl8FavUkQ1R44pOHNhOUtQ";

Office.initialize = function () {};

function onItemSend(event) {
  var item     = Office.context.mailbox.item;
  var gonderen = Office.context.mailbox.userProfile.emailAddress || "";

  item.subject.getAsync(function (subjectResult) {
    var konu = subjectResult.value || "";

    item.body.getAsync(Office.CoercionType.Text, function (bodyResult) {
      var icerik = bodyResult.value || "";
      analiz({ konu: konu, icerik: icerik, gonderen: gonderen }, event);
    });
  });
}

function analiz(payload, event) {
  fetch(API_BASE + "/analiz", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + JWT_TOKEN
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
    var mesaj1 = "Mailiniz Rekabet Hukuku Mevzuati kapsaminda uygun degil.";
    var mesaj2 = "Risk Skoru: " + (karar.skor || "?") + "/10 | " + aciklama;

    if (aktifMod === "SERBEST") {
      event.completed({ allowEvent: true });

    } else if (aktifMod === "KONTROLLU") {
      Office.context.mailbox.item.notificationMessages.replaceAsync("mailguard_1", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: mesaj1,
        icon: "Icon.16x16",
        persistent: true
      });
      Office.context.mailbox.item.notificationMessages.replaceAsync("mailguard_2", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: mesaj2,
        icon: "Icon.16x16",
        persistent: true
      });
      event.completed({ allowEvent: true });

    } else {
      Office.context.mailbox.item.notificationMessages.replaceAsync("mailguard_1", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: mesaj1
      });
      Office.context.mailbox.item.notificationMessages.replaceAsync("mailguard_2", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: mesaj2,
        icon: "Icon.16x16",
        persistent: true
      });
      event.completed({ allowEvent: false });
    }
  })
  .catch(function (err) {
    console.error("MailGuard API hatasi:", err);
    event.completed({ allowEvent: true });
  });
}

Office.onReady(function () {
  Office.actions.associate("onItemSend", onItemSend);
});
