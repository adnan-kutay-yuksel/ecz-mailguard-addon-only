const API_BASE = "https://unworthy-dreamily-calculus.ngrok-free.dev";

Office.onReady(function () {
  Office.actions.associate("onItemSend", onItemSend);
});

Office.initialize = function () {};

function fetchWithTimeout(url, options, ms) {
  return new Promise(function (resolve, reject) {
    var timer = setTimeout(function () {
      reject(new Error("timeout"));
    }, ms);
    fetch(url, options).then(function (res) {
      clearTimeout(timer);
      resolve(res);
    }).catch(function (err) {
      clearTimeout(timer);
      reject(err);
    });
  });
}

function onItemSend(event) {
  var item     = Office.context.mailbox.item;
  var gonderen = Office.context.mailbox.userProfile.emailAddress || "bilinmiyor";

  item.subject.getAsync(function (subjectResult) {
    var konu = subjectResult.value || "";

    item.body.getAsync(Office.CoercionType.Text, function (bodyResult) {
      var icerik = bodyResult.value || "";

      fetchWithTimeout(API_BASE + "/token", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ kullanici: gonderen })
      }, 8000)
      .then(function (res) { return res.json(); })
      .then(function (tokenData) {
        var jwt = tokenData.token;
        analiz({ konu: konu, icerik: icerik, gonderen: gonderen }, jwt, event);
      })
      .catch(function (err) {
        console.error("CATCH TETIKLENDI token:", err.message);
        gosterFallbackUyari(event);
      });
    });
  });
}

function analiz(payload, jwt, event) {
  fetchWithTimeout(API_BASE + "/analiz", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + jwt
    },
    body: JSON.stringify(payload)
  }, 20000)
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
    var mesaj = "Yazışmanı kontrol etmeni öneririz! Risk skoru: " + (karar.skor || "?") + "/10 - " + aciklama;

    if (aktifMod === "SERBEST") {
      event.completed({ allowEvent: true });

    } else if (aktifMod === "KONTROLLU") {
      var dialogUrl = "https://adnan-kutay-yuksel.github.io/ecz-mailguard-addon-only/confirm.html"
        + "?mesaj=" + encodeURIComponent(mesaj);

      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 30, width: 40, promptBeforeOpen: false },
        function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            Office.context.mailbox.item.notificationMessages.replaceAsync("mailguard_1", {
              type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
              message: mesaj + " | Onay icin Outlook uygulamasini kullanin."
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
        message: mesaj
      });
      event.completed({ allowEvent: false });
    }
  })
  .catch(function (err) {
    console.error("CATCH TETIKLENDI analiz:", err.message);
    gosterFallbackUyari(event);
  });
}

function gosterFallbackUyari(event) {
  console.log("FALLBACK TETIKLENDI");
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "mailguard_offline",
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: "MailGuard erisilemuyor. DLP politikalari gecerlidir. Sorumlusunuz."
    },
    function (result) {
      console.log("notification result:", JSON.stringify(result));
      event.completed({ allowEvent: true });
    }
  );
}
