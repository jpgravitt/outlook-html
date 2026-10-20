/*
 * JP Gravitt signature add-in: event-based runtime
 * -------------------------------------------------
 * Adapted from Microsoft's "outlook-set-signature" sample (MIT License).
 * Written in plain ES5 so the same file also runs in the JavaScript-only
 * runtime used by classic Outlook on Windows.
 *
 * What it does:
 *   New message          -> "default" signature
 *   Reply / forward      -> "External" signature
 *   Any recipient outside @adem.alabama.gov (To/Cc/Bcc) -> "External"
 *   (On a new message, removing every outside recipient switches back to "default".)
 */

var INTERNAL_DOMAIN = "adem.alabama.gov";
var STATE_KEY = "jpSigState";

/* ---------- Signature content ---------- */

// Public Sans throughout; Aptos only as the fallback.
var FONT_BODY = "'Public Sans',Aptos,sans-serif";
var FONT_NAME = "'Public Sans Black','Public Sans',Aptos,sans-serif";

var CSS_BODY = "margin:0;font-family:" + FONT_BODY + ";font-size:11pt;color:#000000;";
var CSS_NAME = "margin:0;font-family:" + FONT_NAME + ";font-size:15pt;font-variant:small-caps;color:#000000;";
var CSS_SPACER = "margin:0;font-family:" + FONT_BODY + ";font-size:2pt;line-height:2pt;color:#000000;";
var CSS_EMAIL = "color:blue;text-decoration:underline;";
var CSS_LOGO_LINK = "color:#000000;text-decoration:none;";
var CSS_LOGO_IMG = "border:0;width:78px;height:23px;";

// Microsoft's docs say setSignatureAsync doesn't support inline CSS, so the
// styling lives in an internal <style> block. Inline styles are kept too as a
// harmless backup for clients that drop <style>.
var STYLE_BLOCK =
  "<style>" +
  ".jpsig-line{" + CSS_BODY + "}" +
  ".jpsig-name{" + CSS_NAME + "}" +
  ".jpsig-spacer{" + CSS_SPACER + "}" +
  ".jpsig-email{" + CSS_EMAIL + "}" +
  ".jpsig-logolink{" + CSS_LOGO_LINK + "}" +
  ".jpsig-logo{" + CSS_LOGO_IMG + "}" +
  "</style>";

function line(html) {
  return "<div class=\"jpsig-line\" style=\"" + CSS_BODY + "\">" + html + "</div>";
}

function buildHtml(name, logoName) {
  var h = STYLE_BLOCK + "<div class=\"jpsig\">";
  h += line("&nbsp;");
  if (name === "external") {
    h += line("If you have any questions or concerns, feel free to contact me via email or phone (listed below).");
    h += "<div class=\"jpsig-spacer\" style=\"" + CSS_SPACER + "\">&nbsp;</div>";
  }
  h += "<div class=\"jpsig-name\" style=\"" + CSS_NAME + "\">JP Gravitt, E.I.</div>";
  h += line("Environmental Engineer, Sr.");
  h += line("Industrial Minerals Section");
  h += line("Air Division &ndash; Energy Branch");
  h += line("Alabama Department of Environmental Management");
  h += line("<a class=\"jpsig-email\" style=\"" + CSS_EMAIL + "\" href=\"mailto:John.Gravitt@adem.alabama.gov\">John.Gravitt@adem.alabama.gov</a>");
  h += line("(334) 271-7952");
  h += line(
    "<a class=\"jpsig-logolink\" style=\"" + CSS_LOGO_LINK + "\" href=\"https://adem.alabama.gov/\">" +
      "<img class=\"jpsig-logo\" style=\"" + CSS_LOGO_IMG + "\" src=\"cid:" + logoName + "\" alt=\"ADEM\" width=\"78\" height=\"23\" border=\"0\">" +
    "</a>"
  );
  h += "</div>";
  h += line("&nbsp;");
  return h;
}

function buildText(name) {
  var t = "\n";
  if (name === "external") {
    t += "If you have any questions or concerns, feel free to contact me via email or phone (listed below).\n\n";
  }
  t += "JP GRAVITT, E.I.\n" +
    "Environmental Engineer, Sr.\n" +
    "Industrial Minerals Section\n" +
    "Air Division – Energy Branch\n" +
    "Alabama Department of Environmental Management\n" +
    "John.Gravitt@adem.alabama.gov\n" +
    "(334) 271-7952\n";
  return t;
}

// ADEM logo, 78x23 PNG, taken from the Windows signature folder.
var LOGO_BASE64 = "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";

/* ---------- Session state (survives between events for the same draft) ---------- */

function hasSessionData() {
  var item = Office.context.mailbox.item;
  return !!(item && item.sessionData && item.sessionData.getAsync);
}

function loadState(cb) {
  if (!hasSessionData()) { cb(null); return; }
  Office.context.mailbox.item.sessionData.getAsync(STATE_KEY, function (r) {
    var s = null;
    if (r.status === Office.AsyncResultStatus.Succeeded && r.value) {
      try { s = JSON.parse(r.value); } catch (e) { s = null; }
    }
    cb(s);
  });
}

function saveState(state, cb) {
  if (!hasSessionData()) { cb(); return; }
  Office.context.mailbox.item.sessionData.setAsync(STATE_KEY, JSON.stringify(state), function () { cb(); });
}

/* ---------- Helpers ---------- */

function getComposeInfo(cb) {
  var item = Office.context.mailbox.item;
  if (!item.getComposeTypeAsync) { cb({ composeType: "newMail", coercion: "html" }); return; }
  item.getComposeTypeAsync(function (r) {
    var info = { composeType: "newMail", coercion: "html" };
    if (r.status === Office.AsyncResultStatus.Succeeded && r.value) {
      info.composeType = r.value.composeType || "newMail";
      info.coercion = r.value.coercionType || "html";
    }
    cb(info);
  });
}

function isExternalAddress(address) {
  if (!address) return false; // unresolved entry: ignore until it resolves
  var a = String(address).toLowerCase().replace(/^\s+|\s+$/g, "");
  var at = a.lastIndexOf("@");
  if (at < 0) return false;
  var domain = a.substring(at + 1);
  return !(domain === INTERNAL_DOMAIN || domain.slice(-(INTERNAL_DOMAIN.length + 1)) === "." + INTERNAL_DOMAIN);
}

function hasExternalRecipient(cb) {
  var item = Office.context.mailbox.item;
  var fields = [item.to, item.cc, item.bcc];
  var pending = 0, external = false, finished = false;
  function doneOne() {
    pending--;
    if (pending === 0 && !finished) { finished = true; cb(external); }
  }
  for (var i = 0; i < fields.length; i++) {
    if (fields[i] && fields[i].getAsync) pending++;
  }
  if (pending === 0) { cb(false); return; }
  for (var j = 0; j < fields.length; j++) {
    if (!(fields[j] && fields[j].getAsync)) continue;
    fields[j].getAsync(function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded && r.value) {
        for (var k = 0; k < r.value.length; k++) {
          if (isExternalAddress(r.value[k].emailAddress)) external = true;
        }
      }
      doneOne();
    });
  }
}

function newLogoName() {
  return "adem-logo-" + new Date().getTime().toString(36) + ".png";
}

function ensureLogo(state, cb) {
  if (state.logoName) { cb(true); return; }
  var name = newLogoName();
  Office.context.mailbox.item.addFileAttachmentFromBase64Async(LOGO_BASE64, name, { isInline: true }, function (r) {
    if (r.status === Office.AsyncResultStatus.Succeeded) {
      state.logoName = name;
      cb(true);
    } else {
      cb(false);
    }
  });
}

function applySignature(state, name, cb) {
  var body = Office.context.mailbox.item.body;
  if (state.coercion === "text") {
    body.setSignatureAsync(buildText(name), { coercionType: Office.CoercionType.Text }, function (r) {
      cb(r.status === Office.AsyncResultStatus.Succeeded);
    });
    return;
  }
  ensureLogo(state, function () {
    // If the logo upload failed, the signature still goes in; only the image is missing.
    body.setSignatureAsync(buildHtml(name, state.logoName || "adem-logo.png"), { coercionType: Office.CoercionType.Html }, function (r) {
      cb(r.status === Office.AsyncResultStatus.Succeeded);
    });
  });
}

/* ---------- Core decision ---------- */

function evaluate(event, isNewCompose) {
  var completed = false;
  function done() {
    if (completed) return;
    completed = true;
    try { event.completed(); } catch (e) { /* ignore */ }
  }
  try {
    loadState(function (state) {
      function decide(st) {
        function finish(desired) {
          if (st.current === desired) { done(); return; }
          applySignature(st, desired, function (ok) {
            if (ok) st.current = desired;
            saveState(st, done);
          });
        }
        if (st.composeType === "reply" || st.composeType === "forward") {
          finish("external");
        } else {
          hasExternalRecipient(function (ext) { finish(ext ? "external" : "default"); });
        }
      }
      if (state && !isNewCompose) {
        decide(state);
      } else {
        getComposeInfo(function (info) {
          decide({ composeType: info.composeType, coercion: info.coercion, current: null, logoName: null });
        });
      }
    });
  } catch (e) {
    done();
  }
}

/* ---------- Event handlers (names must match manifest.xml) ---------- */

function onNewMessageComposeHandler(event) {
  evaluate(event, true);
}

function onMessageRecipientsChangedHandler(event) {
  // Without sessionData the add-in can't tell what's already inserted, so it skips
  // switching rather than risk attaching duplicate logos.
  if (!hasSessionData()) { event.completed(); return; }
  evaluate(event, false);
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
}

// Exposed for local tests only (no effect inside Outlook).
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    buildHtml: buildHtml, buildText: buildText, isExternalAddress: isExternalAddress,
    onNewMessageComposeHandler: onNewMessageComposeHandler,
    onMessageRecipientsChangedHandler: onMessageRecipientsChangedHandler
  };
}
