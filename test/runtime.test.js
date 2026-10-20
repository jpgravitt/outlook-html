// Local tests: runs runtime.js against a mocked Office.js. Usage: node test/runtime.test.js
"use strict";
const assert = require("assert");
const path = require("path");

function makeMock({ composeType = "newMail", coercion = "html", to = [], cc = [], bcc = [], withSessionData = true } = {}) {
  const log = { signatures: [], attachments: [], completed: 0 };
  const session = {};
  const field = (arr) => ({ getAsync: (cb) => setImmediate(() => cb({ status: "succeeded", value: arr.map((e) => ({ emailAddress: e })) })) });
  const item = {
    getComposeTypeAsync: (cb) => setImmediate(() => cb({ status: "succeeded", value: { composeType, coercionType: coercion } })),
    to: field(to), cc: field(cc), bcc: field(bcc),
    body: {
      setSignatureAsync: (data, opts, cb) => { log.signatures.push({ data, type: opts.coercionType }); setImmediate(() => cb({ status: "succeeded" })); },
    },
    addFileAttachmentFromBase64Async: (b64, name, opts, cb) => { log.attachments.push({ name, inline: opts.isInline, len: b64.length }); setImmediate(() => cb({ status: "succeeded", value: "att1" })); },
  };
  if (withSessionData) {
    item.sessionData = {
      getAsync: (k, cb) => setImmediate(() => cb(k in session ? { status: "succeeded", value: session[k] } : { status: "failed", value: null })),
      setAsync: (k, v, cb) => { session[k] = v; setImmediate(() => cb({ status: "succeeded" })); },
    };
  }
  global.Office = {
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    CoercionType: { Html: "html", Text: "text" },
    context: { mailbox: { item } },
    actions: { associate: () => {} },
  };
  const setRecipients = (t, c = [], b = []) => { item.to = field(t); item.cc = field(c); item.bcc = field(b); };
  return { log, setRecipients };
}

function fire(handler) {
  return new Promise((resolve) => handler({ completed: () => resolve() }));
}

function load() {
  const p = path.join(__dirname, "..", "src", "runtime.js");
  delete require.cache[require.resolve(p)];
  return require(p);
}

const lastSigName = (log) => {
  const d = log.signatures[log.signatures.length - 1].data;
  return d.indexOf("questions or concerns") >= 0 ? "external" : "default";
};

(async () => {
  let passed = 0;
  const t = async (name, fn) => { await fn(); passed++; console.log("  ok -", name); };

  await t("domain check", async () => {
    const m = makeMock(); const rt = load();
    assert.strictEqual(rt.isExternalAddress("Someone@ADEM.alabama.gov"), false);
    assert.strictEqual(rt.isExternalAddress("x@sub.adem.alabama.gov"), false);
    assert.strictEqual(rt.isExternalAddress("x@gmail.com"), true);
    assert.strictEqual(rt.isExternalAddress("x@notadem.alabama.gov"), true);
    assert.strictEqual(rt.isExternalAddress(""), false);
  });

  await t("new message, no recipients -> default + one inline logo", async () => {
    const m = makeMock(); const rt = load();
    await fire(rt.onNewMessageComposeHandler);
    assert.strictEqual(m.log.signatures.length, 1);
    assert.strictEqual(lastSigName(m.log), "default");
    assert.strictEqual(m.log.attachments.length, 1);
    assert.strictEqual(m.log.attachments[0].inline, true);
    assert.ok(m.log.signatures[0].data.indexOf("cid:" + m.log.attachments[0].name) >= 0, "img references the attached logo");
    assert.ok(m.log.signatures[0].data.length < 30000, "under the 30,000-character limit");
  });

  await t("reply -> External", async () => {
    const m = makeMock({ composeType: "reply", to: ["boss@adem.alabama.gov"] }); const rt = load();
    await fire(rt.onNewMessageComposeHandler);
    assert.strictEqual(lastSigName(m.log), "external");
  });

  await t("forward -> External", async () => {
    const m = makeMock({ composeType: "forward" }); const rt = load();
    await fire(rt.onNewMessageComposeHandler);
    assert.strictEqual(lastSigName(m.log), "external");
  });

  await t("new message: add outside recipient -> External, remove -> default, logo attached once", async () => {
    const m = makeMock(); const rt = load();
    await fire(rt.onNewMessageComposeHandler);
    m.setRecipients(["coworker@adem.alabama.gov"]);
    await fire(rt.onMessageRecipientsChangedHandler);
    assert.strictEqual(m.log.signatures.length, 1, "internal recipient: no change");
    m.setRecipients(["coworker@adem.alabama.gov"], ["plant.manager@example.com"]);
    await fire(rt.onMessageRecipientsChangedHandler);
    assert.strictEqual(m.log.signatures.length, 2);
    assert.strictEqual(lastSigName(m.log), "external");
    m.setRecipients(["coworker@adem.alabama.gov"]);
    await fire(rt.onMessageRecipientsChangedHandler);
    assert.strictEqual(lastSigName(m.log), "default");
    assert.strictEqual(m.log.attachments.length, 1, "logo not duplicated");
  });

  await t("reply stays External when recipients change", async () => {
    const m = makeMock({ composeType: "reply", to: ["x@example.com"] }); const rt = load();
    await fire(rt.onNewMessageComposeHandler);
    m.setRecipients(["coworker@adem.alabama.gov"]);
    await fire(rt.onMessageRecipientsChangedHandler);
    assert.strictEqual(m.log.signatures.length, 1);
    assert.strictEqual(lastSigName(m.log), "external");
  });

  await t("plain-text message -> text signature, no logo", async () => {
    const m = makeMock({ coercion: "text" }); const rt = load();
    await fire(rt.onNewMessageComposeHandler);
    assert.strictEqual(m.log.signatures[0].type, "text");
    assert.ok(m.log.signatures[0].data.indexOf("<") < 0);
    assert.strictEqual(m.log.attachments.length, 0);
  });

  await t("no sessionData support -> recipient switching is skipped safely", async () => {
    const m = makeMock({ withSessionData: false }); const rt = load();
    await fire(rt.onNewMessageComposeHandler);
    m.setRecipients(["x@example.com"]);
    await fire(rt.onMessageRecipientsChangedHandler);
    assert.strictEqual(m.log.signatures.length, 1);
    assert.strictEqual(m.log.attachments.length, 1);
  });

  await t("fonts: Public Sans first, Aptos only as fallback", async () => {
    makeMock(); const rt = load();
    const html = rt.buildHtml("external", "logo.png");
    const families = html.match(/font-family:[^;]+/g);
    assert.ok(families.length > 0);
    families.forEach((f) => {
      assert.ok(/^font-family:'Public Sans/.test(f), "starts with Public Sans: " + f);
      assert.ok(f.indexOf("Aptos") > f.indexOf("Public Sans"), "Aptos after Public Sans");
    });
  });

  console.log(`\n${passed} tests passed`);
})().catch((e) => { console.error("FAIL:", e.message); process.exit(1); });
