# JP Signature: Outlook add-in

Inserts your signature automatically in Outlook for Mac. It also works in Outlook on the web and on Windows, because add-ins follow your mailbox.

| Situation | Signature |
|---|---|
| New message | **default** |
| Reply / reply all / forward | **External** |
| Any To/Cc/Bcc recipient outside `@adem.alabama.gov` | switches to **External** |
| New message where every outside recipient is removed again | switches back to **default** |
| Plain-text message | text version of the same signature (no logo) |

All text is Public Sans. Aptos is only the fallback. The ADEM logo is embedded in each message as an inline image, so recipients don't get a "download pictures" prompt.

---

## Setup (about 15 minutes, done once)

### 1. Install Public Sans on your Mac
1. Download it from Google Fonts (fonts.google.com, search "Public Sans") and unzip it.
2. Open the `static` folder, select **every** `.ttf` file (Black is included), and double-click to install them with Font Book.
3. Quit and reopen Outlook.

Without the font installed, you'll see Aptos while writing. Recipients who don't have Public Sans will also see Aptos (or their system sans-serif).

### 2. Host the files (free, using GitHub Pages)
Outlook loads add-in code from an HTTPS address, so the files need to live online.
1. On github.com, create a **public** repository named exactly **`jp-signature-addin`**.
2. Upload everything in this folder **except** `test/` (it's harmless if you include it).
3. Go to repo **Settings → Pages → Build and deployment** and set Source to **Deploy from a branch**, branch `main`, folder `/ (root)`. Save.
4. After about a minute, check that `https://<your-username>.github.io/jp-signature-addin/index.html` loads.

The hosted files contain only the contact details already in your signature.

### 3. Point the manifest at your host
In Terminal, inside this folder:
```sh
chmod +x configure.sh
./configure.sh <your-username>.github.io
```
Alternatively, open `manifest.xml` and replace every `YOUR-HOST` with `<your-username>.github.io`. Upload the updated `manifest.xml` to the repo too, so the hosted copy matches.

### 4. Turn off Outlook's own signatures (avoids a brief flash or swap)
In Outlook for Mac, go to **Settings → Signatures** and set the default for both **New messages** and **Replies/forwards** to **None**. The add-in replaces whatever is in the signature spot anyway, so this just makes things cleaner.

### 5. Install (sideload) the add-in
1. In a browser, go to **https://aka.ms/olksideload** and sign in with your ADEM account. Outlook on the web opens and the **Add-ins for Outlook** dialog appears.
2. Choose **My add-ins** → **Custom Addins** → **Add a custom add-in** → **Add from File…**, then pick `manifest.xml`.
3. Restart Outlook for Mac.

If there's no **custom add-in** option, ADEM IT has blocked sideloading for your account. In that case they'd have to deploy the add-in for you (they can use this same manifest).

### 6. Test it
- **New message**: the default signature appears.
- **Add a Gmail address**: it switches to External. Remove the Gmail address and it switches back.
- **Reply to anything**: External.

---

## Changing the signature later
Edit the text in `src/runtime.js` (the `buildHtml` and `buildText` functions), then upload the file to the repo again. Outlook caches add-in files, so restart Outlook if the change doesn't show up right away. You only need to bump `<Version>` and reinstall when `manifest.xml` itself changes.

## Files
- `manifest.xml`: tells Outlook which events to listen for and where the code lives (XML format, required for Mac)
- `src/runtime.html`: background page Outlook loads on Mac, web, and new Outlook for Windows
- `src/runtime.js`: all the logic plus the two signatures and the embedded logo
- `index.html`: small info page (required by the manifest)
- `assets/`: add-in icons and the original logo PNG
- `test/runtime.test.js`: local tests (`node test/runtime.test.js`)

Adapted from Microsoft's "outlook-set-signature" sample (MIT License).
