// ==UserScript==
// @name binki-sharefile-microsoft-word-recovery
// @version 1.0
// @match https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?*sfwopionline*.sharefile.com*
// @grant GM.deleteValue
// @grant GM.getValue
// @grant GM.setValue
// @require https://github.com/binki/binki-userscript-delay-async/raw/252c301cdbd21eb41fa0227c49cd53dc5a6d1e58/binki-userscript-delay-async.js
// @require https://github.com/binki/binki-userscript-when-element-changed-async/raw/88cf57674ab8fcaa0e86bdf5209342ec7780739a/binki-userscript-when-element-changed-async.js
// ==/UserScript==
(async () => {
  const tenantParamName = 'x-sharefile-tenant';
  const originalUrl = new URL(location.href);
  
  // This happens because the wordeditorframe.aspx URI is handed an access token with
  // a lifetime of a few hours. Without that access token, it can’t do anything. But,
  // if the passed WOPISrc refers to a domain which we support such as sharefile.com,
  // then we can try to reverse map it. Unfortunately, the WOPISrc is not itself tenanted
  // so we have to try to tenantize it.
  const wopiSrc = originalUrl.searchParams.get('WOPISrc');
  if (!wopiSrc) {
    throw new Error('Unable to extract WOPISrc');
  }
  // Check if this is a recognized WOPISrc.
  const sharefileWopiSrcPrefix = 'https://sfwopionline-usw.sharefile.com/WopiServer/wopi/files/';
  if (!wopiSrc.startsWith(sharefileWopiSrcPrefix)) {
    throw new Error(`Unrecognized WOPSrc: ${wopiSrc}`);
  }
  const documentId = wopiSrc.substring(sharefileWopiSrcPrefix.length);
  
  const maybeHostEditUrlJson = [...document.querySelectorAll('script')].map(x => /HostEditUrl:[^'"]*['"]([^'"]*)'/.exec(x.textContent)).filter(x => x).map(x => `"${x[1]}"`)[0];
  if (maybeHostEditUrlJson) {
    // Extract the tenant name from the URL and detect if it is actually sharefile.
    const maybeTenant = (/https:\/\/([^.]*)\.sharefile\.com\//.exec(JSON.parse(maybeHostEditUrlJson)) || [])[1];
    if (maybeTenant) {
      // We got the value! Store it! Put it in the URI so that when we re-navigate to this document it is there.
      const url = new URL(location.href);
      url.searchParams.set(tenantParamName, maybeTenant);
      history.replaceState(history.state, document.title, url.toString());
    }
  } else {
    // See if we are in an error state where we sit at a URI like
    // https://word-edit.officeapps.live.com/we/wordeditorframe.aspx?WOPISrc=https%3A%2F%2Fsfwopionline-usw.sharefile.com%2FWopiServer%2Fwopi%2Ffiles%2Fste77ca4-2ca4-4972-8579-6631f321acc8&IsLicensedUser=1
    // with a message like “Sorry, we ran into a problem.” displaying ( https://i.imgur.com/Okd5kcy.png ).
    while (true) {
      const throttlePromise = delayAsync(200);

      // I do not know what dialogs might show up and I want to avoid text matching in case of localization,
      // so going to go with not-yet-having-found-a-title and having a dialog with a cancel
      // icon.
      if (!document.querySelector('span[data-unique-id=DocumentTitleContent]') && document.querySelector('#WACDialogOuterContainer #WACDialogIconPanel img[class*=CancelRequest]')) {
        const tenant = originalUrl.searchParams.get(tenantParamName);
        if (tenant) {
          document.location.href = `https://${tenant}.sharefile.com/e/${documentId}`;
        } else {
          console.log(`documentId: ${documentId}, no tenant found.`);
          // Add logic to prompt the user about it here?
        }
        break;
      }

      // Wait for this element to appear.
      await whenElementChangedAsync(document.body);

      // Don’t go too fast. If we are reacting to the first event in a while,
      // this may already be resolved.
      await throttlePromise;
    }
  }
})();
