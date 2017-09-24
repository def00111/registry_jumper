chrome.contextMenus.create({
  title: "Open %s in Registry Editor",
  contexts:["selection"],
  onclick: function(info, tab) {
    browser.runtime.sendNativeMessage('com.def00111.browser.regjump', {
      text: escape(info.selectionText.trim())
    });
  }
});
