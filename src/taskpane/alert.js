import "core-js";
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    Office.context.ui.messageParent("alert");
  }
});
