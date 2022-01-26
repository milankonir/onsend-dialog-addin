function onSend(event: Office.AddinCommands.Event) {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/dialog.html",
    { displayInIframe: true, width: 30, height: 30 },
    (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("[Dialog add-in] Dialog opened");

        result.value.addEventHandler(Office.EventType.DialogEventReceived, (args) => {
          if ("error" in args && args.error === 12006) {
            console.log("[Dialog add-in] Dialog closed");
            event.completed({ allowEvent: false });
          }
        });
      } else {
        console.log(`[Dialog add-in] ${JSON.stringify(result.error)}`);
        event.completed({ allowEvent: false });
      }
    }
  );
}

Object.assign(globalThis, {
  onSend,
});

Office.onReady();
