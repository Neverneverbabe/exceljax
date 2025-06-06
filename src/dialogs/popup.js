/* global Office, document, console */

Office.onReady(() => {
  document.getElementById("ok-button").onclick = () => tryCatch(sendStringToParentPage);
});

function sendStringToParentPage() {
  const userName = document.getElementById("name-box").value;
  Office.context.ui.messageParent(userName);
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
