chrome.action.onClicked.addListener(async (tab) => {
  if (!tab.id) {
    return;
  }

  try {
    await chrome.tabs.sendMessage(tab.id, { type: "teams-export/toggle-panel" });
  } catch (error) {
    console.warn("Unable to toggle Teams Selected Messages Export panel.", error);
  }
});
