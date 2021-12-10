const onOpen = (e) => {
  Logger.log(`onOpen e.authmode: ${e.authMode}`);
  const ui = SpreadsheetApp.getUi();
  let menu = ui.createAddonMenu();

  menu.addItem("Hello world", "hello");

  /*
  menu.addSubMenu(
    ui
      .createMenu("Developer")
      .addItem("Execution logs", "openExecutionLogs")
      .addSeparator()
      .addSubMenu(
        ui
          .createMenu("DeveloperMetaData")
          .addItem("Set DeveloperMetadata", "setDeveloperMetadataUI")
          .addItem("Remove DeveloperMetadata", "removeDeveloperMetadataUI")
          .addItem("Log DevelperMetadata", "logDeveloperMetadata")
      )
  );
  */
  menu.addToUi();
};

const onInstall = (e) => {
  onOpen(e);
};
