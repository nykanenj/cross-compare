const showPopup = (popupObj) => {
  const { title, messages, file, folder, modal, height, width } = popupObj;

  const spaces = `<br><br><br>`;

  let messageHtml = !messages.length
    ? ""
    : messages.reduce((text, msg) => {
        const appendable = msg
          ? `<div>
       ${msg}
     <div>`
          : "";
        return `${text}${appendable}`;
      }, "");
  messageHtml = `${messageHtml}${spaces}`;

  const fileHtml = file
    ? `<a href="${file.getUrl()}" target="blank" style="font-size: 20px; color: #fff; background-color: #343a40; border-color: #343a40; padding: 0.7rem 1.1rem; border-radius: 0.7rem; text-decoration:none">
         Avaa tiedosto:
         ${file.getName()}
       </a>
       ${spaces}`
    : "";

  const folderHtml = folder
    ? `<a href="${folder.getUrl()}" target="blank" style="font-size: 20px; color: #fff; background-color: #343a40; border-color: #343a40; padding: 0.7rem 1.1rem; border-radius: 0.7rem; text-decoration:none">
         Avaa kansio:
         ${folder.getName()}
       </a>
       ${spaces}`
    : "";

  const html = `
    <html>
      <body>
        ${messageHtml}
        ${fileHtml}
        ${folderHtml}
      </body>
    </html>`;

  let ui = HtmlService.createHtmlOutput(html)
    .setHeight(height || 150)
    .setWidth(width || 1300);

  if (modal) {
    SpreadsheetApp.getUi().showModalDialog(ui, title);
  } else {
    SpreadsheetApp.getUi().showModelessDialog(ui, title);
  }
};
