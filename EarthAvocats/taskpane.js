Office.onReady(() => {
  loadTemplates();
});

async function loadTemplates() {
  const container = document.getElementById("templates");

  const sharepointUrl = "https://earthavocats.sharepoint.com/ModelesMails/_api/web/GetFolderByServerRelativeUrl('ModelesMails')/Files";

  const resp = await fetch(sharepointUrl, {
    headers: { "accept": "application/json;odata=verbose" }
  });

  const data = await resp.json();
  const files = data.d.results;

  files.forEach(file => {
    const button = document.createElement("button");
    button.textContent = file.Name.replace(".html", "");
    button.onclick = () => insertTemplate(file.ServerRelativeUrl);
    container.appendChild(button);
  });
}

async function insertTemplate(relativeUrl) {
  const url = "https://earthavocats.sharepoint.com" + relativeUrl;
  const resp = await fetch(url);
  const html = await resp.text();

  Office.context.mailbox.item.body.setAsync(html, { coercionType: Office.CoercionType.Html });
}