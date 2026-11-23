"use strict";

const search = document.getElementById("search");
const status = document.getElementById("status");
const results = document.getElementById("results");
const snapshot = document.getElementById("snapshot");
let records = [];

function field(container, label, value) {
  const heading = document.createElement("h3");
  heading.textContent = label;
  const content = document.createElement("pre");
  content.textContent = value || "（無）";
  container.append(heading, content);
}

function render() {
  const query = search.value.trim().toLocaleLowerCase();
  const matches = records.filter(record =>
    [record.id, record.short, record.long, record.steps]
      .some(value => value.toLocaleLowerCase().includes(query))
  );
  const list = document.createDocumentFragment();
  for (const record of matches) {
    const details = document.createElement("details");
    const title = document.createElement("summary");
    title.textContent = `${record.id} — ${record.short}`;
    details.append(title);
    field(details, "OBS ID", record.id);
    field(details, "Short Description", record.short);
    field(details, "Long Description", record.long);
    field(details, "Steps to Reproduce", record.steps);
    field(details, "建立時間", record.created);

    const heading = document.createElement("h3");
    heading.textContent = "附件";
    details.append(heading);
    if (record.attachments.length === 0) {
      const empty = document.createElement("p");
      empty.textContent = "無附件。";
      details.append(empty);
    } else {
      const files = document.createElement("ul");
      for (const attachment of record.attachments) {
        const item = document.createElement("li");
        const link = document.createElement("a");
        link.textContent = attachment.name;
        link.href = attachment.path;
        link.download = attachment.name;
        item.append(link);
        files.append(item);
      }
      details.append(files);
    }
    list.append(details);
  }
  results.replaceChildren(list);
  status.textContent = `符合 ${matches.length} 筆，共 ${records.length} 筆。`;
}

fetch("data.json")
  .then(response => {
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return response.json();
  })
  .then(data => {
    records = data.records;
    snapshot.textContent = `快照：${data.snapshot}`;
    search.disabled = false;
    search.addEventListener("input", render);
    render();
  })
  .catch(error => {
    status.textContent = `無法讀取快照資料：${error.message}。請確認所有批次已上傳並完成部署。`;
  });
