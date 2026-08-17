function onEdit(e) {
  const sheet = e.range.getSheet();

  switch (sheet.getName()) {

    case "Studio Master":
      assignStudioId(e);
      break;

    case "Staff Master":
      assignStaffId(e);
      break;

    case "Services Master":
      assignServiceId(e);
      break;

    case "Lists":
      handleLists(e);
      break;
  }
}

function assignStudioId(e) {
  const sheet = e.range.getSheet();

  if (e.range.getColumn() !== 2 || e.range.getRow() === 1) return;

  const idCell = sheet.getRange(e.range.getRow(), 1);

  if (idCell.getValue() !== "") return;
  if (e.range.getValue() === "") return;

  idCell.setValue(generateNextId(sheet, "ST"));
}

function assignStaffId(e) {
  const sheet = e.range.getSheet();

  if (e.range.getColumn() !== 2 || e.range.getRow() === 1) return;

  const idCell = sheet.getRange(e.range.getRow(), 1);

  if (idCell.getValue() !== "") return;
  if (e.range.getValue() === "") return;

  idCell.setValue(generateNextId(sheet, "SF"));
}

function assignServiceId(e) {
  const sheet = e.range.getSheet();

  if (e.range.getColumn() !== 2 || e.range.getRow() === 1) return;

  const idCell = sheet.getRange(e.range.getRow(), 1);

  if (idCell.getValue() !== "") return;
  if (e.range.getValue() === "") return;

  idCell.setValue(generateNextId(sheet, "SV"));
}