function saveStudio(data) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Studio Master");

  if (!sheet) {
    throw new Error("Studio Master sheet not found.");
  }

  const values = sheet.getDataRange().getValues();
  const studio = data.studioName.trim().toLowerCase();
  const owner = data.ownerName.trim().toLowerCase();
  const city = data.city.trim().toLowerCase();

  // Check for duplicate Studio + Owner + City
  for (let i = 1; i < values.length; i++) {
    const existingStudio = String(values[i][1]).trim().toLowerCase();
    const existingOwner = String(values[i][3]).trim().toLowerCase();
    const existingCity = String(values[i][6]).trim().toLowerCase();

    if (
      existingStudio === studio &&
      existingOwner === owner &&
      existingCity === city
    ) {
      throw new Error("Studio already exists.");
    }
  }

  // Generate Studio ID
  const studioId = generateNextId(sheet, "ST");

  // Generate Display Name
  const displayName =
    data.studioName.trim() +
    " - " +
    data.ownerName.trim() +
    " - " +
    data.city.trim();

  // Find last Studio ID in Column A
  const ids = sheet.getRange("A:A").getValues();
  let lastStudioRow = 1;

  for (let i = ids.length - 1; i >= 1; i--) {
    if (String(ids[i][0]).trim() !== "") {
      lastStudioRow = i + 1;
      break;
    }
  }

  // Insert after last Studio ID
  sheet.getRange(lastStudioRow + 1, 1, 1, 8).setValues([[
    studioId,
    data.studioName.trim(),
    displayName,
    data.ownerName.trim(),
    data.mobile.trim(),
    data.whatsapp.trim(),
    data.city.trim(),
    data.notes.trim()
  ]]);

  return {
    success: true,
    studioId: studioId
  };
}