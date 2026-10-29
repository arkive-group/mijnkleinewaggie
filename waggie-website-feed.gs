/**
 * Waggi-E website feed
 * Reads the Products tab of the Waggi-E sales infrastructure sheet and returns
 * only the public fields as JSON for mijnkleinewaggie.nl.
 * Supplier cost and margin are never sent to the website.
 *
 * Shown on the website: rows with Status "Live" or "Coming soon".
 */
const SHEET_ID = '1cksGD_Dr5EgAdfLd7cEy8PJ7qJsrSsMgcdIRqMjFXFk';
const TAB_NAME = 'Products';

function doGet() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_NAME);
  const values = sheet.getDataRange().getDisplayValues();

  // The header row is the row that starts with "Product ID" (row 1 holds group headers)
  const headerIndex = values.findIndex(r => String(r[0]).trim() === 'Product ID');
  const headers = values[headerIndex].map(h => String(h).trim());
  const col = name => headers.indexOf(name);

  const pick = (row, name) => (col(name) > -1 ? String(row[col(name)]).trim() : '');

  const products = values.slice(headerIndex + 1)
    .filter(r => r[0])
    .filter(r => ['live', 'coming soon'].includes(pick(r, 'Status').toLowerCase()))
    .map(r => ({
      id: pick(r, 'Product ID'),
      status: pick(r, 'Status'),
      brand: pick(r, 'Brand'),
      model: pick(r, 'Model'),
      variant: pick(r, 'Variant'),
      title: pick(r, 'Ad title'),
      description: pick(r, 'Description'),
      price: pick(r, 'Selling price (EUR)'),
      condition: pick(r, 'Condition'),
      stock: pick(r, 'Stock'),
      images: [pick(r, 'Image URL 1'), pick(r, 'Image URL 2'), pick(r, 'Image URL 3')].filter(Boolean),
      testDrive: pick(r, 'Test drive location'),
      vehicleClass: pick(r, 'Vehicle class'),
      licence: pick(r, 'Licence'),
      topSpeed: pick(r, 'Top speed (km/h)'),
      range: pick(r, 'Range (km)'),
      battery: pick(r, 'Battery (kWh)'),
      seats: pick(r, 'Seats'),
      warranty: pick(r, 'Warranty (years)'),
      colour: pick(r, 'Colour')
    }));

  return ContentService
    .createTextOutput(JSON.stringify({ updated: new Date().toISOString(), products }))
    .setMimeType(ContentService.MimeType.JSON);
}
