function generateProductConsumptionReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const consumptionSheet = ss.getSheetByName('wpConsumption');
  const productMasterSheet = ss.getSheetByName('beData'); // Optional
  const reportSheet = ss.getSheetByName('rProduct');

  // Date inputs from D2 (From Date) and F2 (To Date)
  const fromDate = new Date(reportSheet.getRange('D2').getValue());
  const toDate = new Date(reportSheet.getRange('F2').getValue());
  toDate.setHours(23,59,59,999);

  // Filters from rProduct sheet
  const filterProdName = reportSheet.getRange('B2').getValue().toString().trim().toLowerCase();
  const filterProdType = reportSheet.getRange('C2').getValue().toString().trim().toLowerCase();
  const filterBrand = reportSheet.getRange('G2').getValue().toString().trim().toLowerCase();

  // Read all consumption data
  const rows = consumptionSheet.getDataRange().getValues();
  const header = rows[0];
  const dataRows = rows.slice(1);

  // Find column indexes in wpConsumption
  const idxTimestamp = header.indexOf('Timestamp');
  const idxProdName = header.indexOf('Product Names');
  const idxProdType = header.indexOf('Product Types');
  const idxQtyUsed = header.indexOf('Quantities Used');
  const idxUnits = header.indexOf('Units');

  // Load product master data (beData) into a map with cost and brand per product
  let prodMaster = {};
  if (productMasterSheet) {
    const mRows = productMasterSheet.getDataRange().getValues();
    for (let i = 1; i < mRows.length; i++) {
      const prodName = mRows[i][1];  // B column
      const cost = parseFloat(mRows[i][5]); // F column
      const brand = mRows[i][6]; // G column
      if (prodName && !isNaN(cost)) {
        prodMaster[prodName] = { cost: cost, brand: brand || '' };
      }
    }
  }

  // Aggregate quantities by product + unit
  let tally = {}; // key: prodName|unit => {name, type, qtySum, unit, cost, brand}

  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    let ts = row[idxTimestamp];
    let dateObj = (ts instanceof Date) ? ts : new Date(ts);
    if (isNaN(dateObj) || dateObj < fromDate || dateObj > toDate) continue;

    const name = row[idxProdName];
    const type = row[idxProdType];
    const qty = parseFloat(row[idxQtyUsed]);
    const unit = row[idxUnits];

    if (!name || isNaN(qty)) continue;

    // Lookup brand for current product
    const brand = (prodMaster[name] && prodMaster[name].brand) ? prodMaster[name].brand.toString().toLowerCase() : '';

    // Apply dynamic filters (skip if filter value is blank)
    if (filterProdName && !name.toString().toLowerCase().includes(filterProdName)) continue;
    if (filterProdType && !type.toString().toLowerCase().includes(filterProdType)) continue;
    if (filterBrand && !brand.includes(filterBrand)) continue;

    const key = name + '|' + unit;

    if (!tally[key]) {
      let cost = '';
      let brandText = '';
      if (prodMaster[name]) {
        cost = prodMaster[name].cost;
        brandText = prodMaster[name].brand || '';
      }
      tally[key] = { name, type, qtySum: 0, unit, cost, brand: brandText };
    }

    tally[key].qtySum += qty;
  }

  // Prepare output with Total Cost calculated (qtySum * cost per unit)
  const output = [
    ['Product Name', 'Product Type', 'Total Quantity Used', 'Units', 'Total Cost', 'Brand']
  ];

  Object.values(tally).forEach(obj => {
    let totalCost = '';
    if (obj.cost !== '' && !isNaN(obj.cost)) {
      totalCost = obj.qtySum * obj.cost;
      totalCost = Math.round(totalCost * 100) / 100; // round to 2 decimals
    }
    output.push([
      obj.name,
      obj.type,
      obj.qtySum,
      obj.unit,
      totalCost,
      obj.brand || ''
    ]);
  });

  // Clear old data and write report from B5
  reportSheet.getRange(5, 2, reportSheet.getMaxRows() - 4, 6).clearContent();
  reportSheet.getRange(5, 2, output.length, output[0].length).setValues(output);
}
