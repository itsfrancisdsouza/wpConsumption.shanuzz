function searchClientVisitDataBlockFormat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const consultSheet = ss.getSheetByName('wpConsultation');
  const consumeSheet = ss.getSheetByName('wpConsumption');
  const outSheet = ss.getSheetByName('clientVisitData');

  // Clear previous output (A4:C)
  outSheet.getRange(3, 1, outSheet.getMaxRows() - 2, 3).clearContent();

  const searchValue = outSheet.getRange("A2").getValue().toString().trim();
  const searchType = outSheet.getRange("A1").getValue().toString().toLowerCase();

  if (!searchValue || !["submit id", "client name", "phone number"].includes(searchType)) {
    outSheet.getRange("A4").setValue("Please enter value in A1 and valid search type (Submit ID, Client Name, or Phone Number) in B1.");
    return;
  }

  // Read consultation data
  const consultData = consultSheet.getDataRange().getValues();
  const consultHeader = consultData[0];
  const cRows = consultData.slice(1);

  // Read consumption data
  const consumeData = consumeSheet.getDataRange().getValues();
  const consumeHeader = consumeData[0];
  const cnRows = consumeData.slice(1);

  // Columns for looking up in consultation
  const col = {};
  consultHeader.forEach((x,i)=>col[x.trim().toLowerCase()] = i);

  // Find matching consultation rows
  let matches = [];
  if (searchType === "submit id") {
    matches = cRows.filter(r => String(r[col["submit id"]]).trim() === searchValue);
  } else if (searchType === "client name") {
    matches = cRows.filter(r => String(r[col["client name"]]).toLowerCase().includes(searchValue.toLowerCase()));
  } else if (searchType === "phone number") {
    matches = cRows.filter(r => String(r[col["phone number"]]).includes(searchValue));
  }

  if (matches.length === 0) {
    outSheet.getRange("A4").setValue("No matching client visit data found.");
    return;
  }

  // For each match, output block details
  let out = [];
  matches.forEach(m => {
    // Block fields
    out.push(['Submit ID', m[col["submit id"]]]);
    out.push(['Timestamp', m[col["timestamp"]]]);
    out.push(['Client Name', m[col["client name"]]]);
    out.push(['Phone Number', m[col["phone number"]]]);
    out.push(['Email Address', m[col["email address"]]]);
    out.push(['Visit Date', m[col["visit date"]]]);
    out.push(['Client Type', m[col["client type"]]]);
    out.push(['Status', m[col["status"]]]);
    out.push(['Desired Services', m[col["desired services"]]]);
    out.push(['Hair Type', m[col["hair type"]]]);
    out.push(['Hair Texture', m[col["hair texture"]]]);
    out.push(['Scalp Condition', m[col["scalp condition"]]]);
    out.push(['Previous Treatments', m[col["previous treatments"]]]);
    out.push(['Allergies/Sensitivities', m[col["allergies/sensitivities"]]]);
    out.push(['Client Expectations', m[col["client expectations"]]]);
    out.push(['Consent Given', m[col["consent given"]]]);
    out.push(['Digital Signature', m[col["digital signature"]]]);
    out.push(['Before Photo', m[col["before photo"]]]);

    // --- Consumption details ---
    out.push(['---Consumption Details---', '']);

    // Find all consumptions for this submit id
    let cMatches = cnRows.filter(cn => String(cn[consumeHeader.indexOf('Consultation ID')]) === String(m[col["submit id"]]));
    // Gather combined fields
    let prodNames = [];
    let prodTypes = [];
    let prodQuant = [];
    let stylists = [];
    let totalCost = '';
    let notes='', sat='', followup='', afterPhoto='', stat='';

    cMatches.forEach(cn => {
      // If this row is a product/consumption row
      if (cn[consumeHeader.indexOf('Product Names')]) {
        prodNames.push(cn[consumeHeader.indexOf('Product Names')]);
        prodTypes.push(cn[consumeHeader.indexOf('Product Types')]);
        prodQuant.push(cn[consumeHeader.indexOf('Quantities Used')] + ' ' + cn[consumeHeader.indexOf('Units')]);
      }
      // If stylist/incentive split row
      if (cn[consumeHeader.indexOf('Stylists/Technicians')] && cn[consumeHeader.indexOf('Incentive Splits')]) {
        stylists.push(String(cn[consumeHeader.indexOf('Stylists/Technicians')]) +
            '(' + cn[consumeHeader.indexOf('Incentive Splits')] + ')');
      }
      // Cost, notes, satisfaction, etc (only once)
      if (cn[consumeHeader.indexOf('Total Cost')] && !totalCost)
        totalCost = cn[consumeHeader.indexOf('Total Cost')];
      if (cn[consumeHeader.indexOf('Service Notes')] && !notes)
        notes = cn[consumeHeader.indexOf('Service Notes')];
      if (cn[consumeHeader.indexOf('Client Satisfaction')] && !sat)
        sat = cn[consumeHeader.indexOf('Client Satisfaction')];
      if (cn[consumeHeader.indexOf('Follow-up Recommendations')] && !followup)
        followup = cn[consumeHeader.indexOf('Follow-up Recommendations')];
      if (cn[consumeHeader.indexOf('After Photo')] && !afterPhoto)
        afterPhoto = cn[consumeHeader.indexOf('After Photo')];
      if (cn[consumeHeader.indexOf('Status')] && !stat)
        stat = cn[consumeHeader.indexOf('Status')];
    });

    out.push(['Product Names', prodNames.join(', ') || '']);
    out.push(['Product Types', prodTypes.join(', ') || '']);
    out.push(['Quantities Used', prodQuant.join(', ') || '']);
    out.push(['Stylists/Technicians', stylists.join(', ') || '']);
    out.push(['Total Cost', totalCost]);
    out.push(['Service Notes', notes]);
    out.push(['Client Satisfaction', sat]);
    out.push(['Follow-up Recommendations', followup]);
    out.push(['After Photo', afterPhoto]);

    // Row spacer after each block
    out.push(['', '']);
  });

  // Write output
  outSheet.getRange(4,1,out.length,2).setValues(out);
}
