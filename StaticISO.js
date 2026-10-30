/**
 * Static ISO service helpers.
 * LiDAR.js and WiFi.js keep their domain-specific readers, while shared
 * Static ISO sheet discovery lives here so it is defined only once.
 */
function getSheetFlexible(ss, candidateNames) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var exactName = sheets[i].getName().trim().toLowerCase();
    for (var j = 0; j < candidateNames.length; j++) {
      if (exactName === String(candidateNames[j]).toLowerCase()) return sheets[i];
    }
  }
  for (var k = 0; k < sheets.length; k++) {
    var partialName = sheets[k].getName().toLowerCase();
    for (var n = 0; n < candidateNames.length; n++) {
      if (partialName.indexOf(String(candidateNames[n]).toLowerCase()) !== -1) return sheets[k];
    }
  }
  return null;
}
