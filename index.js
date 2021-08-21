const fs = require('fs');
const Excel = require('exceljs');

const raw = fs.readFileSync('./qldGeoJson.json');

const se2Map = {};

const geoJson = JSON.parse(raw);

const rawLocalityFilename = './meshblock-correspondence-file-asgs-2016.xlsx';

const normaliseString = (str) => {
  return str.toLowerCase().replace(/ *\([^)]*\) */g, "").replace(/[^\w]/gi, '');
}

const random_hex_color_code = () => {
  let n = (Math.random() * 0xfffff * 1000000).toString(16);
  return '#' + n.slice(0, 6);
};

const run = async () => {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(rawLocalityFilename);
  const workSheet = workbook.getWorksheet('Data');
  workSheet.eachRow(row => {
    const SE2 = normaliseString(row.getCell('R').value);
    const SE4 = row.getCell('K').value;
    e4map[SE4] = random_hex_color_code();

    se2Map[SE2] = SE4;
  });

  for (let i = 0; i < geoJson.features.length; i++){
    const geoJsonSE2 = normaliseString(geoJson.features[i].properties.qld_loca_2);

    if (se2Map[geoJsonSE2]) {
      geoJson.features[i].properties = {
        se4: se2Map[geoJsonSE2],
        se2: geoJson.features[i].properties.qld_loca_2,
      };
    }
  }

  fs.writeFileSync('./output.json', JSON.stringify(geoJson))
}

let e4map = {};

const getSe4 = async () => {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(rawLocalityFilename);
  const workSheet = workbook.getWorksheet('Data');
  workSheet.eachRow(row => {
    const SE4 = row.getCell('K').value;
    e4map[SE4] = random_hex_color_code();
  });

  fs.writeFileSync('./e4.json', JSON.stringify(e4map))
}

getSe4();