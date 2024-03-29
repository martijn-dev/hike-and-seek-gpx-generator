import * as XLSX from 'xlsx';
import fs from 'fs';
// eslint-disable-next-line @typescript-eslint/no-var-requires
const rdToWgs84 = require('rd-to-wgs84');

// eslint-disable-next-line @typescript-eslint/no-var-requires
const pjson = require('./package.json');
type Settings = {
  'insert-pre-coord-number-x': number;
  'insert-pre-coord-number-y': number;
  'insert-after-coord-number-x': number;
  'insert-after-coord-number-y': number;
};
type Type = {
  type: string;
  icon: string;
  color: string;
  osmand_icon: string;
};

type Data = {
  'coord-x': number;
  'coord-y': number;
  coord: string;
  post: number;
  type: string;
  description: string;
};

type Wgs84Coord = {
  error: any;
  lat: number;
  lon: number;
};

const workbook = XLSX.readFile('insert-file.xlsx');
const sheet_name_list = workbook.SheetNames;
console.log(workbook.SheetNames);
const settings = XLSX.utils.sheet_to_json<Settings>(workbook.Sheets['Settings'])[0];
const types = XLSX.utils.sheet_to_json<Type>(workbook.Sheets['Types']);
const data = XLSX.utils.sheet_to_json<Data>(workbook.Sheets['Data']);

let finalGPXString = `<?xml version='1.0' encoding='UTF-8' standalone='yes' ?>
<gpx version="1.1" creator="${
  pjson.author
}" xmlns="http://www.topografix.com/GPX/1/1"  xmlns:osmand="https://osmand.net">
<metadata>
<name>Hike and Seek ${new Date().getFullYear()}</name>
<desc />
<time>${new Date().toISOString()}</time>
</metadata>`;

data.forEach((spot) => {
  if (!spot.coord) return;
  spot.coord = spot.coord.toString().replaceAll('-', '');
  spot.coord = spot.coord.toString().replaceAll(' ', '');
  let coordX = spot.coord.toString().substring(0, spot.coord.length * 0.5);
  let coordY = spot.coord.toString().substring(spot.coord.length * 0.5);
  coordX = `${settings['insert-pre-coord-number-x'] ?? ''}${coordX}${settings['insert-after-coord-number-x'] ?? ''}`;
  coordY = `${settings['insert-pre-coord-number-y'] ?? ''}${coordY}${settings['insert-after-coord-number-y'] ?? ''}`;
  console.log(`Insert (${coordX}, ${coordY}) for spot ${spot.post}`);
  const wgsCoord = rdToWgs84(coordX, coordY) as Wgs84Coord;
  if (wgsCoord.error) {
    console.error(
      `Failed to convert coordinate for spot ${spot.post} (${spot['coord-x']},${spot['coord-y']})`,
      wgsCoord.error,
    );
    return;
  }
  const type = types.find((type) => type.type === spot.type);
  const wptString = `<wpt lat="${wgsCoord.lat}" lon="${wgsCoord.lon}">
  <ele>0.0</ele>
  <time>${new Date().toISOString()}</time>
  <name>${spot.post}</name>
  <desc>${spot.description}</desc>
  <sym>${type?.icon ?? 'Flag, Blue'}</sym>
  <extensions>
  <osmand:color>${type?.color ?? '#ff0000'}</osmand:color>
  <osmand:icon>${type?.osmand_icon ?? 'special_flag_stroke'}</osmand:icon>
  <osmand:background>circle</osmand:background>
  </extensions>
  </wpt>`;

  finalGPXString += wptString;
  console.log('data', wptString);

  console.log(spot['coord-y'], coordY);
  console.log(spot['coord-x'], coordX);
});

finalGPXString += '</gpx>';

fs.writeFile(`${new Date().getFullYear()}.gpx`, finalGPXString, function (err) {
  if (err) return console.error('Error writing gpx data to file', err);
  console.log(`${new Date().getFullYear()}.gpx created!`);
});
