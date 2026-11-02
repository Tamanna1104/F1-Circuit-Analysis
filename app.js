// F1 Circuit Analytics — loads an Excel/CSV file in your browser.
// Expected columns include: Circuit Id, Circuit Name, Country, City,
// Direction, Type, Altitude M, Lap Km, Laps, Latitude, Longitude.

const COLORS = ['#ed1c24','#ff6b35','#63a6ff','#f6bd60','#72c6a1','#b49aff','#f08ab4','#aeb8c4','#d7df23'];
const charts = {};
let allData = [];
let map, mapLayer;

const demoData = [
 {circuit:'Bahrain International Circuit',country:'Bahrain',city:'Sakhir',laps:57,lapKm:5.412,altitude:7,direction:'Anti-clockwise',type:'Permanent',lat:26.0325,lon:50.5106},
 {circuit:'Albert Park Grand Prix Circuit',country:'Australia',city:'Melbourne',laps:58,lapKm:5.278,altitude:10,direction:'Clockwise',type:'Street',lat:-37.8497,lon:144.968},
 {circuit:'Shanghai International Circuit',country:'China',city:'Shanghai',laps:56,lapKm:5.451,altitude:5,direction:'Clockwise',type:'Permanent',lat:31.3389,lon:121.220},
 {circuit:'Suzuka International Racing Course',country:'Japan',city:'Suzuka',laps:53,lapKm:5.807,altitude:45,direction:'Clockwise',type:'Permanent',lat:34.8431,lon:136.541},
 {circuit:'Circuit de Monaco',country:'Monaco',city:'Monte Carlo',laps:78,lapKm:3.337,altitude:10,direction:'Clockwise',type:'Street',lat:43.7347,lon:7.4206},
 {circuit:'Circuit Gilles-Villeneuve',country:'Canada',city:'Montreal',laps:70,lapKm:4.361,altitude:13,direction:'Clockwise',type:'Permanent',lat:45.5000,lon:-73.5228},
 {circuit:'Circuit de Barcelona-Catalunya',country:'Spain',city:'Montmelo',laps:66,lapKm:4.657,altitude:109,direction:'Clockwise',type:'Permanent',lat:41.57,lon:2.2611},
 {circuit:'Red Bull Ring',country:'Austria',city:'Spielberg',laps:71,lapKm:4.318,altitude:678,direction:'Clockwise',type:'Permanent',lat:47.2197,lon:14.7647},
 {circuit:'Silverstone Circuit',country:'UK',city:'Silverstone',laps:52,lapKm:5.891,altitude:153,direction:'Clockwise',type:'Permanent',lat:52.0786,lon:-1.0169},
 {circuit:'Hungaroring',country:'Hungary',city:'Mogyorod',laps:70,lapKm:4.381,altitude:238,direction:'Clockwise',type:'Permanent',lat:47.5789,lon:19.2486},
 {circuit:'Circuit de Spa-Francorchamps',country:'Belgium',city:'Stavelot',laps:44,lapKm:7.004,altitude:401,direction:'Clockwise',type:'Permanent',lat:50.4372,lon:5.9714},
 {circuit:'Autodromo Nazionale Monza',country:'Italy',city:'Monza',laps:53,lapKm:5.793,altitude:162,direction:'Clockwise',type:'Permanent',lat:45.6156,lon:9.2811},
 {circuit:'Marina Bay Street Circuit',country:'Singapore',city:'Singapore',laps:62,lapKm:4.94,altitude:18,direction:'Anti-clockwise',type:'Street',lat:1.2914,lon:103.864},
 {circuit:'Suzuka East',country:'Japan',city:'Suzuka',laps:50,lapKm:3.2,altitude:42,direction:'Clockwise',type:'Permanent',lat:34.84,lon:136.54},
 {circuit:'Autodromo Jose Carlos Pace',country:'Brazil',city:'Sao Paulo',laps:71,lapKm:4.309,altitude:785,direction:'Anti-clockwise',type:'Permanent',lat:-23.7013,lon:-46.697},
 {circuit:'Autodromo Hermanos Rodriguez',country:'Mexico',city:'Mexico City',laps:71,lapKm:4.304,altitude:2220,direction:'Clockwise',type:'Permanent',lat:19.4042,lon:-99.0907},
 {circuit:'Circuit of the Americas',country:'USA',city:'Austin',laps:56,lapKm:5.513,altitude:161,direction:'Anti-clockwise',type:'Permanent',lat:30.1328,lon:-97.6411},
 {circuit:'Yas Marina Circuit',country:'UAE',city:'Abu Dhabi',laps:58,lapKm:5.281,altitude:3,direction:'Anti-clockwise',type:'Permanent',lat:24.4672,lon:54.6031},
 {circuit:'Baku City Circuit',country:'Azerbaijan',city:'Baku',laps:51,lapKm:6.003,altitude:-25,direction:'Anti-clockwise',type:'Street',lat:40.3725,lon:49.8533},
 {circuit:'Jeddah Corniche Circuit',country:'Saudi Arabia',city:'Jeddah',laps:50,lapKm:6.174,altitude:5,direction:'Anti-clockwise',type:'Street',lat:21.6319,lon:39.1044},
 {circuit:'Miami International Autodrome',country:'USA',city:'Miami',laps:57,lapKm:5.412,altitude:2,direction:'Anti-clockwise',type:'Street',lat:25.9581,lon:-80.2389},
 {circuit:'Las Vegas Strip Circuit',country:'USA',city:'Las Vegas',laps:50,lapKm:6.201,altitude:610,direction:'Anti-clockwise',type:'Street',lat:36.1147,lon:-115.1728},
 {circuit:'Circuit Zandvoort',country:'Netherlands',city:'Zandvoort',laps:72,lapKm:4.259,altitude:6,direction:'Clockwise',type:'Permanent',lat:52.3888,lon:4.5409},
 {circuit:'Lusail International Circuit',country:'Qatar',city:'Lusail',laps:57,lapKm:5.419,altitude:12,direction:'Clockwise',type:'Permanent',lat:25.49,lon:51.4542}
];

const $ = id => document.getElementById(id);
const num = v => { const n = Number(String(v ?? '').replace(/,/g,'')); return Number.isFinite(n) && String(v ?? '').trim() !== '' ? n : null; };
const clean = v => String(v ?? '').trim();
function normalizeKey(k){return clean(k).toLowerCase().replace(/[^a-z0-9]/g,'');}
function pick(row, names){const keys=Object.keys(row);for(const name of names){const found=keys.find(k=>normalizeKey(k)===normalizeKey(name));if(found!==undefined)return row[found];}return '';}
function normalizeRows(rows){
 return rows.map(r=>{
  const circuit=clean(pick(r,['Circuit Name','Circuit','Name','circuit_name']));
  const country=clean(pick(r,['Country','country_name']));
  if(!circuit && !country) return null;
  return {
   circuit:circuit || clean(pick(r,['Circuit Id','Circuit ID','Id'])) || 'Unnamed circuit',
   country:country || 'Unknown',
   city:clean(pick(r,['City','Location'])) || '—',
   laps:num(pick(r,['Laps','Lap Count','Number of Laps'])),
   lapKm:num(pick(r,['Lap Km','Lap Distance','Length','Circuit Length','Lap Distance Km'])),
   altitude:num(pick(r,['Altitude M','Altitude','Elevation','Altitude (m)'])),
   direction:clean(pick(r,['Direction','Circuit Direction'])) || 'Unknown',
   type:clean(pick(r,['Type','Circuit Type'])) || 'Unknown',
   lat:num(pick(r,['Latitude','Lat'])),
   lon:num(pick(r,['Longitude','Long','Lon','Lng']))
  };
 }).filter(Boolean);
}
function filteredData(){
 const country=$('countryFilter').value,dir=$('directionFilter').value,type=$('typeFilter').value,q=$('searchFilter').value.toLowerCase();
 return allData.filter(d=>(!country||d.country===country)&&(!dir||d.direction===dir)&&(!type||d.type===type)&&(!q||`${d.circuit} ${d.country} ${d.city}`.toLowerCase().includes(q)));
}
function fillSelect(id, values, label){
 const el=$(id),old=el.value;el.innerHTML=`<option value="">${label}</option>`+[...new Set(values.filter(Boolean))].sort().map(v=>`<option>${escapeHtml(v)}</option>`).join('');if([...el.options].some(o=>o.value===old))el.value=old;
}
function escapeHtml(s){return String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function setStatus(text){$('dataStatus').textContent=text;}
function chart(id,type,data,options={}){
 if(charts[id])charts[id].destroy();
 const ctx=$(id).getContext('2d');
 charts[id]=new Chart(ctx,{type,data,options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#c4cbd3',boxWidth:10,font:{size:10}}},tooltip:{backgroundColor:'#080a0d',titleColor:'#fff',bodyColor:'#d9dee4'}},scales:type==='doughnut'||type==='pie'?{}:{x:{ticks:{color:'#9ba4ae',font:{size:9}},grid:{color:'#252c34'}},y:{ticks:{color:'#9ba4ae',font:{size:9}},grid:{color:'#252c34'}}},...options}});
}
function update(){
 const data=filteredData();$('recordCount').textContent=`${data.length} circuits shown`;
 const countries=[...new Set(data.map(d=>d.country))];
 const validDist=data.filter(d=>d.lapKm!==null),validAlt=data.filter(d=>d.altitude!==null);
 $('kpiCircuits').textContent=data.length;
 $('kpiCountries').textContent=countries.length;
 $('kpiDistance').textContent=validDist.length?(validDist.reduce((s,d)=>s+d.lapKm,0)/validDist.length).toFixed(2)+' km':'—';
 $('kpiAltitude').textContent=validAlt.length?Math.round(validAlt.reduce((s,d)=>s+d.altitude,0)/validAlt.length).toLocaleString()+' m':'—';
 const countryCounts={};data.forEach(d=>countryCounts[d.country]=(countryCounts[d.country]||0)+1);
 const countryEntries=Object.entries(countryCounts).sort((a,b)=>b[1]-a[1]);
 chart('countryChart','bar',{labels:countryEntries.map(x=>x[0]),datasets:[{label:'Circuits',data:countryEntries.map(x=>x[1]),backgroundColor:COLORS[0],borderRadius:2}]},{indexAxis:'y',plugins:{legend:{display:false}}});
 const lapData=data.filter(d=>d.laps!==null).sort((a,b)=>b.laps-a.laps).slice(0,12);
 chart('lapsChart','bar',{labels:lapData.map(d=>shortName(d.circuit)),datasets:[{label:'Laps',data:lapData.map(d=>d.laps),backgroundColor:COLORS[2],borderRadius:2}]},{indexAxis:'y',plugins:{legend:{display:false}}});
 const distData=validDist.slice().sort((a,b)=>b.lapKm-a.lapKm).slice(0,12);
 chart('distanceChart','bar',{labels:distData.map(d=>shortName(d.circuit)),datasets:[{label:'Lap distance (km)',data:distData.map(d=>d.lapKm),backgroundColor:COLORS[1],borderRadius:2}]},{indexAxis:'y',plugins:{legend:{display:false}}});
 chart('altDistanceChart','scatter',{datasets:[{label:'Circuits',data:data.filter(d=>d.altitude!==null&&d.lapKm!==null).map(d=>({x:d.lapKm,y:d.altitude,label:d.circuit})),backgroundColor:COLORS[4],pointRadius:5}]},{plugins:{tooltip:{callbacks:{label:c=>`${c.raw.label}: ${c.raw.x} km, ${c.raw.y} m`}}},scales:{x:{title:{display:true,text:'Lap distance (km)',color:'#aab2bc'},ticks:{color:'#9ba4ae'},grid:{color:'#252c34'}},y:{title:{display:true,text:'Altitude (m)',color:'#aab2bc'},ticks:{color:'#9ba4ae'},grid:{color:'#252c34'}}}});
 const directions={};data.forEach(d=>directions[d.direction]=(directions[d.direction]||0)+1);
 chart('directionChart','doughnut',{labels:Object.keys(directions),datasets:[{data:Object.values(directions),backgroundColor:COLORS,borderColor:'#14181e',borderWidth:3}]},{plugins:{legend:{position:'bottom'}}});
 const types={};data.forEach(d=>types[d.type]=(types[d.type]||0)+1);
 chart('typeChart','bar',{labels:Object.keys(types),datasets:[{label:'Circuits',data:Object.values(types),backgroundColor:COLORS.slice(0,Object.keys(types).length),borderRadius:3}]},{plugins:{legend:{display:false}}});
 const altByCountry={};validAlt.forEach(d=>{if(!altByCountry[d.country])altByCountry[d.country]=[];altByCountry[d.country].push(d.altitude);});
 const altEntries=Object.entries(altByCountry).map(([k,v])=>[k,v.reduce((a,b)=>a+b,0)/v.length]).sort((a,b)=>b[1]-a[1]).slice(0,12);
 chart('altitudeChart','bar',{labels:altEntries.map(x=>x[0]),datasets:[{label:'Average altitude (m)',data:altEntries.map(x=>Math.round(x[1])),backgroundColor:COLORS[2],borderRadius:2}]},{indexAxis:'y',plugins:{legend:{display:false}}});
 $('detailsTable').innerHTML=data.map(d=>`<tr><td>${escapeHtml(d.circuit)}</td><td>${escapeHtml(d.country)}</td><td>${escapeHtml(d.city)}</td><td>${d.laps??'—'}</td><td>${d.lapKm??'—'}</td><td>${d.altitude??'—'}</td><td>${escapeHtml(d.direction)}</td><td>${escapeHtml(d.type)}</td></tr>`).join('')||'<tr><td colspan="8">No matching records.</td></tr>';
 const topCountry=countryEntries[0];$('scenarioCountry').textContent=topCountry?`${topCountry[0]} (${topCountry[1]})`:'No data';$('scenarioCountryNote').textContent=topCountry?'Country with the most circuits in the current selection.':'Upload or select data.';
 const longest=validDist.slice().sort((a,b)=>b.lapKm-a.lapKm)[0];$('scenarioDistance').textContent=longest?shortName(longest.circuit):'No data';$('scenarioDistanceNote').textContent=longest?`${longest.lapKm} km per lap`:'Lap distance is not available.';
 const highest=validAlt.slice().sort((a,b)=>b.altitude-a.altitude)[0];$('scenarioAltitude').textContent=highest?shortName(highest.circuit):'No data';$('scenarioAltitudeNote').textContent=highest?`${highest.altitude.toLocaleString()} metres`:'Altitude is not available.';
 updateMap(data);
}
function shortName(s){return s.length>24?s.slice(0,22)+'…':s;}
function updateMap(data){
 if(!map)return;if(mapLayer)map.removeLayer(mapLayer);
 const points=data.filter(d=>d.lat!==null&&d.lon!==null);
 mapLayer=L.layerGroup(points.map(d=>L.circleMarker([d.lat,d.lon],{radius:7,color:'#fff',weight:1,fillColor:'#ed1c24',fillOpacity:.9}).bindPopup(`<b>${escapeHtml(d.circuit)}</b><br>${escapeHtml(d.city)}, ${escapeHtml(d.country)}<br>Lap: ${d.lapKm??'—'} km · Altitude: ${d.altitude??'—'} m`))).addTo(map);
}
function initMap(){map=L.map('circuitMap',{scrollWheelZoom:false}).setView([22,15],2);L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{attribution:'&copy; OpenStreetMap contributors',maxZoom:18}).addTo(map);}
function initFilters(){fillSelect('countryFilter',allData.map(d=>d.country),'All countries');fillSelect('directionFilter',allData.map(d=>d.direction),'All directions');fillSelect('typeFilter',allData.map(d=>d.type),'All types');}
async function loadExcel(file){
 try{
  const buffer=await file.arrayBuffer();let rows;
  if(file.name.toLowerCase().endsWith('.csv')){const text= new TextDecoder().decode(buffer);rows=XLSX.utils.sheet_to_json(XLSX.read(text,{type:'string'}).Sheets[XLSX.read(text,{type:'string'}).SheetNames[0]],{defval:''});}
  else{const wb=XLSX.read(buffer,{type:'array'});const sheet=wb.Sheets[wb.SheetNames[0]];rows=XLSX.utils.sheet_to_json(sheet,{defval:''});}
  const normalized=normalizeRows(rows);
  if(!normalized.length)throw new Error('No circuit rows found. Check that the first sheet has column headings.');
  allData=normalized;initFilters();update();setStatus(`Loaded ${file.name} — ${allData.length} circuit rows. Data stays in your browser.`);$('excelFile').value='';
 }catch(err){alert('Could not read this file: '+err.message);}
}
const stories=[
 ['INTRODUCTION','Understanding Formula 1 circuits','This project explores Formula 1 circuits using data about their locations, lap counts, lap distances, directions, types and altitudes.'],
 ['GEOGRAPHICAL ANALYSIS','Where are the circuits located?','The country chart and world map show how circuits are distributed across countries and continents.'],
 ['LAP PERFORMANCE','Comparing laps and lap distances','Lap-count and lap-distance charts help compare the racing characteristics of different circuits.'],
 ['TRACK CHARACTERISTICS','Direction and circuit type','The direction and type charts compare clockwise, anti-clockwise and mixed tracks, along with permanent and street circuits.'],
 ['ALTITUDE ANALYSIS','Understanding elevation','The altitude charts compare circuit elevations and help explore whether altitude varies with lap distance.'],
 ['SCENARIO FINDINGS','Answering three key questions','The dashboard identifies the country with the most circuits, the longest lap distance and the highest-altitude circuit in the selected data.']
];
let storyIndex=0;
function renderStory(){const s=stories[storyIndex];$('storyCount').textContent=`${String(storyIndex+1).padStart(2,'0')} / ${String(stories.length).padStart(2,'0')}`;$('storyKicker').textContent=s[0];$('storyTitle').textContent=s[1];$('storyText').textContent=s[2];}
function start(){
 allData=demoData;initMap();initFilters();update();renderStory();
 ['countryFilter','directionFilter','typeFilter','searchFilter'].forEach(id=>$(id).addEventListener(id==='searchFilter'?'input':'change',update));
 $('resetFilters').addEventListener('click',()=>{['countryFilter','directionFilter','typeFilter','searchFilter'].forEach(id=>$(id).value='');update();});
 $('excelFile').addEventListener('change',e=>{if(e.target.files[0])loadExcel(e.target.files[0]);});
 $('prevStory').addEventListener('click',()=>{storyIndex=(storyIndex-1+stories.length)%stories.length;renderStory();});
 $('nextStory').addEventListener('click',()=>{storyIndex=(storyIndex+1)%stories.length;renderStory();});
}
document.addEventListener('DOMContentLoaded',start);