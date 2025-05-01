const CLIENT_ID = '41469672863-oirobf089qejmd3mscdbe61l6d8bdbqi.apps.googleusercontent.com';
const API_KEY = 'AIzaSyC4RNpF79s6A9m9e0bfcErhuYL-M8T5skQ';

const DISCOVERY_DOCS = [
  'https://sheets.googleapis.com/$discovery/rest?version=v4',
];

const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
];

let tokenClient;

function gapiLoaded() {
  gapi.load('client', initializeGapiClient);
}

async function initializeGapiClient() {
  await gapi.client.init({
    apiKey: API_KEY,
    discoveryDocs: DISCOVERY_DOCS,
  });
  handleAuth();
}

function handleAuth() {
  tokenClient.callback = async (resp) => {
    if (resp.error !== undefined) {
      throw (resp);
    }
    vbardh.style.visibility = 'visible';
    vzi.style.visibility = 'visible';
    mverdh.style.visibility = 'visible';
  };
  if (gapi.client.getToken() === null) {
    tokenClient.requestAccessToken({prompt: 'consent'});
  } else {
    tokenClient.requestAccessToken({prompt: ''});
  }
}

function gisLoaded() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES.join(' '),
    callback: '',
  });
}



const Bardh='1tD-nH-r2rX8iaUQp5U9GRkBR9T9HS-jkHLnI9Vf8GXc';
const Zi='1dxolaI00sOo-D8Vt4PTJ-2V1EtifML3O5uCnDFKxd9A';
const Verdh='1PT85vtxPRIWLf27eIhZgkiidvxBYN04EsUzG8LlYml8';
let active_sheet;
let tbody;
let muaji;
let shoferet;
let gropat;
let nafta;
let rruget;
let serviset;
let shkarkimet;
let placeholder;
const months = { JANAR: 31, SHKURT: 28, MARS: 31, PRILL: 30, MAJ: 31, QERSHOR: 30, KORRIK: 31, GUSHT: 31, SHTATOR: 30, TETOR: 31, NENTOR: 30, DHJETOR: 31 };
let shobj={};
let gobj={};
let shaobj={};
let active_months={};

const vbardh = document.getElementById('bardh');
const vzi = document.getElementById('zi');
const mverdh = document.getElementById('verdh');
const header = document.getElementById('header');

vbardh.style.visibility = 'hidden';
vzi.style.visibility = 'hidden';
mverdh.style.visibility = 'hidden';
header.style.visibility = 'hidden';

vbardh.addEventListener('click',async function() {
  header.style.visibility = 'visible';
  header.innerHTML = 'Volvo Bardh';
  let x = document.getElementById('container');
  x.innerHTML = "";
  var htmlContent = `
  <div  id="table-container">
    <table class="general_table" id="myTable">
      <thead>
        <tr>
          <th> Muajt Volvo Bardh </th>
        </tr>
      </thead>
      <tbody></tbody> <!-- Added tbody element -->
    </table>
  </div>
  <div class="form-container">
  <select id="ZgjithMaujin">
  <option value="jo" selected> Muaji </option>
  <option value="JANAR"> JANAR </option>
  <option value="SHKURT"> SHKURT </option>
  <option value="MARS"> MARS </option>
  <option value="PRILL"> PRILL </option>
  <option value="MAJ"> MAJ </option>
  <option value="QERSHOR"> QERSHOR </option>
  <option value="KORRIK"> KORRIK </option>
  <option value="GUSHT"> GUSHT </option>
  <option value="SHTATOR"> SHTATOR </option>
  <option value="TETOR"> TETOR </option>
  <option value="NENTOR"> NENTOR </option>
  <option value="DHJETOR"> DHJETOR </option>
  </select>
    <button id="shto">SHTO</button>
  </div>
  `;
  x.innerHTML = htmlContent;
  active_sheet = Bardh;
  nafta=null;
  shoferet=null;
  gropat=null;
  rruget=null;
  serviset=null;
  shkarkimet=null;
  placeholder=null;
  shobj={};
  gobj={};
  shaobj={};
  active_months={};
  
  let table = document.getElementById('myTable');
  
  tbody = table.querySelector('tbody');

  await populateMonth()
  
  await shikoEvent();

  let button = document.getElementById('shto');

  button.addEventListener('click', async function() {
    try{
      let muajiInput = document.getElementById('ZgjithMaujin');
      let y = muajiInput.value;
      if(active_months.hasOwnProperty(y)){alert("Muaji ekziston!");return;}
      if(y!="jo"){
        let currentDate = new Date();
        let currentYearUTC = currentDate.getUTCFullYear();
        y=currentYearUTC+"/"+y;
        await appendValues(active_sheet, 'Sheet1', 'USER_ENTERED', [[y],]);
        alert("Muaji u shtua.");
        muajiInput.selectedIndex=0;
        populateMonth();
      }
      else{
        alert("Zgjidh muajin.")
      }
    }
    catch(err){
      alert("Muaji nuk u shtua bej refresh.Po su rregullua coj mesazh Klajdit");
    } 
    
  });
});

vzi.addEventListener('click',async function() {
  header.style.visibility = 'visible';
  header.innerHTML = 'Volvo Zeze';
  let x = document.getElementById('container');
  x.innerHTML = "";
  var htmlContent = `
  <div  id="table-container">
    <table class="general_table" id="myTable">
      <thead>
        <tr>
          <th> Muajt Volvo Zeze </th>
        </tr>
      </thead>
      <tbody></tbody> <!-- Added tbody element -->
    </table>
  </div>
  <div class="form-container">
  <select id="ZgjithMaujin">
  <option value="jo" selected> Muaji </option>
  <option value="JANAR"> JANAR </option>
  <option value="SHKURT"> SHKURT </option>
  <option value="MARS"> MARS </option>
  <option value="PRILL"> PRILL </option>
  <option value="MAJ"> MAJ </option>
  <option value="QERSHOR"> QERSHOR </option>
  <option value="KORRIK"> KORRIK </option>
  <option value="GUSHT"> GUSHT </option>
  <option value="SHTATOR"> SHTATOR </option>
  <option value="TETOR"> TETOR </option>
  <option value="NENTOR"> NENTOR </option>
  <option value="DHJETOR"> DHJETOR </option>
  </select>
    <button id="shto">SHTO</button>
  </div>
  `;
  x.innerHTML = htmlContent;
  active_sheet = Zi;
  nafta=null;
  shoferet=null;
  gropat=null;
  rruget=null;
  serviset=null;
  placeholder=null;
  shkarkimet=null;
  shobj={};
  gobj={};
  shaobj={};
  active_months={};
  
  let table = document.getElementById('myTable');
  
  tbody = table.querySelector('tbody');

  await populateMonth()
  
  await shikoEvent();

  let button = document.getElementById('shto');

  button.addEventListener('click', async function() {
    try{
      let muajiInput = document.getElementById('ZgjithMaujin');
      let y = muajiInput.value;
      if(active_months.hasOwnProperty(y)){alert("Muaji ekziston!");return;}
      if(y!="jo"){
        let currentDate = new Date();
        let currentYearUTC = currentDate.getUTCFullYear();
        y=currentYearUTC+"/"+y;
        await appendValues(active_sheet, 'Sheet1', 'USER_ENTERED', [[y],]);
        alert("Muaji u shtua.");
        muajiInput.selectedIndex=0;
        populateMonth();
      }
      else{
        alert("Shkruaj muajin.")
      }
    }
    catch(err){
      alert("Muaji nuk u shtua bej refresh.Po su rregullua coj mesazh Klajdit");
    } 
    
  });
});

mverdh.addEventListener('click',async function() {
  header.style.visibility = 'visible';
  header.innerHTML = 'Mani Verdh';
  let x = document.getElementById('container');
  x.innerHTML = "";
  var htmlContent = `
  <div  id="table-container">
    <table class="general_table" id="myTable">
      <thead>
        <tr>
          <th> Muajt Mani Verdh </th>
        </tr>
      </thead>
      <tbody></tbody> <!-- Added tbody element -->
    </table>
  </div>
  <div class="form-container">
  <select id="ZgjithMaujin">
  <option value="jo" selected> Muaji </option>
  <option value="JANAR"> JANAR </option>
  <option value="SHKURT"> SHKURT </option>
  <option value="MARS"> MARS </option>
  <option value="PRILL"> PRILL </option>
  <option value="MAJ"> MAJ </option>
  <option value="QERSHOR"> QERSHOR </option>
  <option value="KORRIK"> KORRIK </option>
  <option value="GUSHT"> GUSHT </option>
  <option value="SHTATOR"> SHTATOR </option>
  <option value="TETOR"> TETOR </option>
  <option value="NENTOR"> NENTOR </option>
  <option value="DHJETOR"> DHJETOR </option>
  </select>
    <button id="shto">SHTO</button>
  </div>
  `;
  x.innerHTML = htmlContent;
  active_sheet = Verdh;
  nafta=null;
  shoferet=null;
  gropat=null;
  rruget=null;
  serviset=null;
  shkarkimet=null;
  placeholder=null;
  shobj={};
  gobj={};
  shaobj={};
  active_months={};
  
  let table = document.getElementById('myTable');
  
  tbody = table.querySelector('tbody');

  await populateMonth()
  
  await shikoEvent();

  let button = document.getElementById('shto');

  button.addEventListener('click', async function() {
    try{
      let muajiInput = document.getElementById('ZgjithMaujin');
      let y = muajiInput.value;
      if(active_months.hasOwnProperty(y)){alert("Muaji ekziston!");return;}
      if(y!="jo"){
        let currentDate = new Date();
        let currentYearUTC = currentDate.getUTCFullYear();
        y=currentYearUTC+"/"+y;
        await appendValues(active_sheet, 'Sheet1', 'USER_ENTERED', [[y],]);
        alert("Muaji u shtua.");
        muajiInput.selectedIndex=0;
        populateMonth();
      }
      else{
        alert("Zgjidh muajin.")
      }
    }
    catch(err){
      alert("Muaji nuk u shtua bej refresh.Po su rregullua coj mesazh Klajdit");
    } 
    
  });
});

async function appendValues(spreadsheetId, range, valueInputOption, _values) {
  let values = [[],];
  values = _values;
  const body = {
  values: values,
  };
  await gapi.client.sheets.spreadsheets.values.append({
  spreadsheetId: spreadsheetId,
  range: range,
  valueInputOption: valueInputOption,
  resource: body,
  });
}

async function batchGetValues(spreadsheetId, _ranges, condition,type) {
  let ranges = _ranges;
  let result = await gapi.client.sheets.spreadsheets.values.batchGet({
    spreadsheetId: spreadsheetId,
    ranges: ranges,
  });

  let rows = result.result.valueRanges.map(valueRange => valueRange.values);
  if(!condition){
    return rows;}
  else{
    try{
      if(ranges.length==6){
        if(rows[0]){shoferet=rows[0].filter(cell => cell[0].includes(condition)).map(cell => cell[1])};
        if(rows[1]){gropat=rows[1].filter(cell => cell[0].includes(condition)).map(cell => cell[1])};
        if(rows[2]){nafta=rows[2].filter(cell => cell[0].includes(condition)).map(cell => [cell[1],cell[2]])};
        if(rows[3]){rruget=rows[3].filter(cell => cell[0].includes(condition)).map(cell => [cell[1],cell[2],cell[3],cell[4],cell[5],cell[6]])};
        if(rows[4]){serviset=rows[4].filter(cell => cell[0].includes(condition)).map(cell => [cell[1],cell[2]])};
        if(rows[5]){shkarkimet=rows[5].filter(cell => cell[0].includes(condition)).map(cell => cell[1])};
      }
      else if(type=="sh"){
        if(rows[0]){shoferet=rows[0].filter(cell => cell[0].includes(condition)).map(cell => cell[1]);}
      }
      else if (type=="g"){
        if(rows[0]){gropat=rows[0].filter(cell => cell[0].includes(condition)).map(cell => cell[1]);}
      }
      else if(type=="n"){
        if(rows[0]){nafta=rows[0].filter(cell => cell[0].includes(condition)).map(cell => [cell[1],cell[2]]);};
      }
      else if(type=="rr"){
        if(rows[0]){rruget=rows[0].filter(cell => cell[0].includes(condition)).map(cell => [cell[1],cell[2],cell[3],cell[4],cell[5],cell[6]])};
      }
      else if(type=="se"){
        if(rows[0]){serviset=rows[0].filter(cell => cell[0].includes(condition)).map(cell => [cell[1],cell[2]])};
      }
      else if(type=="sha"){
        if(rows[0]){shkarkimet=rows[0].filter(cell => cell[0].includes(condition)).map(cell => cell[1])};
      }
      
    }
    catch(err){
      alert("Ndodhi nje problem bej refresh po su rregullua mer klajdin.")
    }
   }
}

async function populateMonth(){
  let result
  try{
    result = await batchGetValues(active_sheet, ['Sheet1!A2:A37'],null,null);
    tbody.innerHTML = "";
    for (let i of result[0]) {
      let currentDate = new Date();
      let currentYearUTC = currentDate.getUTCFullYear();
      let strarr = i[0].split("/");
      if (strarr[0] == `${currentYearUTC}`) {
          active_months[strarr[1]] = "";
      }
      const newTableRow = tbody.insertRow();
      newTableRow.insertCell(0).innerHTML = '<button id='+i[0]+' class="shiko-button">'+i[0]+'</button>';
    }
  }
  catch(err){
    if(!result){
      alert("Ndodhi nje problem. Bej revresh po su rregullua thuaj Klajdit.");
    }
    else{
      alert("Nuk ka asnje muaj pun.");
    }
  }  
}

async function shikoEvent(){
  tbody.addEventListener('click',async function(event){
    if (event.target.classList.contains('shiko-button')) {
      const button = event.target;
      muaji=button.id;
      header.innerHTML+='/'+muaji;
      placeholder=header.innerHTML;
      load_Navigation();
      try{
        await batchGetValues(active_sheet,["Sheet2!A2:B150", "Sheet3!A2:B200","Sheet4!A2:C500","Sheet5!A2:G1600","Sheet6!A2:C150","Sheet7!A2:B200"],muaji);
        calculate_Values();
        NavigationEvent();
      }
      catch(err){
        alert("Ndodhi nje problem. Bej refresh nese nuk rregullohet thuaj Klajdit.")
      }
    } 
  });
}

function load_Navigation(){
  let x = document.getElementById('container');
  x.innerHTML = "";
  var htmlContent = `
  <button id="sh" class="navigation">SHOFERAT</button>
  <button id="rr" class="navigation">RRUGET</button>
  <button id="g" class="navigation">GROPAT</button>
  <button id="n" class="navigation">NAFTA</button>
  <button id="se" class="navigation">SERVISET</button>
  <button id="shka" class="navigation">SHKARKIMET</button>`;
  
  x.innerHTML =htmlContent ;
}

async function NavigationEvent(){
  let sh=document.getElementById('sh');
  let rr=document.getElementById('rr');
  let g=document.getElementById('g');
  let n=document.getElementById('n');
  let se=document.getElementById('se');
  let shkark=document.getElementById('shka');

  sh.addEventListener('click',function(){
  load_Sh_Or_G(shoferet,
    `<th> Shoferi </th>
    <th> Rruget Naten</th>
    <th> Rruget Diten </th>`,"shshofer");
    header.innerHTML+="/SHOFERET";
});

  g.addEventListener('click',function(){
  load_Sh_Or_G(gropat,
    `<th> Gropa </th>
    <th> Rruget </th>`,"shgrop");
    header.innerHTML+="/GROPAT";
  }); 
  
  n.addEventListener('click',function(){
    load_N();
    header.innerHTML+="/NAFTA";});
  
  rr.addEventListener('click',function(){
    load_Rruget();
    header.innerHTML+="/RRUGET";});

  se.addEventListener('click',function(){
    load_Serviset();
    header.innerHTML+="/SERVISET";}); 
  
  shkark.addEventListener('click',function(){
    load_Sh_Or_G(shkarkimet,
      `<th> Vend Shkarkimet </th>
      <th> Rruget </th>`,"shshkarkim");
      header.innerHTML+="/SHKARKIMET";
    });
}

function load_Serviset(){
  let x = document.getElementById('container');
  x.innerHTML = "";
  var htmlContent = `
  <button id="mbrapa"> Kthehu Mbrapa </button>
  <div id="table-container">
    <table class="general_table" id="myTable">
      <thead>
        <tr>
        <th> Servisi </th>
        <th> Cmimi </th>
        </tr>
      </thead>
      <tbody></tbody> <!-- Added tbody element -->
    </table>
  </div>
  <div class="form-container">
    <input id="info" type="text" placeholder="Shkruaj Servisin...">
    <input id="shuma" type="text" placeholder="Shkruaj Cmimin...">
    <button id="shservis">SHTO</button>
  </div>
  `;
  x.innerHTML=htmlContent;
  shto_Evente("shservis");
  if(!serviset){return}
  let table = document.getElementById('myTable');
  table= table.querySelector('tbody');
  table .innerHTML = "";
  let sum=0
  let newTableRow
  for(i of serviset){
    newTableRow = table.insertRow();
    newTableRow.insertCell(0).innerHTML = i[0];
    newTableRow.insertCell(1).innerHTML = i[1];
    sum+=parseInt(i[1]);
  }
  if(sum!=0){newTableRow.insertCell(2).innerHTML =`Totali: ${sum} Leke`;}
}

function load_N(){
  let x = document.getElementById('container');
  x.innerHTML = "";
  var htmlContent = `
  <button id="mbrapa"> Kthehu Mbrapa </button>
  <div id="table-container">
    <table class="general_table" id="myTable">
      <thead>
        <tr>
        <th> Data </th>
        <th> Sasia </th>
        </tr>
      </thead>
      <tbody></tbody> <!-- Added tbody element -->
    </table>
  </div>
  <div class="form-container">
  <select id="info" >
  <option value="jo" selected> Data </option>
  </select>
    <input id="naft" type="text" placeholder="Shkruaj Sasin...">
    <button id="shnaft">SHTO</button>
  </div>
  `;
  x.innerHTML=htmlContent;
  let nd = document.getElementById('info');
  let m=header.innerHTML.split("/")
  for(let i=1;i<=months[m[2]];i++){
    nd.innerHTML+=`<option value="${i}"> ${i} </option>`;
  }
  shto_Evente("shnaft");
  if(!nafta){return;}
  let table = document.getElementById('myTable');
  table= table.querySelector('tbody');
  table .innerHTML = "";
  let sum=0
  let newTableRow
  for(i of nafta){
    newTableRow = table.insertRow();
    newTableRow.insertCell(0).innerHTML = i[0];
    newTableRow.insertCell(1).innerHTML = i[1];
    sum+=parseInt(i[1]);
  }
  if(sum!=0){newTableRow.insertCell(2).innerHTML =`Totali: ${sum} Litra`;}
}

function load_Rruget(){
  let x = document.getElementById('container');
  let shoferets,gropats,shkarkimets;
  if(shoferet){
    for(let i of shoferet){
    shoferets+=`<option value="${i}">${i}</option>`;
    }}
  if(gropat){
    for(let i of gropat){
    gropats+=`<option value="${i}">${i}</option>`;
    }}
  if(shkarkimet){
    for(let i of shkarkimet){
      shkarkimets+=`<option value="${i}">${i}</option>`;
    }}
  x.innerHTML = "";
  var htmlContent = `
  <button id="mbrapa"> Kthehu Mbrapa </button>
  <div id="table-container">
    <table class="general_table" id="myTable">
      <thead>
        <tr>
        <th> Data </th>
        <th> Shoferi </th>
        <th> Gropa </th>
        <th> Shkarkuar </th>
        <th> Koha </th>
        <th> Rruget </th>
        </tr>
      </thead>
      <tbody></tbody> <!-- Added tbody element -->
    </table>
  </div>
  <div class="form-container1">
    <select id="data" >
    <option value="jo" selected> Data </option>
    </select>
    <select id="shoferi">
    <option value="jo" selected> Shoferi </option>
    ${shoferets}
    </select>
    <select id="gropa">
    <option value="jo" selected> Gropa </option>
    ${gropats}
    </select>
    <select id="shkarkimet">
    <option value="jo" selected> Shkarkimet </option>
    ${shkarkimets}
    </select>
    <select id="koha">
    <option value="jo" selected> Koha </option>
    <option value="Diten"> Diten </option>
    <option value="Naten"> Naten </option>
    </select>
    <select id="rruget">
    <option value="jo" selected> Rruget </option>
    </select>
    <button id="shrrug">SHTO</button>
  </div>
  `;
  x.innerHTML=htmlContent;
  let d = document.getElementById('data');
  let r = document.getElementById('rruget');
  let m=header.innerHTML.split("/")
  for(let i=1;i<=months[m[2]];i++){
    d.innerHTML+=`<option value="${i}"> ${i} </option>`;
    r.innerHTML+=`<option value="${i}"> ${i} </option>`;
  }
  shto_Evente("shrrug");
  if(!rruget){return}
  let table = document.getElementById('myTable');
  table= table.querySelector('tbody');
  table .innerHTML = "";
  for(i of rruget){
    const newTableRow = table.insertRow();
    newTableRow.insertCell(0).innerHTML = i[0];
    newTableRow.insertCell(1).innerHTML = i[1];
    newTableRow.insertCell(2).innerHTML = i[2];
    newTableRow.insertCell(3).innerHTML = i[3];
    newTableRow.insertCell(4).innerHTML = i[4];
    newTableRow.insertCell(5).innerHTML = i[5];
  }
}

function load_Sh_Or_G(arr,head,id){
  let x = document.getElementById('container');
  x.innerHTML = "";
  var htmlContent = `
  <button id="mbrapa"> Kthehu Mbrapa </button>
  <div id="table-container">
    <table class="general_table" id="myTable">
      <thead>
        <tr>`
        +head+
        `</tr>
      </thead>
      <tbody></tbody> <!-- Added tbody element -->
    </table>
  </div>
  <div class="form-container">
    <input id="info" type="text" placeholder="Shto...">
    <button id="`+id+`">SHTO</button>
  </div>`;
  x.innerHTML=htmlContent;
  let table = document.getElementById('myTable');
  table= table.querySelector('tbody');
  table .innerHTML = "";
  shto_Evente(id);
  if(!arr){return};
    for (let i of arr) {
      const newTableRow = table.insertRow();
      if(id=="shshofer"){
        newTableRow.insertCell(0).innerHTML = i;
        newTableRow.insertCell(1).innerHTML = shobj[i].Naten;
        newTableRow.insertCell(2).innerHTML = shobj[i].Diten;
      }
      else if(id=="shgrop"){
        newTableRow.insertCell(0).innerHTML = i;
        newTableRow.insertCell(1).innerHTML = gobj[i];
      }
      else if(id=="shshkarkim"){
        newTableRow.insertCell(0).innerHTML = i;
        newTableRow.insertCell(1).innerHTML = shaobj[i];
      }
    }
}

function  shto_Evente(type){
  let shshofer=document.getElementById('shshofer');
  let shgrop=document.getElementById('shgrop');
  let shnaft=document.getElementById('shnaft');
  let shrrug=document.getElementById('shrrug');
  let shservis=document.getElementById('shservis');
  let shshkarkim=document.getElementById('shshkarkim');
  let mrapa=document.getElementById('mbrapa');

  mrapa.addEventListener('click',function(){
    header.innerHTML=placeholder;
    load_Navigation();
    NavigationEvent();
  })

  if(type=="shshofer"){
    shshofer.addEventListener('click',async function(){
      let info=document.getElementById('info');
      info=info.value;
      try{
        if(info){
          await appendValues(active_sheet, 'Sheet2', 'USER_ENTERED', [[muaji,info],]);
          alert("Shoferi u shtua.");
          await batchGetValues(active_sheet, ['Sheet2!A2:B150'],muaji,"sh");
          calculate_Values();
          load_Sh_Or_G(shoferet,
            `<th> Shoferi </th>
            <th> Rr Naten</th>
            <th> Rr Diten </th>`,"shshofer")
          }
          else{
            alert("Shkruaj shoferin!")
          }
      }
      catch(err){
        alert("Ndodhi nje problem bej refresh po su rregullua mer klajdin.")
      }
    })
  }

  else if(type=="shgrop"){
    shgrop.addEventListener('click',async function(){
      let info=document.getElementById('info');
      info=info.value;
      try{
        if(info){
          await appendValues(active_sheet, 'Sheet3', 'USER_ENTERED', [[muaji,info],]);
          alert("Gropa u shtua.");
          await batchGetValues(active_sheet, ['Sheet3!A2:B200'],muaji,"g");
          calculate_Values();
          load_Sh_Or_G(gropat,
            `<th> Gropa </th>
            <th> Rruget </th>`,"shgrop")
          }
          else{
            alert("Shkruaj gropen!")
          }
      }
      catch(err){
        alert("Ndodhi nje problem bej refresh po su rregullua mer klajdin.")
      }
    })
  }
  else if(type=="shnaft"){
    shnaft.addEventListener('click',async function(){
      let info=document.getElementById('info');
      let info1=document.getElementById('naft');
      info=info.value;
      info1=info1.value;
      try{
        if(info!="jo" && info1 && parseInt(info1)){
          await appendValues(active_sheet, 'Sheet4', 'USER_ENTERED', [[muaji,info,info1],]);
          alert("Nafta u shtua.");
          await batchGetValues(active_sheet, ['Sheet4!A2:C500'],muaji,"n");
          load_N()
        }
          else{
            alert("Ploteso te dhenat sic duhet.")
          }
      }
      catch(err){
        alert("Ndodhi nje problem bej refresh po su rregullua mer klajdin.")
      }
    })
  }
  else if(type=="shrrug"){
    shrrug.addEventListener('click',async function(){
      let data=document.getElementById('data');
      let nr=document.getElementById('rruget');
      let ko=document.getElementById('koha');
      let sho=document.getElementById('shoferi');
      let gro=document.getElementById('gropa');
      let shk=document.getElementById('shkarkimet');
      data=data.value;
      nr=nr.value;
      ko=ko.value;
      sho=sho.value;
      gro=gro.value;
      shk=shk.value
      try{
        if(data!="jo" && nr!="jo" && ko!="jo" && sho!="jo" && gro!="jo" && shk!="jo"){
          await appendValues(active_sheet, 'Sheet5', 'USER_ENTERED', [[muaji,data,sho,gro,shk,ko,nr],]);
          alert("Rruget u shtuan.");
          await batchGetValues(active_sheet, ['Sheet5!A2:G1600'],muaji,"rr");
          load_Rruget()
          calculate_Values()
        }
          else{
            alert("Ploteso te dhenat")
          }
      }
      catch(err){
        alert("Ndodhi nje problem bej refresh po su rregullua mer klajdin.")
      }
    })
  }
  else if(type=="shservis"){
    shservis.addEventListener('click',async function(){
      let info=document.getElementById('info');
      let info1=document.getElementById('shuma');
      info=info.value;
      info1=info1.value;
      try{
        if(info && info1 && parseInt(info1)){
          await appendValues(active_sheet, 'Sheet6', 'USER_ENTERED', [[muaji,info,info1],]);
          alert("Servisi u shtua.");
          await batchGetValues(active_sheet, ['Sheet6!A2:C150'],muaji,"se");
          load_Serviset();
        }
          else{
            alert("Ploteso te dhenat sic duhet.")
          }
      }
      catch(err){
        alert("Ndodhi nje problem bej refresh po su rregullua mer klajdin.")
      }
    })
  }

  else if(type=="shshkarkim"){
    shshkarkim.addEventListener('click',async function(){
      let info=document.getElementById('info');
      info=info.value;
      try{
        if(info){
          await appendValues(active_sheet, 'Sheet7', 'USER_ENTERED', [[muaji,info],]);
          alert("Vend Shkarkimi u shtua.");
          await batchGetValues(active_sheet, ['Sheet7!A2:B200'],muaji,"sha");
          calculate_Values();
          load_Sh_Or_G(shkarkimet,
            `<th> Vend Shkarkimet </th>
            <th> Rruget </th>`,"shshkarkim");}
          else{
            alert("Shkruaj Vend Shkarkimin!")
          }
      }
      catch(err){
        alert("Ndodhi nje problem bej refresh po su rregullua mer klajdin.")
      }
    })
  }
}

function calculate_Values(){
  shobj={};
  gobj={};
  shaobj={};
  if(shoferet){for(let i of shoferet){shobj[i]={};shobj[i].Naten=0;shobj[i].Diten=0;}}
  if(gropat){for(let i of gropat){gobj[i]=0;}}
  if(shkarkimet){for(let i of shkarkimet){shaobj[i]=0;}}
  if (rruget) {
    for (let i of rruget) {
        if (i[4] === "Diten") {
            shobj[i[1]].Diten += parseInt(i[5]);
        } else {
            shobj[i[1]].Naten +=parseInt(i[5]);
        }
          gobj[i[2]] += parseInt(i[5]);
          shaobj[i[3]] += parseInt(i[5]);
    }
}
}
