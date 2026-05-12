import sys
import os
import webview

HTML = """
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>발주 수량 정리</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; background: #f5f5f3; color: #1a1a1a; min-height: 100vh; }
  header { background: #fff; border-bottom: 0.5px solid #e0e0db; padding: 0 2rem; height: 56px; display: flex; align-items: center; gap: 12px; }
  header h1 { font-size: 15px; font-weight: 500; }
  header span { font-size: 12px; color: #888; }
  .container { max-width: 900px; margin: 0 auto; padding: 2rem 1.5rem; }
  .drop { border: 1.5px dashed #ccc; border-radius: 12px; padding: 2.5rem; text-align: center; cursor: pointer; background: #fff; margin-bottom: 1.5rem; }
  .drop:hover { background: #f9f9f7; border-color: #aaa; }
  .drop input[type=file] { display: none; }
  .drop-icon { font-size: 32px; margin-bottom: .5rem; }
  .drop-title { font-size: 15px; font-weight: 500; margin-bottom: 4px; }
  .drop-sub { font-size: 13px; color: #888; }
  .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(110px, 1fr)); gap: 10px; margin-bottom: 1.5rem; }
  .stat { background: #fff; border-radius: 8px; padding: .75rem 1rem; border: 0.5px solid #e8e8e4; }
  .stat-label { font-size: 11px; color: #888; margin-bottom: 4px; }
  .stat-value { font-size: 22px; font-weight: 500; }
  .tabs { display: flex; border-bottom: 0.5px solid #e0e0db; margin-bottom: 1.25rem; }
  .tab { font-size: 13px; padding: 8px 18px; border: none; background: none; cursor: pointer; color: #888; border-bottom: 2px solid transparent; margin-bottom: -1px; font-family: inherit; }
  .tab.active { color: #1a1a1a; border-bottom-color: #1a1a1a; font-weight: 500; }
  .tbl-wrap { background: #fff; border-radius: 12px; border: 0.5px solid #e8e8e4; overflow: hidden; overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; min-width: 400px; }
  th { background: #f9f9f7; font-size: 11px; font-weight: 500; color: #888; text-align: left; padding: 8px 12px; border-bottom: 0.5px solid #e8e8e4; white-space: nowrap; }
  td { padding: 7px 12px; border-bottom: 0.5px solid #f0f0ec; color: #1a1a1a; vertical-align: middle; }
  tr:last-child td { border-bottom: none; }
  .subtotal td { background: #fafaf8; font-weight: 500; }
  .grand-total td { font-weight: 500; border-top: 0.5px solid #ddd; background: #f5f5f3; }
  .center { text-align: center; }
  .muted { color: #888; }
  .tag-type { display: inline-block; font-size: 10px; padding: 2px 7px; border-radius: 4px; background: #e8f0fb; color: #2d6fc9; margin-right: 6px; vertical-align: middle; white-space: nowrap; }
  #result { display: none; }
  .open-btn { display: inline-block; margin-top: 1rem; padding: 10px 24px; background: #1a1a1a; color: #fff; border: none; border-radius: 8px; font-size: 14px; cursor: pointer; font-family: inherit; }
  .open-btn:hover { background: #333; }
</style>
</head>
<body>
<header>
  <h1>발주 수량 정리</h1>
  <span>에이블리 발주 엑셀 파일 분석기</span>
</header>
<div class="container">
  <div class="drop" id="dropZone">
    <div class="drop-icon">📂</div>
    <div class="drop-title">발주 엑셀 파일 열기</div>
    <div class="drop-sub">클릭하거나 파일을 드래그하세요 (.xlsx)</div>
    <br>
    <button class="open-btn" onclick="openFile()">파일 선택</button>
    <input type="file" id="fileInput" accept=".xlsx,.xls">
  </div>
  <div id="result">
    <div class="stats" id="statsBar"></div>
    <div class="tabs">
      <button class="tab active" onclick="switchTab(this,'design')">디자인 / 색상</button>
      <button class="tab" onclick="switchTab(this,'type')">옷종류 / 색상 / 사이즈</button>
    </div>
    <div class="tbl-wrap"><div id="tableWrap"></div></div>
  </div>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script>
let rows=[], currentTab='design';

// 파일 선택 버튼 - pywebview API 사용
async function openFile() {
  try {
    const result = await window.pywebview.api.open_file_dialog();
    if (result) loadFromBase64(result.data, result.name);
  } catch(e) {
    // fallback: input file
    document.getElementById('fileInput').click();
  }
}

// 드래그앤드롭 / input file fallback
const dropZone = document.getElementById('dropZone');
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.style.background='#f0f0ec'; });
dropZone.addEventListener('dragleave', () => { dropZone.style.background=''; });
dropZone.addEventListener('drop', e => {
  e.preventDefault(); dropZone.style.background='';
  processFile(e.dataTransfer.files[0]);
});
document.getElementById('fileInput').addEventListener('change', e => processFile(e.target.files[0]));

function processFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    const wb = XLSX.read(e.target.result, {type:'array'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(ws);
    renderAll(file.name);
  };
  reader.readAsArrayBuffer(file);
}

function loadFromBase64(b64, name) {
  const binary = atob(b64);
  const arr = new Uint8Array(binary.length);
  for (let i=0; i<binary.length; i++) arr[i] = binary.charCodeAt(i);
  const wb = XLSX.read(arr, {type:'array'});
  const ws = wb.Sheets[wb.SheetNames[0]];
  rows = XLSX.utils.sheet_to_json(ws);
  renderAll(name);
}

function splitCode(code){const c=String(code||'');const idx=c.indexOf('_');if(idx<0)return{type:c,design:c};return{type:c.slice(0,idx),design:c.slice(idx+1)};}
function is11Product(code){return String(code||'').includes('1+1');}
function parseOption(optStr,flag11){
  const str=String(optStr||'').trim();const parts=str.split('/').map(p=>p.trim());
  if(flag11&&parts.length>=6)return[{design:parts[0],color:parts[1],size:parts[2]},{design:parts[3],color:parts[4],size:parts[5]}];
  const color=(parts[0]||'').split('_')[0].trim();const size=(parts[1]||'').trim();
  return[{design:null,color,size}];
}

function renderAll(fname){
  let totalQty=0;const designs=new Set(),types=new Set();
  rows.forEach(r=>{totalQty+=Number(r['수량'])||1;designs.add(r['상품명']);types.add(splitCode(r['판매자 상품코드']).type);});
  document.getElementById('statsBar').innerHTML=`
    <div class="stat"><div class="stat-label">총 수량</div><div class="stat-value">${totalQty}</div></div>
    <div class="stat"><div class="stat-label">디자인 수</div><div class="stat-value">${designs.size}</div></div>
    <div class="stat"><div class="stat-label">옷 종류</div><div class="stat-value">${types.size}</div></div>
    <div class="stat"><div class="stat-label">주문 행</div><div class="stat-value">${rows.length}</div></div>`;
  dropZone.innerHTML=`<div class="drop-icon">✅</div><div class="drop-title">${fname}</div><div class="drop-sub">다른 파일을 열려면 클릭</div><br><button class="open-btn" onclick="openFile()">파일 선택</button><input type="file" id="fileInput" accept=".xlsx,.xls">`;
  dropZone.addEventListener('dragover',e=>{e.preventDefault();});
  dropZone.addEventListener('drop',e=>{e.preventDefault();processFile(e.dataTransfer.files[0]);});
  document.getElementById('fileInput').addEventListener('change',e=>processFile(e.target.files[0]));
  document.getElementById('result').style.display='block';
  renderTab();
}

function switchTab(btn,tab){currentTab=tab;document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));btn.classList.add('active');renderTab();}
function renderTab(){currentTab==='design'?renderDesign():renderType();}

function renderDesign(){
  const map={};
  rows.forEach(r=>{
    const qty=Number(r['수량'])||1;const flag11=is11Product(r['판매자 상품코드']);
    const{type,design:cd}=splitCode(r['판매자 상품코드']);
    parseOption(r['옵션 정보'],flag11).forEach(opt=>{
      const dn=opt.design||cd;const key=type+'|'+dn;
      if(!map[key])map[key]={type,design:dn,colors:{}};
      map[key].colors[opt.color]=(map[key].colors[opt.color]||0)+qty;
    });
  });
  let html='<table><thead><tr><th>디자인</th><th>색상</th><th style="text-align:right">수량</th></tr></thead><tbody>';
  let grand=0;
  Object.values(map).sort((a,b)=>a.type.localeCompare(b.type,'ko')||a.design.localeCompare(b.design,'ko')).forEach(({type,design,colors})=>{
    const total=Object.values(colors).reduce((a,b)=>a+b,0);grand+=total;
    const cl=Object.entries(colors);let first=true;
    cl.forEach(([color,qty])=>{
      html+=`<tr>`;
      if(first){html+=`<td rowspan="${cl.length+1}" style="font-weight:500;border-right:0.5px solid #f0f0ec"><span class="tag-type">${type}</span>${design}</td>`;first=false;}
      html+=`<td class="muted">${color}</td><td style="text-align:right">${qty}</td></tr>`;
    });
    html+=`<tr class="subtotal"><td class="muted" style="text-align:right">소계</td><td style="text-align:right">${total}</td></tr>`;
  });
  html+=`<tr class="grand-total"><td colspan="2" style="text-align:right">합계</td><td style="text-align:right">${grand}</td></tr></tbody></table>`;
  document.getElementById('tableWrap').innerHTML=html;
}

function renderType(){
  const sizeOrder=['XS','S','M','L','XL','2XL','3XL'];const map={};
  rows.forEach(r=>{
    const type=splitCode(r['판매자 상품코드']).type;const qty=Number(r['수량'])||1;
    const flag11=is11Product(r['판매자 상품코드']);
    parseOption(r['옵션 정보'],flag11).forEach(opt=>{
      if(!map[type])map[type]={};if(!map[type][opt.color])map[type][opt.color]={};
      map[type][opt.color][opt.size]=(map[type][opt.color][opt.size]||0)+qty;
    });
  });
  const allSizes=new Set();
  Object.values(map).forEach(c=>Object.values(c).forEach(s=>Object.keys(s).forEach(k=>allSizes.add(k))));
  const sizes=[...allSizes].sort((a,b)=>{const ai=sizeOrder.indexOf(a),bi=sizeOrder.indexOf(b);if(ai>=0&&bi>=0)return ai-bi;if(ai>=0)return -1;if(bi>=0)return 1;return a.localeCompare(b);});
  let html=`<table><thead><tr><th>옷 종류</th><th>색상</th>${sizes.map(s=>`<th class="center">${s}</th>`).join('')}<th style="text-align:right">소계</th></tr></thead><tbody>`;
  let grand=0;const grandSize={};
  Object.entries(map).sort((a,b)=>a[0].localeCompare(b[0],'ko')).forEach(([type,colors])=>{
    const typeTotal=Object.values(colors).reduce((s,sz)=>s+Object.values(sz).reduce((a,b)=>a+b,0),0);grand+=typeTotal;
    const typeSize={};const cl=Object.entries(colors);let first=true;
    cl.forEach(([color,szMap])=>{
      const ct=Object.values(szMap).reduce((a,b)=>a+b,0);html+=`<tr>`;
      if(first){html+=`<td rowspan="${cl.length+1}" style="font-weight:500;border-right:0.5px solid #f0f0ec">${type}</td>`;first=false;}
      html+=`<td class="muted">${color}</td>`;
      sizes.forEach(s=>{const v=szMap[s]||0;typeSize[s]=(typeSize[s]||0)+v;grandSize[s]=(grandSize[s]||0)+v;html+=`<td class="center">${v||''}</td>`;});
      html+=`<td style="text-align:right;font-weight:500">${ct}</td></tr>`;
    });
    html+=`<tr class="subtotal"><td class="muted" style="text-align:right">소계</td>${sizes.map(s=>`<td class="center">${typeSize[s]||''}</td>`).join('')}<td style="text-align:right">${typeTotal}</td></tr>`;
  });
  html+=`<tr class="grand-total"><td colspan="2" style="text-align:right">합계</td>${sizes.map(s=>`<td class="center">${grandSize[s]||''}</td>`).join('')}<td style="text-align:right">${grand}</td></tr></tbody></table>`;
  document.getElementById('tableWrap').innerHTML=html;
}
</script>
</body>
</html>
"""

class Api:
    def open_file_dialog(self):
        import base64
        result = window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=('Excel Files (*.xlsx;*.xls)',)
        )
        if result and len(result) > 0:
            path = result[0]
            with open(path, 'rb') as f:
                data = base64.b64encode(f.read()).decode('utf-8')
            return {'data': data, 'name': os.path.basename(path)}
        return None

api = Api()
window = webview.create_window(
    '발주 수량 정리',
    html=HTML,
    js_api=api,
    width=1000,
    height=750,
    min_size=(700, 500)
)

if __name__ == '__main__':
    webview.start()
