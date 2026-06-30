const express = require('express');
const multer  = require('multer');
const xlsx    = require('xlsx');
const cors    = require('cors');
const path    = require('path');
const fs      = require('fs');

const app       = express();
const PORT      = 3000;
const XLSX_PATH = path.join(__dirname, 'DADOSTESTE.xlsx');
const UPLOADS   = path.join(__dirname, 'uploads');
const META_PATH = path.join(__dirname, 'documentos_meta.json');

app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));
app.use('/uploads', express.static(UPLOADS));

// ── Cursos por cargo ─────────────────────────────────────────
const CURSOS_CARGO = {
  'ELETRICISTA DE REDE':  ['ASO','NR_10','NR_35','NR_10SEP','POP00','DIRECAO_DEF','CNH'],
  'AUXILIAR ELETRICISTA': ['ASO','NR_10','NR_35','NR_10SEP','POP00','DIRECAO_DEF','CNH'],
  'ELETRICISTA PODADOR':  ['ASO','NR_10','NR_35','NR_10SEP','POP00','DIRECAO_DEF','PODADOR','CNH'],
  'AUX DE TRANSPORTE':    ['ASO','NR_10','NR_35','NR_10SEP','POP00','DIRECAO_DEF','CNH'],
  'AUX ADM I':            ['ASO','DIRECAO_DEF','OFFICE','5S','ELETROTECNICO','CNH'],
  'SUPERVISOR DE OBRAS':  ['ASO','NR_10','NR_35','POP00','DIRECAO_DEF','CNH','OPERADOR','ELETROTECNICO'],
  'TEC FECHAMENTO I':     ['ASO','NR_10','NR_35','DIRECAO_DEF','OFFICE','5S','ELETROTECNICO','CNH'],
  'MOT OPERADOR':         ['ASO','NR_10','NR_12','NR_35','NR_10SEP','POP00','DIRECAO_DEF','OPERADOR','CNH'],
  'OPERADOR RETRO':       ['ASO','NR_10','NR_12','NR_35','NR_10SEP','POP00','DIRECAO_DEF','RETRO','CNH'],
};

// ── Validade ──────────────────────────────────────────────────
// 'cnh' = data já é vencimento | null = sem validade | número = anos
const VALIDADE = {
  ASO:1, NR_10:2, NR_12:2, NR_35:2, NR_10SEP:2, POP00:2,
  DIRECAO_DEF:2, OPERADOR:2, RETRO:2, PODADOR:2,
  CNH:'cnh', OFFICE:null, '5S':null, ELETROTECNICO:null,
};

const TODOS_CURSOS = ['ASO','NR_10','NR_12','NR_35','NR_10SEP','POP00',
                      'DIRECAO_DEF','OPERADOR','RETRO','CNH','PODADOR',
                      'OFFICE','5S','ELETROTECNICO'];

const COL_MAP = {
  ASO:'ASO', NR_10:'NR_10', NR_12:'NR_12', NR_35:'NR_35',
  NR_10SEP:'NR_10SEP', POP00:'POP00', 'DIREÇÃO_DEF':'DIRECAO_DEF',
  OPERADOR:'OPERADOR', RETRO:'RETRO', CNH:'CNH',
  PODADOR:'PODADOR', OFFICE:'OFFICE', '5S':'5S', ELETROTECNICO:'ELETROTECNICO',
};
const COL_REVERSE = Object.fromEntries(Object.entries(COL_MAP).map(([k,v])=>[v,k]));

// ── Meta helpers ──────────────────────────────────────────────
function lerMeta() {
  if (!fs.existsSync(META_PATH)) return {};
  try { return JSON.parse(fs.readFileSync(META_PATH,'utf8')); } catch { return {}; }
}
function salvarMeta(meta) {
  fs.writeFileSync(META_PATH, JSON.stringify(meta, null, 2));
}

// ── Planilha helpers ──────────────────────────────────────────
function lerPlanilha() {
  const wb   = xlsx.readFile(XLSX_PATH, { cellDates: true });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(ws, { raw: true, cellDates: true, defval: null });
  return { wb, ws, rows };
}

function toDateStr(val) {
  if (!val && val !== 0) return null;
  if (val instanceof Date && !isNaN(val)) return val.toISOString().slice(0,10);
  if (typeof val === 'number') {
    const d = xlsx.SSF.parse_date_code(val);
    if (d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }
  const s = String(val).trim();
  if (!s || s==='null') return null;
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);
  if (m) return `${2000+parseInt(m[3])}-${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}`;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return null;
}

function calcVencimento(emissao, key) {
  const val = VALIDADE[key];
  if (!emissao) return null;
  if (val==='cnh' || val===null) return emissao;
  const d = new Date(emissao+'T00:00:00');
  d.setFullYear(d.getFullYear()+val);
  return d.toISOString().slice(0,10);
}

function statusCurso(vencimento, key) {
  if (!vencimento) return 'sem_data';
  if (VALIDADE[key]===null) return 'apto';
  const venc = new Date(vencimento+'T00:00:00');
  const hoje = new Date(); hoje.setHours(0,0,0,0);
  const d30  = new Date(hoje); d30.setDate(d30.getDate()+30);
  if (venc < hoje)  return 'vencido';
  if (venc <= d30)  return 'a_vencer';
  return 'apto';
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

async function writeXlsxSafe(wb, filePath, retries=5) {
  const tmp = filePath+'.tmp';
  for (let i=0; i<retries; i++) {
    try {
      xlsx.writeFile(wb, tmp, { bookType: 'xlsx' });
      if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
      fs.renameSync(tmp, filePath);
      return;
    } catch(e) {
      try { if (fs.existsSync(tmp)) fs.unlinkSync(tmp); } catch {}
      if ((e.code==='EBUSY'||e.code==='EACCES'||e.code==='EPERM') && i<retries-1) {
        await sleep(300*(i+1)); continue;
      }
      throw e;
    }
  }
}

// ── GET /api/colaboradores ────────────────────────────────────
app.get('/api/colaboradores', (req, res) => {
  try {
    const { rows } = lerPlanilha();
    const meta     = lerMeta();
    let totVencido=0, totAVencer=0, totApto=0, totAceito=0, totRejeitado=0, totSemAnexo=0;

    const colaboradores = rows.map(row => {
      const fkey  = Object.keys(row).find(k=>k.toUpperCase().includes('FUN'));
      const funcao= fkey ? String(row[fkey]||'').trim().toUpperCase() : '';
      const cursosExigidos = CURSOS_CARGO[funcao]||[];
      const matStr = String(row['MAT']).trim();
      const metaCol= meta[matStr]||{};
      const cursos = {};

      TODOS_CURSOS.forEach(key => {
        const colExcel = COL_REVERSE[key];
        const emissao  = toDateStr(colExcel ? row[colExcel] : null);
        const vencimento = calcVencimento(emissao, key);
        const exigido  = cursosExigidos.includes(key);
        const status   = statusCurso(vencimento, key);
        const m        = metaCol[key]||{};

        // Contagem — apenas cursos exigidos
        if (exigido) {
          if (status==='vencido')   totVencido++;
          else if (status==='a_vencer')  totAVencer++;
          else if (status==='apto')      totApto++;
        }
        if (m.validado) {
          if (m.apto) totAceito++; else totRejeitado++;
        }

        // PDFs existentes
        const pdfDir = path.join(UPLOADS, matStr);
        let pdfUrl=null, pdfName=null;
        if (fs.existsSync(pdfDir)) {
          const f = fs.readdirSync(pdfDir).find(f=>f.startsWith(key+'_')&&f.endsWith('.pdf'));
          if (f) { pdfUrl=`/uploads/${matStr}/${f}`; pdfName=f; }
        }

        if (exigido && emissao && !pdfUrl) totSemAnexo++;

        cursos[key] = {
          emissao, vencimento,
          semValidade: VALIDADE[key]===null,
          status: emissao ? status : 'sem_data',
          exigido,
          pdfUrl, pdfName,
          uploadedBy:  m.uploadedBy||null,
          uploadedAt:  m.uploadedAt||null,
          validado:    m.validado||false,
          apto:        m.apto??null,
          obs:         m.obs||'',
          validadoPor: m.validadoPor||null,
          validadoEm:  m.validadoEm||null,
        };
      });

      return {
        mat:matStr, nome:String(row['NOME']||'').trim(),
        funcao, setor:String(row['SETOR']||'').trim(),
        cursosExigidos, cursos,
      };
    });

    res.json({ colaboradores, totais:{
      vencido:totVencido, a_vencer:totAVencer, apto:totApto,
      aceito:totAceito, rejeitado:totRejeitado,
      sem_anexo:totSemAnexo,
    }});
  } catch(e) { console.error(e); res.status(500).json({error:e.message}); }
});

// ── POST upload ───────────────────────────────────────────────
const upload = multer({ dest: path.join(UPLOADS,'_tmp') });

app.post('/api/colaboradores/:mat/curso/:curso', upload.single('pdf'), async (req, res) => {
  try {
    const { mat, curso } = req.params;
    const { data, uploadedBy='João Victor Santos' } = req.body;
    if (!data) return res.status(400).json({ error:'Data obrigatória' });

    // Salva PDF
    let pdfName=null;
    if (req.file) {
      const dest = path.join(UPLOADS, mat);
      fs.mkdirSync(dest,{recursive:true});
      try {
        fs.readdirSync(dest)
          .filter(f=>f.startsWith(curso+'_')&&f.endsWith('.pdf'))
          .forEach(f=>fs.unlinkSync(path.join(dest,f)));
      } catch {}
      pdfName = req.file.originalname||`${curso}_${data}.pdf`;
      fs.renameSync(req.file.path, path.join(dest,`${curso}_${data}.pdf`));
    }

    // Atualiza meta
    const meta = lerMeta();
    if (!meta[mat]) meta[mat]={};
    meta[mat][curso] = {
      ...(meta[mat][curso]||{}),
      uploadedBy, uploadedAt: new Date().toISOString().slice(0,10),
      ...(pdfName ? {pdfName} : {}),
    };
    // Se subir novo PDF, reseta validação
    if (req.file) {
      meta[mat][curso].validado   = false;
      meta[mat][curso].apto       = null;
      meta[mat][curso].obs        = '';
      meta[mat][curso].validadoPor= null;
      meta[mat][curso].validadoEm = null;
    }
    salvarMeta(meta);

    // Atualiza Excel
    const wb    = xlsx.readFile(XLSX_PATH,{cellDates:true});
    const ws    = wb.Sheets[wb.SheetNames[0]];
    const range = xlsx.utils.decode_range(ws['!ref']);
    const targetHeader = COL_REVERSE[curso];
    let colIdx=-1;
    for (let c=range.s.c;c<=range.e.c;c++) {
      const cell = ws[xlsx.utils.encode_cell({r:0,c})];
      if (cell && String(cell.v).trim()===targetHeader) { colIdx=c; break; }
    }
    if (colIdx===-1) return res.status(404).json({error:`Coluna '${curso}' não encontrada`});
    let rowIdx=-1;
    for (let r=1;r<=range.e.r;r++) {
      const cell = ws[xlsx.utils.encode_cell({r,c:0})];
      if (cell && String(cell.v).trim()===mat) { rowIdx=r; break; }
    }
    if (rowIdx===-1) return res.status(404).json({error:`MAT ${mat} não encontrado`});
    const [y,m,d2]=data.split('-');
    ws[xlsx.utils.encode_cell({r:rowIdx,c:colIdx})]={t:'s',v:`${d2}/${m}/${y}`};
    await writeXlsxSafe(wb, XLSX_PATH);

    res.json({ok:true});
  } catch(e) {
    console.error(e);
    res.status(500).json({error: e.code==='EBUSY'||e.code==='EACCES'
      ? 'Arquivo Excel em uso. Feche o Excel e tente novamente.' : e.message});
  }
});

// ── POST validação ────────────────────────────────────────────
app.post('/api/colaboradores/:mat/curso/:curso/validar', (req, res) => {
  try {
    const { mat, curso } = req.params;
    const { apto, obs='', validadoPor='João Victor Santos' } = req.body;
    if (apto===undefined) return res.status(400).json({error:'Campo apto obrigatório'});
    const meta = lerMeta();
    if (!meta[mat]) meta[mat]={};
    if (!meta[mat][curso]) meta[mat][curso]={};
    Object.assign(meta[mat][curso], {
      validado:true, apto:!!apto, obs,
      validadoPor, validadoEm: new Date().toISOString().slice(0,10),
    });
    salvarMeta(meta);
    res.json({ok:true});
  } catch(e) { console.error(e); res.status(500).json({error:e.message}); }
});

app.listen(PORT, ()=>console.log(`Cuidar server → http://localhost:${PORT}`));
