/*
 * Lógica de cruzamento: Takeoff (linhas 100%) × SGJ × Lista de Pastas × Rede de Precedência.
 * Funciona no navegador (window.Core) e no Node (require) para testes.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Core = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------- utilidades ----------
  function colIdx(letra) {
    var n = 0, s = String(letra || '').toUpperCase().trim();
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1;
  }

  // Normaliza código de linha: '' -> ", 1.1/2 == 1 1/2, sem espaços, maiúsculas
  function norm(s) {
    if (s === null || s === undefined) return '';
    s = String(s).replace(/''/g, '"').replace(/[“”″]/g, '"');
    s = s.replace(/(\d)[ .]+(\d+\/\d+)/g, '$1 $2');
    return s.replace(/\s+/g, '').toUpperCase();
  }

  // Normaliza texto livre (cabeçalhos, nomes de aba, status)
  function normTxt(s) {
    return String(s === null || s === undefined ? '' : s)
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/\s+/g, ' ').trim().toUpperCase();
  }

  function cell(row, idx) {
    if (!row) return null;
    var v = row[idx];
    if (v === undefined || v === null) return null;
    if (typeof v === 'string' && v.trim() === '') return null;
    return v;
  }

  function num(v) {
    if (typeof v === 'number') return v;
    if (v === null || v === undefined) return 0;
    var n = parseFloat(String(v).replace(',', '.'));
    return isNaN(n) ? 0 : n;
  }

  // ---------- configuração padrão ----------
  var CONFIG_PADRAO = {
    takeoff: {
      aba: 'PROG TUB',
      linha: 'G', pasta: 'F', spool: 'I', marcador: 'O',
      prevJ: 'AO', realJ: 'AP', avanJ: 'AQ',
      linhaInicial: 6,          // primeira linha de dados no Excel
      avancoMinimo: 1           // 1 = 100%
    },
    sgj: {
      aba: 'SGJ',
      linha: 'C', junta: 'L', status: 'CP',
      statusLiberada: 'Junta Liberada'
    },
    pastas: {
      aba: 'MILPLAN',
      pasta: 'N', lineCode: 'R', subsistema: 'U',
      linhaInicial: 5
    }
  };

  // ---------- leitores ----------
  // rows: array de arrays (linha 1 do Excel = rows[0])

  function lerTakeoff(rows, cfg) {
    cfg = cfg || CONFIG_PADRAO.takeoff;
    var cL = colIdx(cfg.linha), cP = colIdx(cfg.pasta), cS = colIdx(cfg.spool),
        cM = colIdx(cfg.marcador), cPr = colIdx(cfg.prevJ), cRe = colIdx(cfg.realJ);
    var mapa = new Map();
    for (var i = (cfg.linhaInicial || 1) - 1; i < rows.length; i++) {
      var r = rows[i];
      var linha = cell(r, cL);
      if (!linha) continue;
      if (cell(r, cS) !== null) continue;                    // linha de spool / atividade
      var marc = cell(r, cM);
      if (marc === null || String(marc).trim() !== '-') continue; // só a linha-mãe
      var k = norm(linha);
      var e = mapa.get(k);
      if (!e) {
        e = { chave: k, linha: String(linha).trim(), pastaTakeoff: new Set(), prevJ: 0, realJ: 0, linhaExcel: [] };
        mapa.set(k, e);
      }
      var p = cell(r, cP); if (p) e.pastaTakeoff.add(String(p).trim());
      e.prevJ += num(cell(r, cPr));
      e.realJ += num(cell(r, cRe));
      e.linhaExcel.push(i + 1);
    }
    var todas = Array.from(mapa.values()).map(function (e) {
      e.avanco = e.prevJ > 0 ? e.realJ / e.prevJ : 0;
      e.pastaTakeoff = Array.from(e.pastaTakeoff).sort();
      return e;
    });
    var min = cfg.avancoMinimo === undefined ? 1 : cfg.avancoMinimo;
    var selecionadas = todas.filter(function (e) { return e.prevJ > 0 && e.avanco >= min - 1e-9; });
    return { todas: todas, selecionadas: selecionadas };
  }

  function lerSGJ(rows, cfg) {
    cfg = cfg || CONFIG_PADRAO.sgj;
    var cL = colIdx(cfg.linha), cJ = colIdx(cfg.junta), cS = colIdx(cfg.status);
    var lib = normTxt(cfg.statusLiberada);
    var mapa = new Map();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var linha = cell(r, cL);
      if (!linha) continue;
      var k = norm(linha);
      var e = mapa.get(k);
      if (!e) { e = { previstas: 0, liberadas: 0, pendentes: [] }; mapa.set(k, e); }
      e.previstas++;
      var st = cell(r, cS);
      if (normTxt(st) === lib) e.liberadas++;
      else e.pendentes.push({ junta: String(cell(r, cJ) || '(sem nº)'), status: st === null ? '(sem status)' : String(st).trim() });
    }
    return mapa;
  }

  // Procura o cabeçalho da Lista de Pastas para descobrir as colunas sozinho
  function detectarColunasPastas(rows, cfg) {
    var out = Object.assign({}, cfg);
    for (var i = 0; i < Math.min(rows.length, 20); i++) {
      var r = rows[i] || [];
      var achou = {};
      for (var c = 0; c < r.length; c++) {
        var t = normTxt(r[c]);
        if (!t) continue;
        if (t.indexOf('PASTA - PARA') >= 0 || t.indexOf('PASTA PARA') >= 0) achou.pasta = c;
        else if (t === 'LINE CODE') achou.lineCode = c;
        else if (t.indexOf('SUBSYSTEM') === 0) achou.subsistema = c;
      }
      if (achou.pasta !== undefined && achou.lineCode !== undefined) {
        out._pasta = achou.pasta; out._lineCode = achou.lineCode;
        if (achou.subsistema !== undefined) out._subsistema = achou.subsistema;
        out.linhaInicial = i + 2;
        out._detectado = true;
        break;
      }
    }
    return out;
  }

  function lerPastas(rows, cfg) {
    cfg = detectarColunasPastas(rows, cfg || CONFIG_PADRAO.pastas);
    var cP = cfg._pasta !== undefined ? cfg._pasta : colIdx(cfg.pasta);
    var cL = cfg._lineCode !== undefined ? cfg._lineCode : colIdx(cfg.lineCode);
    var cS = cfg._subsistema !== undefined ? cfg._subsistema : colIdx(cfg.subsistema);
    var mapa = new Map();
    for (var i = (cfg.linhaInicial || 1) - 1; i < rows.length; i++) {
      var r = rows[i];
      var linha = cell(r, cL);
      if (!linha) continue;
      var k = norm(linha);
      var e = mapa.get(k);
      if (!e) { e = { pastas: new Set(), subsistemas: new Set() }; mapa.set(k, e); }
      var p = cell(r, cP);
      if (p && /\d/.test(String(p))) e.pastas.add(String(p).trim());
      var s = cell(r, cS);
      if (s && /^\d{3}-/.test(String(s).trim())) e.subsistemas.add(String(s).trim());
    }
    return { mapa: mapa, detectado: !!cfg._detectado };
  }

  // ---------- processamento ----------
  var SIT = {
    OK: 'OK',
    EXECUTAR: 'PODE EXECUTAR',
    SEM_SGJ: 'SEM JUNTAS NO SGJ',
    VERIFICAR: 'VERIFICAR'
  };

  function processar(entrada) {
    var linhas = entrada.takeoff.selecionadas;
    var sgj = entrada.sgj;
    var pastas = entrada.pastas ? entrada.pastas.mapa : new Map();
    var rede = entrada.rede || [];
    var redeMap = new Map();
    rede.forEach(function (n) { if (n && n.subsistema) redeMap.set(String(n.subsistema).trim().toUpperCase(), n); });

    var res = linhas.map(function (t) {
      var s = sgj.get(t.chave) || { previstas: 0, liberadas: 0, pendentes: [] };
      var pend = s.previstas - s.liberadas;
      var p = pastas.get(t.chave);
      var listaPastas = p ? Array.from(p.pastas).sort() : [];
      var subs = p ? Array.from(p.subsistemas).sort() : [];
      var nos = subs.map(function (x) { return redeMap.get(x.toUpperCase()); }).filter(Boolean);
      nos.sort(function (a, b) { return num(a.etapa) - num(b.etapa); });
      var no = nos[0] || null;

      var status;
      if (s.previstas === 0) status = 'NÃO ENCONTRADA';
      else if (pend === 0) status = '100% CONFIRMADO';
      else status = 'DIVERGÊNCIA - ' + pend + ' JUNTAS NÃO LIBERADAS';

      var pastaConfere = null;
      if (t.pastaTakeoff.length && listaPastas.length) {
        pastaConfere = t.pastaTakeoff.some(function (x) { return listaPastas.indexOf(x) >= 0; });
      }

      return {
        linha: t.linha,
        chave: t.chave,
        avanco: t.avanco,
        juntasTakeoff: t.prevJ,
        previstas: s.previstas,
        liberadas: s.liberadas,
        pendentes: pend,
        juntasPendentes: s.pendentes,
        status: status,
        divergente: pend > 0,
        pastas: listaPastas,
        pastaTakeoff: t.pastaTakeoff,
        pastaConfere: pastaConfere,
        subsistemas: subs,
        etapa: no ? num(no.etapa) : null,
        etapaNome: no ? (no.etapaNome || ('Etapa ' + no.etapa)) : (subs.length ? 'Sem posição na rede' : 'Sem subsistema'),
        grupo: no ? no.grupo : '',
        descricao: no ? no.descricao : '',
        predecessoras: no ? no.predecessoras : '',
        sucessoras: no ? no.sucessoras : '',
        situacao: null
      };
    });

    // etapas que ainda têm pendência
    var etapasPend = new Set();
    res.forEach(function (r) { if (r.divergente && r.etapa !== null) etapasPend.add(r.etapa); });
    var listaEtapasPend = Array.from(etapasPend).sort(function (a, b) { return a - b; });

    function aguardando(etapa) {
      var ant = listaEtapasPend.filter(function (e) { return e < etapa; });
      return ant.length ? ant[ant.length - 1] : null;
    }

    res.forEach(function (r) {
      if (r.previstas === 0) r.situacao = SIT.SEM_SGJ;
      else if (!r.divergente) r.situacao = SIT.OK;
      else if (r.etapa === null) r.situacao = SIT.VERIFICAR;
      else {
        var a = aguardando(r.etapa);
        r.situacao = a === null ? SIT.EXECUTAR : 'AGUARDANDO PRIORIDADE ' + a;
      }
    });

    // ordenação: etapa, divergentes primeiro, linha
    res.sort(function (a, b) {
      var ea = a.etapa === null ? 999 : a.etapa, eb = b.etapa === null ? 999 : b.etapa;
      if (ea !== eb) return ea - eb;
      if (a.etapaNome !== b.etapaNome) return a.etapaNome < b.etapaNome ? -1 : 1;
      if (a.divergente !== b.divergente) return a.divergente ? -1 : 1;
      return a.linha < b.linha ? -1 : a.linha > b.linha ? 1 : 0;
    });

    // resumo por etapa
    var resumoMap = new Map();
    res.forEach(function (r) {
      var key = r.etapa === null ? 'z' + r.etapaNome : 'e' + String(r.etapa).padStart(4, '0');
      var e = resumoMap.get(key);
      if (!e) {
        var no = r.etapa !== null ? rede.filter(function (n) { return num(n.etapa) === r.etapa; }) : [];
        e = {
          etapa: r.etapa, etapaNome: r.etapaNome, grupo: r.grupo,
          libera: uniq(no.map(function (n) { return n.libera || n.sucessoras; })).join('; '),
          linhas: 0, linhasPendentes: 0, juntasPendentes: 0, juntasPrevistas: 0
        };
        resumoMap.set(key, e);
      }
      e.linhas++;
      e.juntasPrevistas += r.previstas;
      if (r.divergente) { e.linhasPendentes++; e.juntasPendentes += r.pendentes; }
    });
    var resumo = Array.from(resumoMap.entries()).sort(function (a, b) { return a[0] < b[0] ? -1 : 1; }).map(function (x) {
      var e = x[1];
      if (e.etapa === null) e.situacao = SIT.VERIFICAR;
      else if (e.linhasPendentes === 0) e.situacao = 'SEM PENDÊNCIAS';
      else { var a = aguardando(e.etapa); e.situacao = a === null ? SIT.EXECUTAR : 'AGUARDANDO PRIORIDADE ' + a; }
      return e;
    });

    // indicadores
    var kpi = {
      linhas: res.length,
      linhasDivergentes: res.filter(function (r) { return r.divergente; }).length,
      linhasSemSGJ: res.filter(function (r) { return r.previstas === 0; }).length,
      linhasSemPasta: res.filter(function (r) { return !r.pastas.length; }).length,
      juntasPrevistas: sum(res, 'previstas'),
      juntasLiberadas: sum(res, 'liberadas'),
      juntasPendentes: sum(res, 'pendentes'),
      podeExecutar: res.filter(function (r) { return r.situacao === SIT.EXECUTAR; }).length,
      pastaDivergente: res.filter(function (r) { return r.pastaConfere === false; }).length,
      linhasTakeoff: entrada.takeoff.todas.length
    };

    // status das juntas pendentes
    var statusJuntas = {};
    res.forEach(function (r) { r.juntasPendentes.forEach(function (j) { statusJuntas[j.status] = (statusJuntas[j.status] || 0) + 1; }); });

    // pastas com mais juntas pendentes
    var porPasta = {};
    res.forEach(function (r) {
      if (!r.divergente) return;
      var p = r.pastas.length ? r.pastas[0].replace(/(\.\d+)+$/, '') : '(sem pasta)';
      porPasta[p] = (porPasta[p] || 0) + r.pendentes;
    });

    return { linhas: res, resumo: resumo, kpi: kpi, statusJuntas: statusJuntas, porPasta: porPasta };
  }

  function uniq(a) { var s = []; a.forEach(function (x) { if (x && s.indexOf(x) < 0) s.push(x); }); return s; }
  function sum(a, k) { return a.reduce(function (t, r) { return t + (r[k] || 0); }, 0); }

  return {
    colIdx: colIdx, norm: norm, normTxt: normTxt,
    CONFIG_PADRAO: CONFIG_PADRAO, SIT: SIT,
    lerTakeoff: lerTakeoff, lerSGJ: lerSGJ, lerPastas: lerPastas,
    processar: processar
  };
});
