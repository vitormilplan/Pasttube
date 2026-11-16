/* Interface do painel. Depende de: XLSX (SheetJS), Chart (Chart.js), Core, REDE_PADRAO */
(function () {
  'use strict';

  var LS_REDE = 'painelJuntas.rede.v1';
  var LS_CFG = 'painelJuntas.config.v1';

  var estado = {
    arquivos: { takeoff: null, sgj: null, pastas: null },
    dados: { takeoff: null, sgj: null, pastas: null },
    resultado: null,
    graficos: {},
    ordem: { k: null, asc: true },
    abertos: new Set()
  };

  // ---------- armazenamento ----------
  function lsGet(k, padrao) {
    try { var v = localStorage.getItem(k); return v ? JSON.parse(v) : padrao; } catch (e) { return padrao; }
  }
  function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  function getRede() { return lsGet(LS_REDE, null) || clone(window.REDE_PADRAO); }
  function getCfg() {
    var salvo = lsGet(LS_CFG, null), pad = clone(Core.CONFIG_PADRAO);
    if (!salvo) return pad;
    Object.keys(pad).forEach(function (k) { pad[k] = Object.assign(pad[k], salvo[k] || {}); });
    return pad;
  }

  // ---------- helpers de DOM ----------
  function $(s, el) { return (el || document).querySelector(s); }
  function $$(s, el) { return Array.prototype.slice.call((el || document).querySelectorAll(s)); }
  function esc(s) { return String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function fmt(n) { return Number(n || 0).toLocaleString('pt-BR'); }
  function pct(n) { return (n * 100).toLocaleString('pt-BR', { maximumFractionDigits: 1 }) + '%'; }
  function msg(el, txt, tipo) { el.textContent = txt || ''; el.className = 'msg' + (tipo ? ' ' + tipo : ''); }
  function carregando(on, txt) { $('#carregando').hidden = !on; if (txt) $('#carregando-txt').textContent = txt; }
  function pausa() { return new Promise(function (r) { setTimeout(r, 30); }); }

  function irPara(aba) {
    $$('#abas button').forEach(function (b) { b.classList.toggle('ativa', b.dataset.aba === aba); });
    $$('.aba').forEach(function (s) { s.classList.toggle('ativa', s.id === 'aba-' + aba); });
    window.scrollTo(0, 0);
  }

  // ---------- leitura de planilhas ----------
  function acharAba(nomes, alvo, contem) {
    var a = Core.normTxt(alvo).replace(/ /g, '');
    var n = nomes.find(function (x) { return Core.normTxt(x).replace(/ /g, '') === a; });
    if (!n && contem) n = nomes.find(function (x) { return Core.normTxt(x).replace(/ /g, '').indexOf(contem) >= 0; });
    return n;
  }

  function lerPlanilha(buf, alvo, contem, usarPrimeira) {
    var nomes = XLSX.read(buf, { type: 'array', bookSheets: true }).SheetNames;
    var nome = acharAba(nomes, alvo, contem) || (usarPrimeira ? nomes[0] : null);
    if (!nome) throw new Error('Aba "' + alvo + '" não encontrada. Abas do arquivo: ' + nomes.join(', '));
    var wb = XLSX.read(buf, { type: 'array', sheets: [nome], cellFormula: false, cellHTML: false, cellText: false, cellStyles: false });
    var ws = wb.Sheets[nome];
    if (!ws || !ws['!ref']) throw new Error('A aba "' + nome + '" está vazia.');
    var ref = XLSX.utils.decode_range(ws['!ref']); ref.s.r = 0; ref.s.c = 0;
    return { nome: nome, rows: XLSX.utils.sheet_to_json(ws, { header: 1, range: ref, defval: null, raw: true, blankrows: true }) };
  }

  async function carregarArquivo(tipo, file) {
    var drop = $('.drop[data-tipo="' + tipo + '"]');
    var status = $('.drop-status', drop);
    drop.classList.remove('carregado', 'erro');
    status.textContent = 'Lendo ' + file.name + '…';
    carregando(true, 'Lendo ' + file.name + (file.size > 5e6 ? ' (arquivo grande, pode levar alguns segundos)…' : '…'));
    await pausa();
    try {
      var cfg = getCfg();
      var buf = new Uint8Array(await file.arrayBuffer());
      var p, info;
      if (tipo === 'takeoff') {
        p = lerPlanilha(buf, cfg.takeoff.aba, 'PROGTUB');
        estado.dados.takeoff = Core.lerTakeoff(p.rows, cfg.takeoff);
        info = estado.dados.takeoff.selecionadas.length + ' linhas em 100% (de ' + estado.dados.takeoff.todas.length + ')';
      } else if (tipo === 'sgj') {
        p = lerPlanilha(buf, cfg.sgj.aba, 'SGJ', true);
        estado.dados.sgj = Core.lerSGJ(p.rows, cfg.sgj);
        var tot = 0; estado.dados.sgj.forEach(function (e) { tot += e.previstas; });
        info = fmt(tot) + ' juntas em ' + fmt(estado.dados.sgj.size) + ' linhas';
      } else {
        p = lerPlanilha(buf, cfg.pastas.aba, 'MILPLAN', true);
        estado.dados.pastas = Core.lerPastas(p.rows, cfg.pastas);
        info = fmt(estado.dados.pastas.mapa.size) + ' linhas' + (estado.dados.pastas.detectado ? ' (colunas detectadas pelo cabeçalho)' : '');
      }
      estado.arquivos[tipo] = file.name;
      drop.classList.add('carregado');
      status.textContent = '✓ ' + file.name + ' — aba "' + p.nome + '": ' + info;
    } catch (e) {
      console.error(e);
      estado.dados[tipo] = null;
      drop.classList.add('erro');
      status.textContent = '✗ ' + (e.message || e);
    }
    carregando(false);
    $('#btn-processar').disabled = !(estado.dados.takeoff && estado.dados.sgj);
    msg($('#msg'), estado.dados.pastas ? '' : (estado.dados.takeoff && estado.dados.sgj ? 'Sem a Lista de Pastas o painel sai sem pasta, subsistema e prioridade.' : ''));
  }

  function configurarDrops() {
    $$('.drop').forEach(function (drop) {
      var tipo = drop.dataset.tipo, input = $('input', drop);
      input.addEventListener('change', function () { if (input.files[0]) carregarArquivo(tipo, input.files[0]); input.value = ''; });
      ['dragenter', 'dragover'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('sobre'); }); });
      ['dragleave', 'drop'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('sobre'); }); });
      drop.addEventListener('drop', function (e) { var f = e.dataTransfer.files[0]; if (f) carregarArquivo(tipo, f); });
    });
  }

  // ---------- processamento ----------
  function processar() {
    estado.resultado = Core.processar({
      takeoff: estado.dados.takeoff, sgj: estado.dados.sgj, pastas: estado.dados.pastas, rede: getRede()
    });
    $$('#abas button').forEach(function (b) { b.disabled = false; });
    renderPainel();
    montarFiltros();
    renderTabela();
    irPara('painel');
  }

  function classeSit(s) {
    if (s === Core.SIT.OK || s === 'SEM PENDÊNCIAS') return 'ok';
    if (s === Core.SIT.EXECUTAR) return 'executar';
    if (s === Core.SIT.SEM_SGJ) return 'semsgj';
    if (s === Core.SIT.VERIFICAR) return 'verificar';
    return 'aguardando';
  }

  // ---------- painel ----------
  function renderPainel() {
    var R = estado.resultado, k = R.kpi;
    var lib = k.juntasPrevistas ? k.juntasLiberadas / k.juntasPrevistas : 0;
    $('#kpis').innerHTML = [
      ['Linhas 100% no takeoff', fmt(k.linhas), 'de ' + fmt(k.linhasTakeoff) + ' linhas no takeoff', ''],
      ['Juntas liberadas (SGJ)', pct(lib), fmt(k.juntasLiberadas) + ' de ' + fmt(k.juntasPrevistas), lib >= 0.999 ? 'bom' : ''],
      ['Juntas pendentes', fmt(k.juntasPendentes), 'em ' + fmt(k.linhasDivergentes) + ' linhas', k.juntasPendentes ? 'destaque' : 'bom'],
      ['Pode executar agora', fmt(k.podeExecutar), 'linhas sem bloqueio de etapa anterior', k.podeExecutar ? 'destaque' : ''],
      ['Sem juntas no SGJ', fmt(k.linhasSemSGJ), 'linhas 100% no takeoff não achadas no SGJ', ''],
      ['Sem pasta', fmt(k.linhasSemPasta), 'linhas fora da Lista de Pastas', '']
    ].map(function (x) {
      return '<div class="kpi ' + x[3] + '"><div class="r">' + esc(x[0]) + '</div><div class="v">' + x[1] + '</div><div class="r">' + esc(x[2]) + '</div></div>';
    }).join('');

    $('#escada').innerHTML = R.resumo.map(function (e, i) {
      var cls = e.etapa === null ? 'verificar' : e.situacao === Core.SIT.EXECUTAR ? 'executar' : e.situacao === 'SEM PENDÊNCIAS' ? 'limpo' : 'aguardando';
      return '<div class="degrau ' + cls + '" data-i="' + i + '" style="margin-left:' + (e.etapa === null ? 0 : Math.min((e.etapa - 1) * 24, 160)) + 'px">' +
        '<div class="n">' + (e.etapa === null ? '—' : e.etapa) + '</div>' +
        '<div><div class="t">' + esc(e.etapaNome) + '</div><div class="d">' + esc(e.situacao) + (e.libera ? ' · libera: ' + esc(e.libera) : '') + '</div></div>' +
        '<div class="nums"><b>' + fmt(e.linhasPendentes) + '</b> de ' + fmt(e.linhas) + ' linhas pendentes<br><b>' + fmt(e.juntasPendentes) + '</b> juntas</div></div>';
    }).join('');
    $$('#escada .degrau').forEach(function (d) {
      d.addEventListener('click', function () {
        var e = R.resumo[+d.dataset.i];
        $('#f-etapa').value = e.etapaNome; $('#f-situacao').value = ''; $('#f-pend').checked = true; $('#f-busca').value = '';
        renderTabela(); irPara('linhas');
      });
    });

    // gráficos
    Object.values(estado.graficos).forEach(function (g) { g.destroy(); });
    var cores = { executar: '#b5590f', aguardando: '#9aa3b2', limpo: '#23784a', verificar: '#b3261e' };
    estado.graficos.etapa = new Chart($('#g-etapa'), {
      type: 'bar',
      data: {
        labels: R.resumo.map(function (e) { return (e.etapa === null ? '' : e.etapa + ' · ') + e.etapaNome; }),
        datasets: [{ label: 'Juntas pendentes', data: R.resumo.map(function (e) { return e.juntasPendentes; }),
          backgroundColor: R.resumo.map(function (e) { return e.etapa === null ? cores.verificar : e.situacao === Core.SIT.EXECUTAR ? cores.executar : e.linhasPendentes ? cores.aguardando : cores.limpo; }) }]
      },
      options: { maintainAspectRatio: false, indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { beginAtZero: true, ticks: { precision: 0 } } } }
    });
    var st = Object.entries(R.statusJuntas).sort(function (a, b) { return b[1] - a[1]; });
    estado.graficos.status = new Chart($('#g-status'), {
      type: 'doughnut',
      data: { labels: st.map(function (x) { return x[0]; }), datasets: [{ data: st.map(function (x) { return x[1]; }),
        backgroundColor: ['#1f5fa8', '#b5590f', '#23784a', '#7a4fb3', '#b3261e', '#2a8a8a', '#8a7a2a', '#5c6474'] }] },
      options: { maintainAspectRatio: false, plugins: { legend: { position: 'right' } } }
    });
    var pp = Object.entries(R.porPasta).sort(function (a, b) { return b[1] - a[1]; }).slice(0, 12);
    estado.graficos.pastas = new Chart($('#g-pastas'), {
      type: 'bar',
      data: { labels: pp.map(function (x) { return x[0]; }), datasets: [{ label: 'Juntas pendentes', data: pp.map(function (x) { return x[1]; }), backgroundColor: '#1f5fa8' }] },
      options: { maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, ticks: { precision: 0 } } } }
    });

    // alertas
    var al = [];
    if (k.linhasSemSGJ) al.push(['aviso', k.linhasSemSGJ + ' linhas estão 100% no takeoff mas não têm nenhuma junta no SGJ — o código da linha pode estar escrito diferente ou a linha não foi cadastrada.']);
    if (k.pastaDivergente) {
      var l = R.linhas.filter(function (r) { return r.pastaConfere === false; });
      al.push(['', k.pastaDivergente + ' linhas com pasta diferente entre takeoff e Lista de Pastas: ' +
        l.map(function (r) { return r.linha + ' (takeoff ' + r.pastaTakeoff.join('/') + ' × lista ' + r.pastas.join('/') + ')'; }).join('; ')]);
    }
    if (!estado.dados.pastas) al.push(['aviso', 'Lista de Pastas não carregada: sem pasta, subsistema e prioridade.']);
    var semRede = {};
    R.linhas.forEach(function (r) { if (r.etapa === null && r.subsistemas.length) r.subsistemas.forEach(function (s) { semRede[s] = 1; }); });
    if (Object.keys(semRede).length) al.push(['aviso', 'Subsistemas que não constam na Rede de Precedência: ' + Object.keys(semRede).join(', ') + '. Cadastre na aba "Rede de precedência" para entrarem na ordem.']);
    $('#alertas').innerHTML = al.map(function (a) { return '<div class="alerta ' + a[0] + '">' + esc(a[1]) + '</div>'; }).join('');
  }

  // ---------- tabela de linhas ----------
  function montarFiltros() {
    var R = estado.resultado;
    var sits = Array.from(new Set(R.linhas.map(function (r) { return r.situacao; })));
    var ordemSit = ['PODE EXECUTAR', 'AGUARDANDO', 'VERIFICAR', 'SEM JUNTAS', 'OK'];
    sits.sort(function (a, b) {
      var ia = ordemSit.findIndex(function (x) { return a.indexOf(x) === 0; }), ib = ordemSit.findIndex(function (x) { return b.indexOf(x) === 0; });
      return ia - ib || (a < b ? -1 : 1);
    });
    $('#f-situacao').innerHTML = '<option value="">Todas as situações</option>' + sits.map(function (s) { return '<option>' + esc(s) + '</option>'; }).join('');
    $('#f-etapa').innerHTML = '<option value="">Todas as etapas</option>' + R.resumo.map(function (e) { return '<option value="' + esc(e.etapaNome) + '">' + esc((e.etapa === null ? '' : e.etapa + ' · ') + e.etapaNome) + '</option>'; }).join('');
  }

  function filtradas() {
    var R = estado.resultado; if (!R) return [];
    var b = Core.normTxt($('#f-busca').value), s = $('#f-situacao').value, e = $('#f-etapa').value, p = $('#f-pend').checked;
    var l = R.linhas.filter(function (r) {
      if (s && r.situacao !== s) return false;
      if (e && r.etapaNome !== e) return false;
      if (p && !r.divergente) return false;
      if (b) {
        var txt = Core.normTxt([r.linha, r.pastas.join(' '), r.subsistemas.join(' '), r.status, r.situacao, r.etapaNome,
          r.juntasPendentes.map(function (j) { return j.junta + ' ' + j.status; }).join(' ')].join(' '));
        if (txt.indexOf(b) < 0) return false;
      }
      return true;
    });
    var o = estado.ordem;
    if (o.k) {
      l = l.slice().sort(function (x, y) {
        var a = x[o.k], c = y[o.k];
        if (Array.isArray(a)) { a = a.join(','); c = c.join(','); }
        if (a === null) a = 999; if (c === null) c = 999;
        var r = typeof a === 'number' && typeof c === 'number' ? a - c : String(a).localeCompare(String(c), 'pt-BR', { numeric: true });
        return o.asc ? r : -r;
      });
    }
    return l;
  }

  function renderTabela() {
    var l = filtradas();
    $('#f-cont').textContent = l.length + ' linhas · ' + fmt(l.reduce(function (t, r) { return t + r.pendentes; }, 0)) + ' juntas pendentes';
    var html = l.map(function (r) {
      var aberto = estado.abertos.has(r.chave);
      var pastaHtml = esc(r.pastas.join(', ') || '—') + (r.pastaConfere === false ? ' <span class="warn" title="Takeoff: ' + esc(r.pastaTakeoff.join(', ')) + '">≠ takeoff</span>' : '');
      var linha = '<tr class="l' + (r.divergente ? ' div' : '') + '" data-k="' + esc(r.chave) + '">' +
        '<td class="num">' + (r.etapa === null ? '—' : r.etapa) + '</td><td>' + esc(r.etapaNome) + '</td>' +
        '<td><span class="pill ' + classeSit(r.situacao) + '">' + esc(r.situacao) + '</span></td>' +
        '<td><b>' + esc(r.linha) + '</b></td><td class="num">' + (r.pendentes || '') + '</td><td class="num">' + r.previstas + '</td>' +
        '<td class="num">' + fmt(r.juntasTakeoff) + (r.previstas && r.juntasTakeoff !== r.previstas ? ' <span class="warn" title="Takeoff e SGJ com quantidades diferentes">≠</span>' : '') + '</td>' +
        '<td>' + pastaHtml + '</td><td>' + esc(r.subsistemas.join(', ') || '—') + '</td><td>' + esc(r.status) + '</td></tr>';
      if (aberto) {
        linha += '<tr class="det"><td></td><td colspan="9"><div class="det-grid"><div><b>Juntas pendentes (' + r.juntasPendentes.length + ')</b>' +
          (r.juntasPendentes.length ? '<ul>' + r.juntasPendentes.map(function (j) { return '<li>' + esc(j.junta) + ' — ' + esc(j.status) + '</li>'; }).join('') + '</ul>' : '<p class="nota">Nenhuma.</p>') +
          '</div><div><b>Predecessoras</b><p>' + esc(r.predecessoras || '—') + '</p><b>Grupo</b><p>' + esc(r.grupo || '—') + '</p></div>' +
          '<div><b>Sucessoras</b><p>' + esc(r.sucessoras || '—') + '</p><b>Takeoff</b><p>Pasta ' + esc(r.pastaTakeoff.join(', ') || '—') + ' · ' + fmt(r.juntasTakeoff) + ' juntas · avanço ' + pct(r.avanco) + '</p></div></div></td></tr>';
      }
      return linha;
    }).join('');
    $('#tabela tbody').innerHTML = html || '<tr><td colspan="10" class="nota">Nenhuma linha com esses filtros.</td></tr>';
  }

  function configurarTabela() {
    ['#f-busca', '#f-situacao', '#f-etapa', '#f-pend'].forEach(function (s) { $(s).addEventListener('input', renderTabela); });
    $('#tabela tbody').addEventListener('click', function (e) {
      var tr = e.target.closest('tr.l'); if (!tr) return;
      var k = tr.dataset.k; if (estado.abertos.has(k)) estado.abertos.delete(k); else estado.abertos.add(k);
      renderTabela();
    });
    $$('#tabela th').forEach(function (th) {
      th.addEventListener('click', function () {
        var k = th.dataset.k; if (estado.ordem.k === k) estado.ordem.asc = !estado.ordem.asc; else { estado.ordem.k = k; estado.ordem.asc = true; }
        renderTabela();
      });
    });
  }

  // ---------- exportação ----------
  function exportar() {
    var R = estado.resultado; if (!R) return;
    var wb = XLSX.utils.book_new();
    function aba(nome, linhas, larguras) {
      var ws = XLSX.utils.aoa_to_sheet(linhas);
      ws['!cols'] = larguras.map(function (w) { return { wch: w }; });
      ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: linhas.length - 1, c: linhas[0].length - 1 } }) };
      XLSX.utils.book_append_sheet(wb, ws, nome);
    }
    var res = [['ORDEM', 'ETAPA', 'LINHAS', 'LINHAS PENDENTES', 'JUNTAS PENDENTES', 'SITUAÇÃO', 'LIBERA']];
    R.resumo.forEach(function (e) { res.push([e.etapa === null ? '—' : e.etapa, e.etapaNome, e.linhas, e.linhasPendentes, e.juntasPendentes, e.situacao, e.libera || '—']); });
    res.push(['', 'TOTAL', R.kpi.linhas, R.kpi.linhasDivergentes, R.kpi.juntasPendentes, '', '']);
    aba('RESUMO', res, [8, 40, 10, 18, 18, 28, 50]);

    var det = [['ORDEM', 'ETAPA', 'SITUAÇÃO', 'LINHA', 'STATUS', 'JUNTAS PREVISTAS (SGJ)', 'JUNTAS LIBERADAS', 'JUNTAS PENDENTES', 'JUNTAS TAKEOFF',
      'TAG (JUNTAS PENDENTES)', 'PASTAS', 'PASTA TAKEOFF', 'SUBSISTEMA', 'GRUPO (REDE)', 'DESCRIÇÃO (REDE)', 'PREDECESSORAS', 'SUCESSORAS']];
    R.linhas.forEach(function (r) {
      det.push([r.etapa === null ? '—' : r.etapa, r.etapaNome, r.situacao, r.linha, r.status, r.previstas, r.liberadas, r.pendentes, r.juntasTakeoff,
        r.previstas === 0 ? 'NÃO ENCONTRADA' : r.pendentes === 0 ? 'TODAS LIBERADAS' : r.juntasPendentes.map(function (j) { return j.junta + ' (' + j.status + ')'; }).join(', '),
        r.pastas.join(', ') || 'NÃO ENCONTRADA', r.pastaTakeoff.join(', '), r.subsistemas.join('; ') || 'NÃO ENCONTRADO',
        r.grupo || '—', r.descricao || '—', r.predecessoras || '—', r.sucessoras || '—']);
    });
    aba('PRIORIDADES', det, [8, 34, 26, 32, 38, 12, 12, 12, 12, 70, 30, 18, 24, 34, 30, 50, 50]);

    var jp = [['ORDEM', 'ETAPA', 'SITUAÇÃO', 'LINHA', 'JUNTA', 'STATUS DA JUNTA', 'PASTAS']];
    R.linhas.forEach(function (r) { r.juntasPendentes.forEach(function (j) { jp.push([r.etapa === null ? '—' : r.etapa, r.etapaNome, r.situacao, r.linha, j.junta, j.status, r.pastas.join(', ')]); }); });
    aba('JUNTAS PENDENTES', jp, [8, 34, 26, 32, 40, 36, 30]);

    var d = new Date(), dd = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    XLSX.writeFile(wb, 'Divergencias_Prioridades_' + dd + '.xlsx');
  }

  // ---------- editor da rede ----------
  var CAMPOS_REDE = ['subsistema', 'etapa', 'etapaNome', 'grupo', 'descricao', 'predecessoras', 'sucessoras', 'libera'];
  function renderRede(rede) {
    $('#tabela-rede tbody').innerHTML = rede.map(function (n, i) {
      return '<tr data-i="' + i + '">' + CAMPOS_REDE.map(function (c) {
        return '<td><input data-c="' + c + '" class="' + c + '" ' + (c === 'etapa' ? 'type="number" min="1"' : '') + ' value="' + esc(n[c]) + '"></td>';
      }).join('') + '<td><button data-rm="' + i + '" title="Remover">✕</button></td></tr>';
    }).join('');
  }
  function lerRedeDaTela() {
    return $$('#tabela-rede tbody tr').map(function (tr) {
      var n = {}; $$('input', tr).forEach(function (inp) { n[inp.dataset.c] = inp.dataset.c === 'etapa' ? Number(inp.value) || null : inp.value.trim(); });
      return n;
    }).filter(function (n) { return n.subsistema; });
  }
  function configurarRede() {
    renderRede(getRede());
    $('#rede-add').onclick = function () { var r = lerRedeDaTela(); r.push({ subsistema: '', etapa: '' }); renderRede(r); };
    $('#tabela-rede').addEventListener('click', function (e) {
      var b = e.target.closest('button[data-rm]'); if (!b) return;
      var r = lerRedeDaTela(); r.splice(+b.dataset.rm, 1); renderRede(r);
    });
    $('#rede-salvar').onclick = function () {
      var r = lerRedeDaTela();
      var semEtapa = r.filter(function (n) { return !n.etapa; });
      if (semEtapa.length) { msg($('#rede-msg'), 'Preencha a etapa de: ' + semEtapa.map(function (n) { return n.subsistema; }).join(', '), 'erro'); return; }
      lsSet(LS_REDE, r); renderRede(r);
      msg($('#rede-msg'), 'Rede salva (' + r.length + ' subsistemas).' + (estado.resultado ? ' Painel recalculado.' : ''), 'ok');
      if (estado.resultado) { processar(); irPara('rede'); }
    };
    $('#rede-exportar').onclick = function () {
      var blob = new Blob([JSON.stringify(lerRedeDaTela(), null, 2)], { type: 'application/json' });
      var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'rede-precedencia.json'; a.click();
    };
    $('#rede-importar').onchange = async function (e) {
      try {
        var r = JSON.parse(await e.target.files[0].text());
        if (!Array.isArray(r)) throw new Error('o arquivo precisa ser uma lista');
        renderRede(r); msg($('#rede-msg'), 'Importado — confira e clique em Salvar.', 'ok');
      } catch (err) { msg($('#rede-msg'), 'JSON inválido: ' + err.message, 'erro'); }
      e.target.value = '';
    };
    $('#rede-padrao').onclick = function () {
      if (!confirm('Voltar a rede para o padrão do arquivo js/rede-padrao.js? As edições salvas neste navegador serão perdidas.')) return;
      try { localStorage.removeItem(LS_REDE); } catch (e) { }
      renderRede(getRede()); msg($('#rede-msg'), 'Rede padrão restaurada.', 'ok');
      if (estado.resultado) { processar(); irPara('rede'); }
    };
  }

  // ---------- configurações ----------
  var ROTULOS = {
    takeoff: ['Takeoff', { aba: 'Aba', linha: 'Coluna linha', pasta: 'Coluna pasta', spool: 'Coluna spool (vazia na linha-mãe)', marcador: 'Coluna com "-" na linha-mãe',
      prevJ: 'Coluna juntas previstas', realJ: 'Coluna juntas realizadas', linhaInicial: 'Primeira linha de dados', avancoMinimo: 'Avanço mínimo (1 = 100%)' }],
    sgj: ['SGJ', { aba: 'Aba', linha: 'Coluna linha', junta: 'Coluna junta', status: 'Coluna status', statusLiberada: 'Texto do status liberado' }],
    pastas: ['Lista de Pastas', { aba: 'Aba', pasta: 'Coluna N° da pasta - para', lineCode: 'Coluna Line Code', subsistema: 'Coluna Subsystem', linhaInicial: 'Primeira linha de dados' }]
  };
  function renderConfig() {
    var cfg = getCfg();
    $('#config').innerHTML = Object.keys(ROTULOS).map(function (g) {
      return '<fieldset><legend>' + ROTULOS[g][0] + '</legend>' + Object.keys(ROTULOS[g][1]).map(function (k) {
        return '<label>' + ROTULOS[g][1][k] + '<input data-g="' + g + '" data-k="' + k + '" value="' + esc(cfg[g][k]) + '"></label>';
      }).join('') + '</fieldset>';
    }).join('') + '<p class="nota" style="grid-column:1/-1">Na Lista de Pastas as colunas são achadas sozinhas pelo cabeçalho ("Line Code", "N° DA PASTA - PARA", "Subsystem"); as letras acima só valem se o cabeçalho não for encontrado.</p>';
  }
  function configurarConfig() {
    renderConfig();
    $('#cfg-salvar').onclick = function () {
      var cfg = getCfg();
      $$('#config input').forEach(function (i) {
        var v = i.value.trim();
        cfg[i.dataset.g][i.dataset.k] = (i.dataset.k === 'linhaInicial' || i.dataset.k === 'avancoMinimo') ? Number(v.replace(',', '.')) : v;
      });
      lsSet(LS_CFG, cfg);
      msg($('#cfg-msg'), 'Salvo. Recarregue os arquivos para aplicar.', 'ok');
    };
    $('#cfg-padrao').onclick = function () { try { localStorage.removeItem(LS_CFG); } catch (e) { } renderConfig(); msg($('#cfg-msg'), 'Configuração padrão restaurada.', 'ok'); };
  }

  // ---------- início ----------
  function iniciar() {
    if (typeof XLSX === 'undefined' || typeof Chart === 'undefined') {
      document.body.insertAdjacentHTML('afterbegin', '<div class="alerta" style="margin:12px">Não foi possível carregar as bibliotecas (SheetJS/Chart.js). Verifique a conexão com a internet.</div>');
    }
    $$('#abas button').forEach(function (b) { b.addEventListener('click', function () { if (!b.disabled) irPara(b.dataset.aba); }); });
    configurarDrops();
    configurarTabela();
    configurarRede();
    configurarConfig();
    $('#btn-processar').addEventListener('click', processar);
    $('#btn-exportar').addEventListener('click', exportar);
    $('#btn-exportar2').addEventListener('click', exportar);
  }
  document.addEventListener('DOMContentLoaded', iniciar);
})();
