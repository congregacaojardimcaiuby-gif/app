/**
 * Programação da congregação — ponte entre a planilha e o app.
 * v2 — leitura via JSONP (suporte a callback=) para evitar CORS no navegador.
 *
 * Como atualizar:
 * 1. No Apps Script, substitua todo o conteúdo pelo texto abaixo.
 * 2. Implantar > Gerenciar implantações > editar (lápis) > Nova versão > Salvar.
 *    O endereço /exec NÃO muda — só o código que ele serve.
 */

const ABAS = {
  config: 'Config', pessoas: 'Pessoas', semanas: 'Semanas', partes: 'Partes',
  fimDeSemana: 'FimDeSemana', multimidia: 'Multimidia', indicadores: 'Indicadores',
  limpezaSemanal: 'LimpezaSemanal', limpezaPeriodica: 'LimpezaPeriodica', relatorios: 'Relatorios'
};

function planilha_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function aba_(nome){
  const ws = planilha_().getSheetByName(nome);
  if (!ws) throw new Error('Aba não encontrada: ' + nome);
  return ws;
}
function paraBool_(v){ return String(v).trim().toUpperCase() === 'SIM'; }
function paraSimNao_(v){ return v ? 'SIM' : 'NAO'; }
function paraTexto_(v){ return v === null || v === undefined ? '' : String(v); }

function lerAba_(nome){
  const ws = aba_(nome);
  const vals = ws.getDataRange().getValues();
  if (vals.length < 2) return [];
  const cab = vals[0].map(h => String(h).trim());
  return vals.slice(1)
    .filter(linha => linha.some(c => c !== '' && c !== null))
    .map(linha => Object.fromEntries(cab.map((h,i) => [h, linha[i]])));
}

function escreverAba_(nome, cabecalho, linhas){
  const ws = aba_(nome);
  const ultimaLinha = ws.getMaxRows();
  if (ultimaLinha > 1) ws.getRange(2, 1, ultimaLinha - 1, ws.getMaxColumns()).clearContent();
  if (linhas.length){
    ws.getRange(2, 1, linhas.length, cabecalho.length).setValues(linhas);
  }
}

function montarDB_(){
  const cfgLinhas = lerAba_(ABAS.config);
  const cfg = {}; const pins = {};
  cfgLinhas.forEach(l => {
    const chave = paraTexto_(l.Chave);
    if (chave.indexOf('pin_') === 0) pins[chave.slice(4)] = paraTexto_(l.Valor);
    else cfg[chave] = paraTexto_(l.Valor);
  });
  cfg.pins = pins;

  const pessoas = lerAba_(ABAS.pessoas).map(l => ({
    id: Number(l.ID), nome: paraTexto_(l.Nome), grupo: paraTexto_(l.Grupo),
    podePresidente: paraBool_(l.PodePresidente), podeDirigenteEstudo: paraBool_(l.PodeDirigenteEstudo),
    podeVolante: paraBool_(l.PodeVolante), podeIndicador: paraBool_(l.PodeIndicador),
    podeMesa: paraBool_(l.PodeMesa), podeLeitorSentinela: paraBool_(l.PodeLeitorSentinela),
    podeLeitorEstudo: paraBool_(l.PodeLeitorEstudo)
  }));

  const partesLinhas = lerAba_(ABAS.partes);
  const semanas = lerAba_(ABAS.semanas).map(l => {
    const id = paraTexto_(l.IdSemana);
    const partes = partesLinhas.filter(p => paraTexto_(p.IdSemana) === id)
      .sort((a,b) => Number(a.Ordem) - Number(b.Ordem))
      .map(p => ({ secao:paraTexto_(p.Secao), titulo:paraTexto_(p.Titulo), min:Number(p.Minutos)||0,
                   tipo:paraTexto_(p.Tipo), nome:paraTexto_(p.Designado), ajudante:paraTexto_(p.Ajudante) }));
    return { id, inicio:paraTexto_(l.Inicio), fim:paraTexto_(l.Fim), textoBiblico:paraTexto_(l.TextoBiblico),
      presidente:paraTexto_(l.Presidente), oracaoInicial:paraTexto_(l.OracaoInicial), oracaoFinal:paraTexto_(l.OracaoFinal),
      canticoInicial:paraTexto_(l.CanticoInicial), canticoMeio:paraTexto_(l.CanticoTransicao), canticoFinal:paraTexto_(l.CanticoFinal),
      partes };
  });

  const fimDeSemana = lerAba_(ABAS.fimDeSemana).map(l => ({ data:paraTexto_(l.Data), presidente:paraTexto_(l.Presidente), leitor:paraTexto_(l.Leitor) }));
  const multimidia = lerAba_(ABAS.multimidia).map(l => ({ data:paraTexto_(l.Data), mesa:paraTexto_(l.Mesa), volante1:paraTexto_(l.Volante1), volante2:paraTexto_(l.Volante2) }));
  const indicadores = lerAba_(ABAS.indicadores).map(l => ({ data:paraTexto_(l.Data), indicador1:paraTexto_(l.Indicador1), indicador2:paraTexto_(l.Indicador2) }));
  const limpezaSemanal = lerAba_(ABAS.limpezaSemanal).map(l => ({ data:paraTexto_(l.Data), grupo:paraTexto_(l.Grupo) }));
  const limpezaPeriodica = lerAba_(ABAS.limpezaPeriodica).map(l => ({ data:paraTexto_(l.Data), observacoes:paraTexto_(l.Observacoes) }));
  const relatorios = lerAba_(ABAS.relatorios).map(l => ({ id:paraTexto_(l.Id), mesAno:paraTexto_(l.MesAno), nome:paraTexto_(l.Nome),
    tipo:paraTexto_(l.Tipo), horaCampo: l.HoraCampo === '' ? '' : Number(l.HoraCampo), estudos: l.Estudos === '' ? '' : Number(l.Estudos), status:paraTexto_(l.Status) }));

  return { config: cfg, pessoas, semanasMS: semanas, fimDeSemana, multimidia, indicadores, limpezaSemanal, limpezaPeriodica, relatorios };
}

function checarPin_(pin, cfg){
  pin = paraTexto_(pin).trim();
  if (!pin) return '';
  if (pin === cfg.pinMestre) return 'todos';
  for (const k in cfg.pins) if (cfg.pins[k] && pin === cfg.pins[k]) return k;
  return '';
}

function salvarPorEscopo_(escopo, dados){
  if (escopo === 'todos' || escopo === 'ms'){
    const semLin = dados.semanasMS.map(s => [s.id, s.inicio, s.fim, s.textoBiblico, s.presidente, s.oracaoInicial, s.oracaoFinal, s.canticoInicial, s.canticoMeio, s.canticoFinal]);
    escreverAba_(ABAS.semanas, ['IdSemana','Inicio','Fim','TextoBiblico','Presidente','OracaoInicial','OracaoFinal','CanticoInicial','CanticoTransicao','CanticoFinal'], semLin);
    const parLin = [];
    dados.semanasMS.forEach(s => s.partes.forEach((p,i) => parLin.push([s.id, p.secao, i+1, p.titulo, p.min, p.tipo, p.nome, p.ajudante])));
    escreverAba_(ABAS.partes, ['IdSemana','Secao','Ordem','Titulo','Minutos','Tipo','Designado','Ajudante'], parLin);
  }
  if (escopo === 'todos' || escopo === 'fs'){
    escreverAba_(ABAS.fimDeSemana, ['Data','Presidente','Leitor'], dados.fimDeSemana.map(x => [x.data, x.presidente, x.leitor]));
  }
  if (escopo === 'todos' || escopo === 'av'){
    escreverAba_(ABAS.multimidia, ['Data','Mesa','Volante1','Volante2'], dados.multimidia.map(x => [x.data, x.mesa, x.volante1, x.volante2]));
  }
  if (escopo === 'todos' || escopo === 'in'){
    escreverAba_(ABAS.indicadores, ['Data','Indicador1','Indicador2'], dados.indicadores.map(x => [x.data, x.indicador1, x.indicador2]));
  }
  if (escopo === 'todos' || escopo === 'lp'){
    escreverAba_(ABAS.limpezaSemanal, ['Data','Grupo'], dados.limpezaSemanal.map(x => [x.data, x.grupo]));
    escreverAba_(ABAS.limpezaPeriodica, ['Data','Observacoes'], dados.limpezaPeriodica.map(x => [x.data, x.observacoes]));
  }
  if (escopo === 'todos' || escopo === 'rm'){
    escreverAba_(ABAS.relatorios, ['Id','MesAno','Nome','Tipo','HoraCampo','Estudos','Status'],
      dados.relatorios.map(x => [x.id, x.mesAno, x.nome, x.tipo, x.horaCampo, x.estudos, x.status]));
  }
  if (escopo === 'todos'){
    const pesLin = dados.pessoas.map(p => [p.id, p.nome, p.grupo, paraSimNao_(p.podePresidente), paraSimNao_(p.podeDirigenteEstudo),
      paraSimNao_(p.podeVolante), paraSimNao_(p.podeIndicador), paraSimNao_(p.podeMesa), paraSimNao_(p.podeLeitorSentinela), paraSimNao_(p.podeLeitorEstudo)]);
    escreverAba_(ABAS.pessoas, ['ID','Nome','Grupo','PodePresidente','PodeDirigenteEstudo','PodeVolante','PodeIndicador','PodeMesa','PodeLeitorSentinela','PodeLeitorEstudo'], pesLin);
    const cfgLin = [['congregacao', dados.config.congregacao], ['pinMestre', dados.config.pinMestre]];
    for (const k in dados.config.pins) cfgLin.push(['pin_'+k, dados.config.pins[k]]);
    escreverAba_(ABAS.config, ['Chave','Valor'], cfgLin);
  }
}

function salvarRelatorio_(registro){
  const hoje = new Date();
  const mesAtual = Utilities.formatDate(hoje, Session.getScriptTimeZone() || 'America/Sao_Paulo', 'yyyy-MM');
  if (registro.mesAno !== mesAtual) throw new Error('Só é possível enviar o relatório do mês atual.');
  const id = registro.mesAno + '|' + String(registro.nome).trim().toLowerCase();
  const ws = aba_(ABAS.relatorios);
  const vals = ws.getDataRange().getValues();
  let linhaAlvo = -1;
  for (let i = 1; i < vals.length; i++) if (vals[i][0] === id) { linhaAlvo = i + 1; break; }
  const linha = [id, registro.mesAno, registro.nome, registro.tipo, registro.horaCampo, registro.estudos, 'Entregue'];
  if (linhaAlvo > 0) ws.getRange(linhaAlvo, 1, 1, linha.length).setValues([linha]);
  else ws.appendRow(linha);
  return id;
}

/* ============================================================ endpoints HTTP */

/**
 * GET — leitura pública.
 * Suporta ?callback=nome para JSONP (necessário quando o app está em
 * outro domínio e o navegador bloqueia fetch() por CORS).
 * Exemplo: /exec?callback=_cb123
 * O servidor responde: _cb123({"ok":true,"dados":{...}})
 */
function doGet(e){
  const dados = montarDB_();
  const payload = JSON.stringify({ ok:true, dados });
  const cb = e && e.parameter && e.parameter.callback;
  if (cb){
    // JSONP: devolve JavaScript puro, não JSON
    return ContentService.createTextOutput(cb + '(' + payload + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  try {
    const corpo = JSON.parse(e.postData.contents);
    const cfg = montarDB_().config;

    if (corpo.acao === 'relatorio'){
      const id = salvarRelatorio_(corpo.registro);
      return saida_({ ok:true, id });
    }

    if (corpo.acao === 'salvar'){
      const escopo = checarPin_(corpo.pin, cfg);
      if (!escopo) return saida_({ ok:false, erro:'PIN inválido.' });
      salvarPorEscopo_(escopo, corpo.dados);
      return saida_({ ok:true, dados: montarDB_() });
    }

    return saida_({ ok:false, erro:'Ação desconhecida.' });
  } catch (err){
    return saida_({ ok:false, erro: String(err) });
  }
}

function saida_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
