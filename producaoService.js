const ABA_PRODUCAO = 'PRODUCAO';
const ABA_PRODUCAO_ETAPAS = 'PRODUCAO_ETAPAS';
const ABA_PRODUCAO_CONSUMO = 'PRODUCAO_CONSUMO';
const ABA_PRODUCAO_MATERIAIS = 'PRODUCAO_MATERIAIS';
const ABA_PRODUCAO_MATERIAIS_PREVISTOS = 'PRODUCAO_MATERIAIS_PREVISTOS';
const ABA_PRODUCAO_VINCULOS = 'PRODUCAO_VINCULOS_MATERIAIS';
const ABA_PRODUCAO_NECESSIDADES_ENTRADA = 'PRODUCAO_NECESSIDADES_ENTRADA';
const ABA_PRODUCAO_RESERVAS_ENTRADA = 'PRODUCAO_RESERVAS_ENTRADA';
const ABA_PRODUCAO_DESTINOS = 'PRODUCAO_DESTINOS';
const ABA_PRODUCAO_SAIDAS_LOTES = 'PRODUCAO_SAIDAS_LOTES';
const PRODUCAO_CACHE_SCOPE = 'PRODUCAO_LISTA_ATIVAS';
const PRODUCAO_CACHE_TTL_SEC = 90;

const PRODUCAO_SCHEMA = [
  'producao_id',
  'produto_id',
  'receita_id',
  'nome_ordem',
  'qtd_planejada',
  'qtd_produzida_acumulada',
  'qtd_restante',
  'status',
  'data_inicio',
  'data_prevista_termino',
  'data_conclusao',
  'observacao',
  'estoque_atualizado',
  'data_estoque_atualizado',
  'data_ultima_movimentacao_estoque',
  'ativo',
  'criado_em'
];

const PRODUCAO_ETAPAS_SCHEMA = [
  'id',
  'producao_id',
  'nome_etapa',
  'feito',
  'ordem',
  'data_feito',
  'ativo'
];

const PRODUCAO_CONSUMO_SCHEMA = [
  'id',
  'producao_id',
  'estoque_id',
  'quantidade_consumida',
  'valor_unit_snapshot',
  'total_snapshot',
  'criado_em',
  'ativo'
];

const PRODUCAO_MATERIAIS_SCHEMA = [
  'id',
  'producao_id',
  'estoque_id',
  'quantidade',
  'unidade',
  'criado_em',
  'ativo'
];

const PRODUCAO_MATERIAIS_PREVISTOS_SCHEMA = [
  'id',
  'producao_id',
  'estoque_id',
  'quantidade',
  'unidade',
  'item_snapshot',
  'valor_unit_snapshot',
  'criado_em',
  'ativo'
];

const PRODUCAO_VINCULOS_SCHEMA = [
  'id',
  'producao_id',
  'receita_id',
  'receita_entrada_id',
  'estoque_id',
  'tipo_item',
  'origem_item',
  'unidade',
  'quantidade_prevista',
  'quantidade_consumida',
  'item_snapshot',
  'valor_unit_snapshot',
  'status',
  'criado_em',
  'ativo'
];

const PRODUCAO_NECESSIDADES_ENTRADA_SCHEMA = [
  'id',
  'producao_id',
  'receita_id',
  'receita_entrada_id',
  'tipo_item',
  'categoria',
  'nome_item',
  'unidade',
  'modo_atendimento',
  'quantidade_prevista',
  'quantidade_baixada',
  'valor_unit_referencia',
  'comprimento_min_cm',
  'largura_min_cm',
  'espessura_min_cm',
  'serie_item',
  'origem_manual',
  'criado_em',
  'ativo'
];

const PRODUCAO_RESERVAS_ENTRADA_SCHEMA = [
  'id',
  'producao_id',
  'necessidade_id',
  'estoque_id',
  'quantidade_reservada',
  'quantidade_consumida',
  'reserva_exclusiva',
  'item_snapshot',
  'unidade_snapshot',
  'valor_unit_snapshot',
  'criado_em',
  'ativo'
];

const PRODUCAO_DESTINOS_SCHEMA = [
  'id',
  'producao_id',
  'receita_id',
  'receita_saida_id',
  'estoque_id',
  'nome_saida_snapshot',
  'unidade',
  'quantidade_prevista',
  'quantidade_produzida',
  'status',
  'criado_em',
  'ativo'
];

const PRODUCAO_SAIDAS_LOTES_SCHEMA = [
  'id',
  'producao_id',
  'estoque_id',
  'nome_saida',
  'produto_ref_id',
  'unidade',
  'quantidade',
  'custo_unitario',
  'custo_total',
  'criado_em',
  'ativo'
];

function formatDateSafe(value, pattern) {
  if (!value) return '';
  const dt = value instanceof Date ? value : new Date(value);
  if (!(dt instanceof Date) || isNaN(dt)) return '';
  const fmt = pattern || 'yyyy-MM-dd';
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), fmt);
}

function normalizarQuantidadeInteiraProducao(valor) {
  const n = parseNumeroBR(valor);
  if (!isFinite(n) || n <= 0) return 0;
  return Math.max(0, Math.round(n));
}

function validarQuantidadeInteiraProducao(valor, campoLabel) {
  const n = parseNumeroBR(valor);
  if (!isFinite(n) || n <= 0) {
    throw new Error(`${campoLabel} invalida.`);
  }
  if (Math.abs(n - Math.round(n)) > 0.000001) {
    throw new Error(`${campoLabel} deve ser inteira.`);
  }
  return Math.round(n);
}

function lerCacheProducao() {
  return appCacheGetJson(PRODUCAO_CACHE_SCOPE);
}

function salvarCacheProducao(lista) {
  appCachePutJson(PRODUCAO_CACHE_SCOPE, Array.isArray(lista) ? lista : [], PRODUCAO_CACHE_TTL_SEC);
}

function limparCacheProducao() {
  return appCacheRemove(PRODUCAO_CACHE_SCOPE);
}

function recarregarCacheProducao() {
  limparCacheProducao();
  const dados = listarProducao(true);
  return {
    ok: true,
    scope: PRODUCAO_CACHE_SCOPE,
    ttl_segundos: PRODUCAO_CACHE_TTL_SEC,
    total_itens: Array.isArray(dados) ? dados.length : 0
  };
}

function assertCanWriteProducao(acao) {
  assertCanWrite(acao || 'Operacao de producao');
}

function isTipoEntradaMadeiraProducao(tipo) {
  return String(tipo || '').trim().toUpperCase() === 'MADEIRA';
}

function getModoAtendimentoEntradaProducao(tipo) {
  return isTipoEntradaMadeiraProducao(tipo) ? 'PECA_UNICA' : 'QUANTIDADE';
}

function ensureSheetComSchemaProducao(nomeAba, schema) {
  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(nomeAba);
  if (!sheet) {
    sheet = ss.insertSheet(nomeAba);
  }
  ensureSchema(sheet, schema);
  return sheet;
}

function listarOrdensAtivasProducaoMap_(contextoLeitura) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO);
  if (!sheet) return {};

  const mapa = {};
  rowsToObjects(sheet, { context: contextoLeitura })
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(i => {
      const id = String(i.producao_id || '').trim();
      if (!id) return;
      mapa[id] = i;
    });
  return mapa;
}

function listarNecessidadesEntradaRows_(producaoId, opcoes) {
  const opts = opcoes || {};
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_NECESSIDADES_ENTRADA);
  if (!sheet) return [];

  const alvo = String(producaoId || '').trim();
  return rowsToObjects(sheet, { context: opts.context })
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => !alvo || String(i.producao_id || '').trim() === alvo);
}

function listarReservasEntradaRows_(producaoId, opcoes) {
  const opts = opcoes || {};
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_RESERVAS_ENTRADA);
  if (!sheet) return [];

  const alvo = String(producaoId || '').trim();
  return rowsToObjects(sheet, { context: opts.context })
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => !alvo || String(i.producao_id || '').trim() === alvo);
}

function listarVinculosMateriaisRowsLegados_(producaoId, opcoes) {
  const opts = opcoes || {};
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_VINCULOS);
  if (!sheet) return [];

  const alvo = String(producaoId || '').trim();
  return rowsToObjects(sheet, { context: opts.context })
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => !alvo || String(i.producao_id || '').trim() === alvo);
}

function migrarVinculosLegadosProducaoPorId(producaoId, opcoes) {
  assertCanWriteProducao('Migracao de vinculos legados da producao');
  const opts = opcoes || {};
  const contexto = opts.context || null;
  const prodId = String(producaoId || '').trim();
  if (!prodId) {
    throw new Error('Producao invalida para migracao.');
  }

  const necessidadesAtuais = listarNecessidadesEntradaRows_(prodId, { context: contexto });
  if (necessidadesAtuais.length > 0) {
    return {
      ok: true,
      migrado: false,
      producao_id: prodId,
      motivo: 'A OP ja utiliza a nova estrutura de necessidades.'
    };
  }

  const vinculosLegados = listarVinculosMateriaisRowsLegados_(prodId, { context: contexto });
  if (vinculosLegados.length === 0) {
    return {
      ok: true,
      migrado: false,
      producao_id: prodId,
      motivo: 'Nenhum vinculo legado encontrado.'
    };
  }

  ensureSheetComSchemaProducao(
    ABA_PRODUCAO_NECESSIDADES_ENTRADA,
    PRODUCAO_NECESSIDADES_ENTRADA_SCHEMA
  );
  ensureSheetComSchemaProducao(
    ABA_PRODUCAO_RESERVAS_ENTRADA,
    PRODUCAO_RESERVAS_ENTRADA_SCHEMA
  );

  const sheetEntradas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS_ENTRADAS);
  const entradasModelo = sheetEntradas ? rowsToObjects(sheetEntradas, { context: contexto }) : [];
  const entradasPorId = {};
  entradasModelo
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(i => {
      const id = String(i.id || '').trim();
      if (!id) return;
      entradasPorId[id] = i;
    });

  const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
  const estoqueRows = sheetEstoque ? rowsToObjects(sheetEstoque, { context: contexto }) : [];
  const estoquePorId = {};
  estoqueRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(i => {
      const id = String(i.ID || '').trim();
      if (!id) return;
      estoquePorId[id] = i;
    });

  const seriePorEntrada = {};
  let reservasCriadas = 0;

  vinculosLegados.forEach((v, idx) => {
    const receitaEntradaId = String(v.receita_entrada_id || '').trim();
    const entradaModelo = entradasPorId[receitaEntradaId] || {};
    const tipoItem = String(
      entradaModelo.tipo_item || v.tipo_item || ''
    ).trim().toUpperCase();
    const modoAtendimento = getModoAtendimentoEntradaProducao(tipoItem);
    const nomeItem = String(
      entradaModelo.nome_item ||
      v.item_snapshot ||
      v.origem_item ||
      v.estoque_id ||
      `Entrada ${idx + 1}`
    ).trim();
    const unidade = String(
      entradaModelo.unidade || v.unidade || (modoAtendimento === 'PECA_UNICA' ? 'M3' : '')
    ).trim();
    const quantidadePrevista = parseNumeroBR(v.quantidade_prevista);
    const quantidadeConsumida = parseNumeroBR(v.quantidade_consumida);
    const serieItem = modoAtendimento === 'PECA_UNICA'
      ? (seriePorEntrada[receitaEntradaId || `MANUAL_${idx}`] = (seriePorEntrada[receitaEntradaId || `MANUAL_${idx}`] || 0) + 1)
      : '';
    const necessidadeId = gerarId('PNE');

    insert(ABA_PRODUCAO_NECESSIDADES_ENTRADA, {
      id: necessidadeId,
      producao_id: prodId,
      receita_id: String(v.receita_id || '').trim(),
      receita_entrada_id: receitaEntradaId,
      tipo_item: tipoItem,
      categoria: String(entradaModelo.categoria || '').trim(),
      nome_item: nomeItem,
      unidade,
      modo_atendimento: modoAtendimento,
      quantidade_prevista: quantidadePrevista,
      quantidade_baixada: quantidadeConsumida,
      valor_unit_referencia: parseNumeroBR(
        entradaModelo.custo_manual || v.valor_unit_snapshot
      ),
      comprimento_min_cm: parseNumeroBR(entradaModelo.comprimento_cm),
      largura_min_cm: parseNumeroBR(entradaModelo.largura_cm),
      espessura_min_cm: parseNumeroBR(entradaModelo.espessura_cm),
      serie_item: serieItem,
      origem_manual: !receitaEntradaId,
      criado_em: v.criado_em || new Date(),
      ativo: true
    }, PRODUCAO_NECESSIDADES_ENTRADA_SCHEMA);

    const estoqueId = String(v.estoque_id || '').trim();
    if (!estoqueId) return;

    const itemEstoque = estoquePorId[estoqueId] || {};
    const quantidadeReservadaBase = Math.max(quantidadePrevista, quantidadeConsumida);
    const quantidadeReservada = modoAtendimento === 'PECA_UNICA'
      ? Math.max(parseNumeroBR(itemEstoque.quantidade), quantidadeReservadaBase)
      : quantidadeReservadaBase;

    insert(ABA_PRODUCAO_RESERVAS_ENTRADA, {
      id: gerarId('PRE'),
      producao_id: prodId,
      necessidade_id: necessidadeId,
      estoque_id: estoqueId,
      quantidade_reservada: quantidadeReservada,
      quantidade_consumida: quantidadeConsumida,
      reserva_exclusiva: modoAtendimento === 'PECA_UNICA',
      item_snapshot: String(
        v.item_snapshot || itemEstoque.item || nomeItem || estoqueId
      ).trim(),
      unidade_snapshot: String(itemEstoque.unidade || unidade || '').trim(),
      valor_unit_snapshot: parseNumeroBR(
        v.valor_unit_snapshot || itemEstoque.custo_unitario || itemEstoque.valor_unit
      ),
      criado_em: v.criado_em || new Date(),
      ativo: true
    }, PRODUCAO_RESERVAS_ENTRADA_SCHEMA);
    reservasCriadas += 1;
  });

  const materiais = recalcularMateriaisPrevistosFromVinculos(prodId);
  return {
    ok: true,
    migrado: true,
    producao_id: prodId,
    necessidades_criadas: vinculosLegados.length,
    reservas_criadas: reservasCriadas,
    materiaisPrevistos: materiais.itens || [],
    custoPrevisto: parseNumeroBR(materiais.custoPrevisto)
  };
}

function migrarVinculosLegadosProducao() {
  assertCanWriteProducao('Migracao de vinculos legados de producao');
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO);
  if (!sheet) {
    return { ok: true, total_ops: 0, migradas: 0, ignoradas: 0, detalhes: [] };
  }

  const ordens = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .map(i => String(i.producao_id || '').trim())
    .filter(i => i);

  const detalhes = [];
  let migradas = 0;
  let ignoradas = 0;

  ordens.forEach(producaoId => {
    const resultado = migrarVinculosLegadosProducaoPorId(producaoId);
    detalhes.push(resultado);
    if (resultado?.migrado) {
      migradas += 1;
    } else {
      ignoradas += 1;
    }
  });

  return {
    ok: true,
    total_ops: ordens.length,
    migradas,
    ignoradas,
    detalhes
  };
}

function limparNecessidadesEntradaProducao(producaoId, opcoes) {
  assertCanWriteProducao('Limpeza de necessidades da producao');
  const cfg = opcoes || {};
  const preservarManuais = !!cfg.preservarManuais;

  const sheet = ensureSheetComSchemaProducao(
    ABA_PRODUCAO_NECESSIDADES_ENTRADA,
    PRODUCAO_NECESSIDADES_ENTRADA_SCHEMA
  );

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idxProducao = headers.indexOf('producao_id');
  const idxOrigemManual = headers.indexOf('origem_manual');
  const idxAtivo = headers.indexOf('ativo');

  if (idxProducao === -1 || idxAtivo === -1 || data.length <= 1) return;

  const producaoIdAlvo = String(producaoId || '').trim();
  const linhasAlvo = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxProducao] || '').trim() !== producaoIdAlvo) continue;
    if (String(data[i][idxAtivo]).toLowerCase() === 'false') continue;
    if (preservarManuais && idxOrigemManual !== -1 && String(data[i][idxOrigemManual]).toLowerCase() === 'true') {
      continue;
    }
    linhasAlvo.push(i + 1);
  }

  if (linhasAlvo.length === 0) return;
  linhasAlvo.forEach(rowNumber => {
    sheet.getRange(rowNumber, idxAtivo + 1).setValue(false);
  });
}

function limparReservasEntradaProducao(producaoId, opcoes) {
  assertCanWriteProducao('Limpeza de reservas da producao');
  const cfg = opcoes || {};
  const preservarConsumidas = !!cfg.preservarConsumidas;

  const sheet = ensureSheetComSchemaProducao(
    ABA_PRODUCAO_RESERVAS_ENTRADA,
    PRODUCAO_RESERVAS_ENTRADA_SCHEMA
  );

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idxProducao = headers.indexOf('producao_id');
  const idxQtdConsumida = headers.indexOf('quantidade_consumida');
  const idxAtivo = headers.indexOf('ativo');

  if (idxProducao === -1 || idxAtivo === -1 || data.length <= 1) return;

  const producaoIdAlvo = String(producaoId || '').trim();
  const linhasAlvo = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxProducao] || '').trim() !== producaoIdAlvo) continue;
    if (String(data[i][idxAtivo]).toLowerCase() === 'false') continue;
    if (preservarConsumidas && idxQtdConsumida !== -1 && parseNumeroBR(data[i][idxQtdConsumida]) > 0) {
      continue;
    }
    linhasAlvo.push(i + 1);
  }

  if (linhasAlvo.length === 0) return;
  linhasAlvo.forEach(rowNumber => {
    sheet.getRange(rowNumber, idxAtivo + 1).setValue(false);
  });
}

function gerarNecessidadesEntradaProducao(produtoId, receitaId, qtdPlanejada) {
  const prodId = String(produtoId || '').trim();
  const recId = String(receitaId || '').trim();
  const qtdPlanejadaInt = normalizarQuantidadeInteiraProducao(qtdPlanejada);

  if (!prodId || !recId || qtdPlanejadaInt <= 0) return [];

  const sheetEntradas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS_ENTRADAS);
  const entradas = sheetEntradas ? rowsToObjects(sheetEntradas) : [];
  const entradasAtivas = entradas
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => String(i.receita_id || '').trim() === recId);

  const necessidades = [];
  entradasAtivas.forEach((entrada, idx) => {
    const tipoItem = String(entrada.tipo_item || '').trim().toUpperCase();
    const categoria = String(entrada.categoria || '').trim();
    const nomeItem = String(entrada.nome_item || '').trim();
    const unidade = String(entrada.unidade || '').trim() || (isTipoEntradaMadeiraProducao(tipoItem) ? 'M3' : '');
    const quantidadeBase = parseNumeroBR(entrada.qtd_pecas);
    const valorUnitRef = parseNumeroBR(entrada.custo_manual);
    const comprimentoMin = parseNumeroBR(entrada.comprimento_cm);
    const larguraMin = parseNumeroBR(entrada.largura_cm);
    const espessuraMin = parseNumeroBR(entrada.espessura_cm);
    const modoAtendimento = getModoAtendimentoEntradaProducao(tipoItem);

    if (!tipoItem || !nomeItem) return;

    if (modoAtendimento === 'PECA_UNICA') {
      for (let serie = 1; serie <= qtdPlanejadaInt; serie++) {
        necessidades.push({
          id: gerarId('PNE'),
          producao_id: '',
          receita_id: recId,
          receita_entrada_id: String(entrada.id || '').trim(),
          tipo_item: tipoItem,
          categoria,
          nome_item: nomeItem,
          unidade: unidade || 'M3',
          modo_atendimento: modoAtendimento,
          quantidade_prevista: quantidadeBase > 0 ? quantidadeBase : 0,
          quantidade_baixada: 0,
          valor_unit_referencia: valorUnitRef,
          comprimento_min_cm: comprimentoMin,
          largura_min_cm: larguraMin,
          espessura_min_cm: espessuraMin,
          serie_item: serie,
          origem_manual: false,
          criado_em: new Date(),
          ativo: true,
          _ordem_base: idx
        });
      }
      return;
    }

    const quantidadePrevista = quantidadeBase * qtdPlanejadaInt;
    if (quantidadePrevista <= 0) return;

    necessidades.push({
      id: gerarId('PNE'),
      producao_id: '',
      receita_id: recId,
      receita_entrada_id: String(entrada.id || '').trim(),
      tipo_item: tipoItem,
      categoria,
      nome_item: nomeItem,
      unidade,
      modo_atendimento: modoAtendimento,
      quantidade_prevista: quantidadePrevista,
      quantidade_baixada: 0,
      valor_unit_referencia: valorUnitRef,
      comprimento_min_cm: 0,
      largura_min_cm: 0,
      espessura_min_cm: 0,
      serie_item: '',
      origem_manual: false,
      criado_em: new Date(),
      ativo: true,
      _ordem_base: idx
    });
  });

  return necessidades;
}

function salvarNecessidadesEntradaProducao(producaoId, receitaId, necessidades, opcoes) {
  assertCanWriteProducao('Salvamento de necessidades da producao');
  const cfg = opcoes || {};
  const preservarManuais = !!cfg.preservarManuais;
  const prodId = String(producaoId || '').trim();
  if (!prodId) return true;

  ensureSheetComSchemaProducao(
    ABA_PRODUCAO_NECESSIDADES_ENTRADA,
    PRODUCAO_NECESSIDADES_ENTRADA_SCHEMA
  );
  limparNecessidadesEntradaProducao(prodId, { preservarManuais });

  const linhas = Array.isArray(necessidades) ? necessidades : [];
  linhas.forEach(n => {
    const quantidadePrevista = parseNumeroBR(n.quantidade_prevista);
    if (quantidadePrevista <= 0 && String(n.modo_atendimento || '').trim().toUpperCase() !== 'PECA_UNICA') return;

    insert(ABA_PRODUCAO_NECESSIDADES_ENTRADA, {
      id: String(n.id || gerarId('PNE')).trim(),
      producao_id: prodId,
      receita_id: String(receitaId || n.receita_id || '').trim(),
      receita_entrada_id: String(n.receita_entrada_id || '').trim(),
      tipo_item: String(n.tipo_item || '').trim().toUpperCase(),
      categoria: String(n.categoria || '').trim(),
      nome_item: String(n.nome_item || '').trim(),
      unidade: String(n.unidade || '').trim(),
      modo_atendimento: String(n.modo_atendimento || '').trim().toUpperCase() || 'QUANTIDADE',
      quantidade_prevista: quantidadePrevista,
      quantidade_baixada: parseNumeroBR(n.quantidade_baixada),
      valor_unit_referencia: parseNumeroBR(n.valor_unit_referencia),
      comprimento_min_cm: parseNumeroBR(n.comprimento_min_cm),
      largura_min_cm: parseNumeroBR(n.largura_min_cm),
      espessura_min_cm: parseNumeroBR(n.espessura_min_cm),
      serie_item: String(n.serie_item || '').trim(),
      origem_manual: String(n.origem_manual).toLowerCase() === 'true' || n.origem_manual === true,
      criado_em: n.criado_em || new Date(),
      ativo: true
    }, PRODUCAO_NECESSIDADES_ENTRADA_SCHEMA);
  });

  invalidarCachesRelacionadosAba(ABA_PRODUCAO);
  return true;
}

function listarResumoReservasEstoqueProducao(opcoes) {
  const opts = opcoes || {};
  const contexto = opts.context || null;
  const excluirReservaId = String(opts.excluirReservaId || '').trim();
  const producoesAtivas = listarOrdensAtivasProducaoMap_(contexto);
  const reservas = listarReservasEntradaRows_('', { context: contexto });
  const resumo = {};

  reservas.forEach(r => {
    const producaoId = String(r.producao_id || '').trim();
    if (!producoesAtivas[producaoId]) return;
    const reservaId = String(r.id || '').trim();
    if (excluirReservaId && reservaId === excluirReservaId) return;

    const estoqueId = String(r.estoque_id || '').trim();
    if (!estoqueId) return;

    const reservado = Math.max(
      parseNumeroBR(r.quantidade_reservada) - parseNumeroBR(r.quantidade_consumida),
      0
    );
    if (reservado <= 0) return;

    if (!resumo[estoqueId]) {
      resumo[estoqueId] = {
        quantidade_reservada: 0,
        producoes: {}
      };
    }
    resumo[estoqueId].quantidade_reservada += reservado;
    resumo[estoqueId].producoes[producaoId] = true;
  });

  return resumo;
}

function listarEstoqueAtivoComReservaProducao(opcoes) {
  const opts = opcoes || {};
  const contexto = opts.context || null;
  const resumoReservas = listarResumoReservasEstoqueProducao({
    context: contexto,
    excluirReservaId: opts.excluirReservaId
  });

  return (typeof listarEstoque === 'function' ? listarEstoque(!!opts.forcarRecarregar) : [])
    .map(item => {
      const estoqueId = String(item?.ID || '').trim();
      const reservado = parseNumeroBR(resumoReservas[estoqueId]?.quantidade_reservada);
      const quantidade = parseNumeroBR(item?.quantidade);
      return {
        ...item,
        reservado_quantidade: reservado,
        quantidade_disponivel: Math.max(quantidade - reservado, 0),
        reservado_em_op_total: Object.keys(resumoReservas[estoqueId]?.producoes || {}).length
      };
    });
}

function obterNecessidadeEntradaProducao_(producaoId, necessidadeId, contexto) {
  const prodId = String(producaoId || '').trim();
  const necId = String(necessidadeId || '').trim();
  if (!prodId || !necId) return null;

  const necessidade = listarNecessidadesEntradaRows_(prodId, { context: contexto })
    .find(i => String(i.id || '').trim() === necId);
  return necessidade || null;
}

function obterReservaEntradaProducao_(producaoId, reservaId, contexto) {
  const prodId = String(producaoId || '').trim();
  const resId = String(reservaId || '').trim();
  if (!prodId || !resId) return null;

  const reserva = listarReservasEntradaRows_(prodId, { context: contexto })
    .find(i => String(i.id || '').trim() === resId);
  return reserva || null;
}

function itemEstoqueCompativelComNecessidade_(necessidade, itemEstoque) {
  const necessidadeTipo = String(necessidade?.tipo_item || '').trim().toUpperCase();
  const necessidadeCategoria = String(necessidade?.categoria || '').trim().toUpperCase();
  const necessidadeUnidade = String(necessidade?.unidade || '').trim().toUpperCase();
  const estoqueTipo = String(itemEstoque?.tipo || '').trim().toUpperCase();
  const estoqueCategoria = String(itemEstoque?.categoria || '').trim().toUpperCase();
  const estoqueUnidade = String(itemEstoque?.unidade || '').trim().toUpperCase();

  if (!necessidadeTipo || !estoqueTipo || necessidadeTipo !== estoqueTipo) return false;
  if (necessidadeCategoria && estoqueCategoria !== necessidadeCategoria) return false;
  if (necessidadeUnidade && estoqueUnidade !== necessidadeUnidade) return false;

  const modo = String(necessidade?.modo_atendimento || '').trim().toUpperCase();
  if (modo !== 'PECA_UNICA') return true;

  const comprimentoMin = parseNumeroBR(necessidade?.comprimento_min_cm);
  const larguraMin = parseNumeroBR(necessidade?.largura_min_cm);
  const espessuraMin = parseNumeroBR(necessidade?.espessura_min_cm);
  const comprimentoEstoque = parseNumeroBR(itemEstoque?.comprimento_cm);
  const larguraEstoque = parseNumeroBR(itemEstoque?.largura_cm);
  const espessuraEstoque = parseNumeroBR(itemEstoque?.espessura_cm);

  return (
    comprimentoEstoque >= comprimentoMin &&
    larguraEstoque >= larguraMin &&
    espessuraEstoque >= espessuraMin
  );
}

function calcularResumoNecessidadeEntrada_(necessidade, reservas, estoqueMap) {
  const modo = String(necessidade?.modo_atendimento || '').trim().toUpperCase() || 'QUANTIDADE';
  const quantidadePrevista = parseNumeroBR(necessidade?.quantidade_prevista);
  const quantidadeBaixadaAtual = parseNumeroBR(necessidade?.quantidade_baixada);

  let quantidadeReservada = 0;
  const reservasNormalizadas = (Array.isArray(reservas) ? reservas : []).map(reserva => {
    const estoqueId = String(reserva?.estoque_id || '').trim();
    const estoque = estoqueMap[estoqueId] || {};
    const quantidadeReservadaLinha = parseNumeroBR(reserva?.quantidade_reservada);
    const quantidadeConsumidaLinha = parseNumeroBR(reserva?.quantidade_consumida);
    const quantidadeRestanteLinha = Math.max(quantidadeReservadaLinha - quantidadeConsumidaLinha, 0);
    quantidadeReservada += quantidadeRestanteLinha;
    return {
      id: String(reserva?.id || '').trim(),
      necessidade_id: String(reserva?.necessidade_id || '').trim(),
      estoque_id: estoqueId,
      item: String(estoque?.item || reserva?.item_snapshot || estoqueId).trim(),
      unidade: String(estoque?.unidade || reserva?.unidade_snapshot || necessidade?.unidade || '').trim(),
      tipo_item: String(estoque?.tipo || necessidade?.tipo_item || '').trim().toUpperCase(),
      categoria: String(estoque?.categoria || necessidade?.categoria || '').trim(),
      quantidade_reservada: quantidadeReservadaLinha,
      quantidade_consumida: quantidadeConsumidaLinha,
      quantidade_restante: quantidadeRestanteLinha,
      valor_unit: parseNumeroBR(estoque?.custo_unitario || estoque?.valor_unit || reserva?.valor_unit_snapshot),
      quantidade_disponivel_item: parseNumeroBR(estoque?.quantidade_disponivel),
      quantidade_fisica_item: parseNumeroBR(estoque?.quantidade)
    };
  });

  const quantidadePendente = modo === 'PECA_UNICA'
    ? (reservasNormalizadas.length > 0 ? 0 : quantidadePrevista)
    : Math.max(quantidadePrevista - quantidadeReservada, 0);
  const quantidadeBaixada = modo === 'PECA_UNICA'
    ? (quantidadeBaixadaAtual > 0 ? quantidadePrevista : 0)
    : quantidadeBaixadaAtual;

  let status = 'Nao reservado';
  if (modo === 'PECA_UNICA') {
    if (quantidadeBaixada > 0) {
      status = 'Baixado';
    } else if (reservasNormalizadas.length > 0) {
      status = 'Reservado';
    }
  } else if (quantidadeBaixada >= quantidadePrevista && quantidadePrevista > 0) {
    status = 'Baixado';
  } else if (quantidadeReservada >= quantidadePrevista && quantidadePrevista > 0) {
    status = 'Reservado';
  } else if (quantidadeReservada > 0) {
    status = 'Parcial';
  }

  return {
    id: String(necessidade?.id || '').trim(),
    producao_id: String(necessidade?.producao_id || '').trim(),
    receita_id: String(necessidade?.receita_id || '').trim(),
    receita_entrada_id: String(necessidade?.receita_entrada_id || '').trim(),
    tipo_item: String(necessidade?.tipo_item || '').trim().toUpperCase(),
    categoria: String(necessidade?.categoria || '').trim(),
    origem_item: String(necessidade?.nome_item || '').trim(),
    item: String(necessidade?.nome_item || '').trim(),
    unidade: String(necessidade?.unidade || '').trim(),
    modo_atendimento: modo,
    quantidade_prevista: quantidadePrevista,
    quantidade_reservada: quantidadeReservada,
    quantidade_baixada: quantidadeBaixada,
    quantidade_pendente: quantidadePendente,
    valor_unit_referencia: parseNumeroBR(necessidade?.valor_unit_referencia),
    comprimento_min_cm: parseNumeroBR(necessidade?.comprimento_min_cm),
    largura_min_cm: parseNumeroBR(necessidade?.largura_min_cm),
    espessura_min_cm: parseNumeroBR(necessidade?.espessura_min_cm),
    serie_item: String(necessidade?.serie_item || '').trim(),
    origem_manual: String(necessidade?.origem_manual).toLowerCase() === 'true',
    status,
    reservas: reservasNormalizadas
  };
}

function listarProdutosAtivosIndicesProducao() {
  const produtosSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS);
  const produtosRows = produtosSheet ? rowsToObjects(produtosSheet) : [];
  const porId = {};
  const porNome = {};

  produtosRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(p => {
      const produtoId = String(p.produto_id || '').trim();
      if (!produtoId) return;

      const precoVenda = parseNumeroBR(p.preco_venda);
      const registro = {
        produto_id: produtoId,
        nome_produto: String(p.nome_produto || '').trim(),
        preco_venda: (isFinite(precoVenda) && precoVenda > 0)
          ? Number(precoVenda.toFixed(2))
          : ''
      };
      porId[produtoId] = registro;

      const nomeNorm = normalizarTextoSemAcentoProducao(p.nome_produto);
      if (!nomeNorm) return;

      const atual = porNome[nomeNorm];
      if (!atual) {
        porNome[nomeNorm] = registro;
        return;
      }

      const precoAtual = parseNumeroBR(atual.preco_venda);
      const precoProximo = parseNumeroBR(registro.preco_venda);
      if (precoAtual <= 0 && precoProximo > 0) {
        porNome[nomeNorm] = registro;
      }
    });

  return { porId, porNome };
}

function listarProdutosAtivosMapeadosPorNomeProducao() {
  return listarProdutosAtivosIndicesProducao().porNome || {};
}

function listarProdutosAtivosMapeadosPorIdProducao() {
  return listarProdutosAtivosIndicesProducao().porId || {};
}

function obterProdutoAtivoPorIdSaidaProducao(produtoId, mapaProdutosPorId) {
  const id = String(produtoId || '').trim();
  if (!id) return null;

  const mapa = (mapaProdutosPorId && typeof mapaProdutosPorId === 'object')
    ? mapaProdutosPorId
    : listarProdutosAtivosMapeadosPorIdProducao();

  return mapa[id] || null;
}

function obterProdutoAtivoPorNomeSaidaProducao(nomeSaida, mapaProdutosPorNome) {
  const nomeNorm = normalizarTextoSemAcentoProducao(nomeSaida);
  if (!nomeNorm) return null;

  const mapa = (mapaProdutosPorNome && typeof mapaProdutosPorNome === 'object')
    ? mapaProdutosPorNome
    : listarProdutosAtivosMapeadosPorNomeProducao();

  return mapa[nomeNorm] || null;
}

function resolverProdutoSaidaProducao(produtoRefId, nomeSaida, mapas) {
  const porId = mapas?.porId || {};
  const porNome = mapas?.porNome || {};
  const produtoPorId = obterProdutoAtivoPorIdSaidaProducao(produtoRefId, porId);

  if (produtoPorId) {
    return {
      produto: produtoPorId,
      vinculoPorId: true
    };
  }

  const possuiRefId = String(produtoRefId || '').trim() !== '';
  if (possuiRefId) {
    return {
      produto: null,
      vinculoPorId: false
    };
  }

  return {
    produto: obterProdutoAtivoPorNomeSaidaProducao(nomeSaida, porNome),
    vinculoPorId: false
  };
}

function listarProducao(forcarRecarregar) {
  if (!forcarRecarregar) {
    const cached = lerCacheProducao();
    if (Array.isArray(cached)) {
      return cached;
    }
  }

  const ss = getDataSpreadsheet();
  const contextoLeitura = criarContextoLeituraRows();
  const sheet = ss.getSheetByName(ABA_PRODUCAO);
  if (!sheet) {
    salvarCacheProducao([]);
    return [];
  }

  const produtosSheet = ss.getSheetByName(ABA_PRODUTOS);
  const produtos = produtosSheet ? rowsToObjects(produtosSheet, { context: contextoLeitura }) : [];
  const produtosMap = {};
  const produtosAtivosPorNome = {};
  produtos
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(p => {
      produtosMap[p.produto_id] = p;

      const nomeNorm = normalizarTextoSemAcentoProducao(p.nome_produto);
      if (!nomeNorm) return;

      const precoVenda = parseNumeroBR(p.preco_venda);
      const proximo = {
        produto_id: String(p.produto_id || '').trim(),
        nome_produto: String(p.nome_produto || '').trim(),
        preco_venda: (isFinite(precoVenda) && precoVenda > 0)
          ? Number(precoVenda.toFixed(2))
          : ''
      };

      const atual = produtosAtivosPorNome[nomeNorm];
      if (!atual) {
        produtosAtivosPorNome[nomeNorm] = proximo;
        return;
      }

      const precoAtual = parseNumeroBR(atual.preco_venda);
      const precoProximo = parseNumeroBR(proximo.preco_venda);
      if (precoAtual <= 0 && precoProximo > 0) {
        produtosAtivosPorNome[nomeNorm] = proximo;
      }
    });

  const receitasSheet = ss.getSheetByName(ABA_PRODUTOS_RECEITAS);
  const receitasRows = receitasSheet ? rowsToObjects(receitasSheet, { context: contextoLeitura }) : [];
  const receitasMap = {};
  receitasRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(r => {
      receitasMap[r.receita_id] = r;
    });

  const saidasReceitaSheet = ss.getSheetByName(ABA_PRODUTOS_RECEITAS_SAIDAS);
  const saidasReceitaRows = saidasReceitaSheet ? rowsToObjects(saidasReceitaSheet, { context: contextoLeitura }) : [];
  const saidasReceitaMap = {};
  saidasReceitaRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(s => {
      const receitaId = String(s.receita_id || '').trim();
      if (!receitaId) return;
      if (!saidasReceitaMap[receitaId]) saidasReceitaMap[receitaId] = [];
      saidasReceitaMap[receitaId].push({
        nome_saida: String(s.nome_saida || '').trim(),
        produto_ref_id: String(s.produto_ref_id || '').trim(),
        tipo_item: String(s.tipo_item || '').trim().toUpperCase() || 'PRODUTO',
        categoria: String(s.categoria || '').trim(),
        unidade: String(s.unidade || '').trim(),
        quantidade_base: parseNumeroBR(s.quantidade)
      });
    });

  const etapasSheet = ss.getSheetByName(ABA_PRODUCAO_ETAPAS);
  const etapasRows = etapasSheet ? rowsToObjects(etapasSheet, { context: contextoLeitura }) : [];
  const etapasMap = {};
  etapasRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(e => {
      const etapa = {
        ...e,
        feito: String(e.feito).toLowerCase() === 'true',
        data_feito: formatDateSafe(e.data_feito, 'yyyy-MM-dd')
      };
      if (!etapasMap[etapa.producao_id]) etapasMap[etapa.producao_id] = [];
      etapasMap[etapa.producao_id].push(etapa);
    });

  Object.keys(etapasMap).forEach(id => {
    etapasMap[id].sort((a, b) => parseNumeroBR(a.ordem) - parseNumeroBR(b.ordem));
  });

  const rows = rowsToObjects(sheet, { context: contextoLeitura });
  const lista = rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .map(i => {
      const prod = produtosMap[i.produto_id] || {};
      const receita = receitasMap[i.receita_id] || {};
      const qtdPlanejada = normalizarQuantidadeInteiraProducao(i.qtd_planejada);
      const estoqueAtualizado = String(i.estoque_atualizado).toLowerCase() === 'true';
      let qtdProduzidaAcumulada = normalizarQuantidadeInteiraProducao(i.qtd_produzida_acumulada);
      if (qtdProduzidaAcumulada <= 0 && estoqueAtualizado && qtdPlanejada > 0) {
        qtdProduzidaAcumulada = qtdPlanejada;
      }
      if (qtdProduzidaAcumulada < 0) {
        qtdProduzidaAcumulada = 0;
      }

      const qtdRestanteRaw = String(i.qtd_restante || '').trim();
      const qtdRestanteInformada = normalizarQuantidadeInteiraProducao(qtdRestanteRaw);
      let qtdRestante = qtdRestanteRaw !== ''
        ? qtdRestanteInformada
        : Math.max(qtdPlanejada - qtdProduzidaAcumulada, 0);
      if (qtdRestante < 0) qtdRestante = 0;
      if (qtdRestante <= 0) {
        qtdRestante = 0;
      }

      const saidasBaseReceita = Array.isArray(saidasReceitaMap[i.receita_id])
        ? saidasReceitaMap[i.receita_id]
        : [];
      const totalSaidasProdutoBase = saidasBaseReceita
        .map(s => normalizarTipoSaidaProducao(s?.tipo_item))
        .filter(tipo => tipo === 'PRODUTO')
        .length;
      const saidasPrevistasBase = agruparSaidasReceitaParaEstoque(
        saidasBaseReceita
          .map(s => {
            const quantidade = parseNumeroBR(s.quantidade_base) * qtdPlanejada;
            if (quantidade <= 0) return null;
            const nomeSaida = s.nome_saida || prod.nome_produto || 'Saida';
            const tipoSaida = s.tipo_item || 'PRODUTO';
            let produtoRefId = String(s.produto_ref_id || '').trim();
            if (!produtoRefId && normalizarTipoSaidaProducao(tipoSaida) === 'PRODUTO') {
              const nomeSaidaNorm = normalizarTextoSemAcentoProducao(nomeSaida);
              const nomeProdutoNorm = normalizarTextoSemAcentoProducao(prod.nome_produto);
              if ((totalSaidasProdutoBase === 1 || (nomeSaidaNorm && nomeSaidaNorm === nomeProdutoNorm)) && String(prod.produto_id || '').trim()) {
                produtoRefId = String(prod.produto_id || '').trim();
              }
            }
            return {
              nome_saida: nomeSaida,
              produto_ref_id: produtoRefId,
              tipo_item: tipoSaida,
              categoria: s.categoria || '',
              unidade: s.unidade || '',
              quantidade
            };
          })
          .filter(v => !!v)
      );
      const saidasPrevistas = saidasPrevistasBase.map(s => {
        const produtoResolvido = resolverProdutoSaidaProducao(
          s.produto_ref_id,
          s.nome_saida,
          { porId: produtosMap, porNome: produtosAtivosPorNome }
        );
        const produtoVinculado = produtoResolvido.produto;
        return {
          ...s,
          produto_ref_id: String(s.produto_ref_id || '').trim(),
          produto_id_vinculado: produtoResolvido.vinculoPorId ? (produtoVinculado?.produto_id || '') : '',
          vinculo_produto_por_id: produtoResolvido.vinculoPorId,
          preco_venda_produto: produtoVinculado?.preco_venda || '',
          permite_editar_preco_venda: produtoResolvido.vinculoPorId
        };
      });
      const totalSaidasPrevistas = saidasPrevistas
        .reduce((acc, s) => acc + parseNumeroBR(s.quantidade), 0);
      const valorPrevistoVendaPecas = saidasPrevistas
        .reduce((acc, s) => {
          const precoVenda = parseNumeroBR(s.preco_venda_produto);
          if (!isFinite(precoVenda) || precoVenda <= 0) return acc;

          const quantidade = parseNumeroBR(s.quantidade);
          if (!isFinite(quantidade) || quantidade <= 0) return acc;

          return acc + (quantidade * precoVenda);
        }, 0);

      return {
        ...i,
        nome_produto: prod.nome_produto || '',
        unidade_produto: prod.unidade_produto || '',
        receita_nome: receita.nome_receita || '',
        qtd_planejada: qtdPlanejada,
        qtd_produzida_acumulada: qtdProduzidaAcumulada,
        qtd_restante: qtdRestante,
        saidas_previstas: saidasPrevistas,
        saidas_previstas_tipos: saidasPrevistas.length,
        saidas_previstas_total: parseNumeroBR(totalSaidasPrevistas),
        valor_previsto_venda_pecas: Number(valorPrevistoVendaPecas.toFixed(2)),
        estoque_atualizado: estoqueAtualizado,
        criado_em: i.criado_em
          ? formatDateSafe(i.criado_em, 'yyyy-MM-dd HH:mm')
          : '',
        data_inicio: i.data_inicio
          ? formatDateSafe(i.data_inicio, 'yyyy-MM-dd')
          : '',
        data_prevista_termino: i.data_prevista_termino
          ? formatDateSafe(i.data_prevista_termino, 'yyyy-MM-dd')
          : '',
        data_conclusao: i.data_conclusao
          ? formatDateSafe(i.data_conclusao, 'yyyy-MM-dd')
          : '',
        data_estoque_atualizado: i.data_estoque_atualizado
          ? formatDateSafe(i.data_estoque_atualizado, 'yyyy-MM-dd')
          : '',
        data_ultima_movimentacao_estoque: i.data_ultima_movimentacao_estoque
          ? formatDateSafe(i.data_ultima_movimentacao_estoque, 'yyyy-MM-dd')
          : '',
        etapas: etapasMap[i.producao_id] || []
      };
    });

  salvarCacheProducao(lista);
  return lista;
}

function criarProducao(payload) {
  assertCanWriteProducao('Criacao de producao');
  const qtdPlanejada = validarQuantidadeInteiraProducao(payload.qtd_planejada, 'Quantidade planejada');
  const novo = {
    ...payload,
    producao_id: gerarId('OP'),
    status: payload.status || 'Em planejamento',
    qtd_planejada: qtdPlanejada,
    qtd_produzida_acumulada: 0,
    qtd_restante: qtdPlanejada > 0 ? qtdPlanejada : 0,
    estoque_atualizado: false,
    data_estoque_atualizado: '',
    data_ultima_movimentacao_estoque: '',
    ativo: true,
    criado_em: new Date()
  };

  const necessidadesEntrada = gerarNecessidadesEntradaProducao(
    novo.produto_id,
    novo.receita_id,
    novo.qtd_planejada
  );
  const saidasDetalhadas = explodirSaidasReceitaDetalhada(
    novo.produto_id,
    novo.receita_id,
    novo.qtd_planejada
  );

  insert(ABA_PRODUCAO, novo, PRODUCAO_SCHEMA);
  salvarNecessidadesEntradaProducao(novo.producao_id, novo.receita_id, necessidadesEntrada);
  salvarDestinosProducao(novo.producao_id, novo.receita_id, saidasDetalhadas.itens || []);
  const materiaisCalculados = recalcularMateriaisPrevistosFromVinculos(novo.producao_id);

  let nomeProduto = '';
  let unidadeProduto = '';
  const produtosSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS);
  if (produtosSheet) {
    const produtos = rowsToObjects(produtosSheet);
    const prod = produtos.find(p => p.produto_id === novo.produto_id);
    if (prod) {
      nomeProduto = prod.nome_produto || '';
      unidadeProduto = prod.unidade_produto || '';
    }
  }

  let receitaNome = '';
  const receitasSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (receitasSheet) {
    const receitas = rowsToObjects(receitasSheet);
    const rec = receitas.find(r => r.receita_id === novo.receita_id);
    if (rec) {
      receitaNome = rec.nome_receita || '';
    }
  }

  const etapasTemplate = listarEtapasProduto(novo.produto_id) || [];
  const etapasCriadas = etapasTemplate.map(etapa => {
    const nova = {
      id: gerarId('EPR'),
      producao_id: novo.producao_id,
      nome_etapa: etapa.nome_etapa || '',
      feito: false,
      ordem: etapa.ordem || '',
      data_feito: '',
      ativo: true
    };
    insert(ABA_PRODUCAO_ETAPAS, nova, PRODUCAO_ETAPAS_SCHEMA);
    return nova;
  });

  return {
    ...novo,
    nome_produto: nomeProduto,
    unidade_produto: unidadeProduto,
    receita_nome: receitaNome,
    custo_previsto: parseNumeroBR(materiaisCalculados.custoPrevisto),
    criado_em: Utilities.formatDate(novo.criado_em, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    data_inicio: novo.data_inicio || '',
    data_prevista_termino: novo.data_prevista_termino || '',
    data_conclusao: novo.data_conclusao || '',
    etapas: etapasCriadas
  };
}

function atualizarProducao(id, payload) {
  assertCanWriteProducao('Atualizacao de producao');
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO);
  if (!sheet) return false;

  const rows = rowsToObjects(sheet);
  const atual = rows.find(i => i.producao_id === id);
  if (!atual) return false;

  const payloadPersistencia = { ...(payload || {}) };
  if (Object.prototype.hasOwnProperty.call(payloadPersistencia, 'qtd_planejada')) {
    payloadPersistencia.qtd_planejada = validarQuantidadeInteiraProducao(
      payloadPersistencia.qtd_planejada,
      'Quantidade planejada'
    );
  }

  const dadosAtualizados = {
    ...atual,
    ...payloadPersistencia
  };

  const mudouReceita = Object.prototype.hasOwnProperty.call(payloadPersistencia, 'receita_id') &&
    payloadPersistencia.receita_id !== atual.receita_id;
  const mudouProduto = Object.prototype.hasOwnProperty.call(payloadPersistencia, 'produto_id') &&
    payloadPersistencia.produto_id !== atual.produto_id;
  const mudouQtd = Object.prototype.hasOwnProperty.call(payloadPersistencia, 'qtd_planejada') &&
    normalizarQuantidadeInteiraProducao(payloadPersistencia.qtd_planejada) !== normalizarQuantidadeInteiraProducao(atual.qtd_planejada);

  if (Object.prototype.hasOwnProperty.call(payloadPersistencia, 'qtd_planejada')) {
    const qtdPlanejadaAtualizada = normalizarQuantidadeInteiraProducao(payloadPersistencia.qtd_planejada);
    let qtdProduzidaAcumuladaAtual = normalizarQuantidadeInteiraProducao(atual.qtd_produzida_acumulada);
    if (qtdProduzidaAcumuladaAtual <= 0 && String(atual.estoque_atualizado).toLowerCase() === 'true') {
      qtdProduzidaAcumuladaAtual = normalizarQuantidadeInteiraProducao(atual.qtd_planejada);
    }
    if (qtdProduzidaAcumuladaAtual < 0) qtdProduzidaAcumuladaAtual = 0;

    const qtdRestanteAtualizada = Math.max(0, qtdPlanejadaAtualizada - qtdProduzidaAcumuladaAtual);
    payloadPersistencia.qtd_restante = qtdRestanteAtualizada;
    payloadPersistencia.estoque_atualizado = qtdRestanteAtualizada <= 0;
  }

  let materiaisRegerados = null;
  let saidasRegeradas = null;
  if (mudouReceita || mudouProduto || mudouQtd) {
    materiaisRegerados = {
      detalhado: gerarNecessidadesEntradaProducao(
        dadosAtualizados.produto_id,
        dadosAtualizados.receita_id,
        normalizarQuantidadeInteiraProducao(dadosAtualizados.qtd_planejada)
      )
    };
    saidasRegeradas = {
      detalhado: (explodirSaidasReceitaDetalhada(
        dadosAtualizados.produto_id,
        dadosAtualizados.receita_id,
        normalizarQuantidadeInteiraProducao(dadosAtualizados.qtd_planejada)
      )?.itens) || []
    };
  }

  const ok = updateById(
    ABA_PRODUCAO,
    'producao_id',
    id,
    payloadPersistencia,
    PRODUCAO_SCHEMA
  );

  if (ok && materiaisRegerados) {
    limparReservasEntradaProducao(id, { preservarConsumidas: false });
    salvarNecessidadesEntradaProducao(
      id,
      dadosAtualizados.receita_id,
      materiaisRegerados.detalhado,
      { preservarManuais: true }
    );
    if (saidasRegeradas) {
      salvarDestinosProducao(
        id,
        dadosAtualizados.receita_id,
        saidasRegeradas.detalhado,
        { preservarManuais: true }
      );
    }
    recalcularMateriaisPrevistosFromVinculos(id);
  }

  return ok;
}

function deletarProducao(id) {
  assertCanWriteProducao('Exclusao de producao');
  const ok = updateById(
    ABA_PRODUCAO,
    'producao_id',
    id,
    { ativo: false },
    PRODUCAO_SCHEMA
  );
  if (ok) {
    limparNecessidadesEntradaProducao(id, { preservarManuais: false });
    limparReservasEntradaProducao(id, { preservarConsumidas: false });
  }
  return ok;
}

function listarEtapasProducao(producaoId) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_ETAPAS);
  if (!sheet) return [];

  const rows = rowsToObjects(sheet);
  return rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.producao_id === producaoId)
    .map(i => ({
      ...i,
      feito: String(i.feito).toLowerCase() === 'true',
      data_feito: formatDateSafe(i.data_feito, 'yyyy-MM-dd')
    }))
    .sort((a, b) => parseNumeroBR(a.ordem) - parseNumeroBR(b.ordem));
}

function listarMateriaisExtrasProducao(producaoId) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_MATERIAIS);
  if (!sheet) return [];

  const rows = rowsToObjects(sheet);
  return rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.producao_id === producaoId)
    .map(i => ({
      estoque_id: i.estoque_id,
      quantidade: parseNumeroBR(i.quantidade),
      unidade: i.unidade || ''
    }));
}

function listarMateriaisPrevistosSnapshot(producaoId) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_MATERIAIS_PREVISTOS);
  if (!sheet) return [];

  const rows = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.producao_id === producaoId);

  if (!rows || rows.length === 0) return [];

  const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
  const estoqueRows = sheetEstoque ? rowsToObjects(sheetEstoque) : [];
  const estoqueMap = {};
  estoqueRows.forEach(i => {
    estoqueMap[i.ID] = i;
  });

  return rows.map(r => {
    const estoqueItem = estoqueMap[r.estoque_id] || {};
    return {
      estoque_id: r.estoque_id,
      item: r.item_snapshot || estoqueItem.item || r.estoque_id,
      unidade: r.unidade || estoqueItem.unidade || '',
      quantidade: parseNumeroBR(r.quantidade),
      valor_unit: parseNumeroBR(r.valor_unit_snapshot || estoqueItem.valor_unit)
    };
  });
}

function inativarLinhasEmBlocosProducao(sheet, ativoCol, linhasAlvo) {
  if (!sheet || ativoCol < 0) return 0;

  const linhas = Array.isArray(linhasAlvo)
    ? linhasAlvo.filter(l => Number.isFinite(l) && l >= 2).sort((a, b) => a - b)
    : [];
  if (linhas.length === 0) return 0;

  let totalInativado = 0;
  let blocoInicio = linhas[0];
  let blocoTamanho = 1;

  function flushBloco() {
    const valores = Array.from({ length: blocoTamanho }, () => [false]);
    sheet.getRange(blocoInicio, ativoCol + 1, blocoTamanho, 1).setValues(valores);
    totalInativado += blocoTamanho;
  }

  for (let i = 1; i < linhas.length; i++) {
    const linhaAtual = linhas[i];
    const linhaAnterior = linhas[i - 1];
    if (linhaAtual === linhaAnterior + 1) {
      blocoTamanho += 1;
      continue;
    }

    flushBloco();
    blocoInicio = linhaAtual;
    blocoTamanho = 1;
  }

  flushBloco();
  return totalInativado;
}

function limparMateriaisPrevistosSnapshot(producaoId) {
  assertCanWriteProducao('Limpeza de snapshot de materiais previstos');
  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(ABA_PRODUCAO_MATERIAIS_PREVISTOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUCAO_MATERIAIS_PREVISTOS);
  }

  ensureSchema(sheet, PRODUCAO_MATERIAIS_PREVISTOS_SCHEMA);

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idCol = headers.indexOf('producao_id');
  const ativoCol = headers.indexOf('ativo');

  if (idCol !== -1 && ativoCol !== -1 && data.length > 1) {
    const producaoIdAlvo = String(producaoId || '').trim();
    const linhasAlvo = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol] || '').trim() !== producaoIdAlvo) continue;
      if (String(data[i][ativoCol]).toLowerCase() === 'false') continue;
      linhasAlvo.push(i + 1);
    }
    inativarLinhasEmBlocosProducao(sheet, ativoCol, linhasAlvo);
  }
}

function salvarMateriaisPrevistosSnapshot(producaoId, itens) {
  assertCanWriteProducao('Salvamento de snapshot de materiais previstos');
  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(ABA_PRODUCAO_MATERIAIS_PREVISTOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUCAO_MATERIAIS_PREVISTOS);
  }

  ensureSchema(sheet, PRODUCAO_MATERIAIS_PREVISTOS_SCHEMA);
  limparMateriaisPrevistosSnapshot(producaoId);

  const linhasValidas = Array.isArray(itens) ? itens : [];
  linhasValidas.forEach(l => {
    if (!l.estoque_id) return;
    const quantidade = parseNumeroBR(l.quantidade);
    if (!quantidade || quantidade <= 0) return;

    const novo = {
      id: gerarId('PMP'),
      producao_id: producaoId,
      estoque_id: l.estoque_id,
      quantidade,
      unidade: l.unidade || '',
      item_snapshot: l.item || '',
      valor_unit_snapshot: parseNumeroBR(l.valor_unit),
      criado_em: new Date(),
      ativo: true
    };
    insert(ABA_PRODUCAO_MATERIAIS_PREVISTOS, novo, PRODUCAO_MATERIAIS_PREVISTOS_SCHEMA);
  });

  invalidarCachesRelacionadosAba(ABA_PRODUCAO);
  return true;
}

function agruparMateriaisPorEstoque(itens) {
  const map = {};
  let custoPrevisto = 0;

  (Array.isArray(itens) ? itens : []).forEach(i => {
    if (!i || !i.estoque_id) return;

    const estoqueId = String(i.estoque_id);
    const quantidade = parseNumeroBR(i.quantidade);
    if (!quantidade || quantidade <= 0) return;

    const valorUnit = parseNumeroBR(i.valor_unit);

    if (!map[estoqueId]) {
      map[estoqueId] = {
        estoque_id: estoqueId,
        item: i.item || estoqueId,
        unidade: i.unidade || '',
        quantidade: 0,
        valor_unit: valorUnit
      };
    }

    map[estoqueId].quantidade += quantidade;
    custoPrevisto += quantidade * valorUnit;
  });

  return {
    itens: Object.values(map),
    custoPrevisto
  };
}

function limparVinculosMateriaisProducao(producaoId, opcoes) {
  const cfg = opcoes || {};
  limparNecessidadesEntradaProducao(producaoId, {
    preservarManuais: !!cfg.preservarManuais
  });
  limparReservasEntradaProducao(producaoId, {
    preservarConsumidas: !!cfg.preservarManuais
  });
}

function salvarVinculosMateriaisProducao(producaoId, receitaId, vinculos, opcoes) {
  const cfg = opcoes || {};
  return salvarNecessidadesEntradaProducao(
    producaoId,
    receitaId,
    Array.isArray(vinculos) ? vinculos.map(v => ({
      id: String(v.id || gerarId('PNE')).trim(),
      producao_id: producaoId,
      receita_id: receitaId || v.receita_id || '',
      receita_entrada_id: v.receita_entrada_id || '',
      tipo_item: v.tipo_item || '',
      categoria: v.categoria || '',
      nome_item: v.origem_item || v.item || '',
      unidade: v.unidade || '',
      modo_atendimento: getModoAtendimentoEntradaProducao(v.tipo_item),
      quantidade_prevista: parseNumeroBR(v.quantidade),
      quantidade_baixada: parseNumeroBR(v.quantidade_consumida),
      valor_unit_referencia: parseNumeroBR(v.valor_unit),
      comprimento_min_cm: parseNumeroBR(v.comprimento_min_cm),
      largura_min_cm: parseNumeroBR(v.largura_min_cm),
      espessura_min_cm: parseNumeroBR(v.espessura_min_cm),
      serie_item: String(v.serie_item || '').trim(),
      origem_manual: !String(v.receita_entrada_id || '').trim(),
      criado_em: new Date(),
      ativo: true
    })) : [],
    { preservarManuais: !!cfg.preservarManuais }
  );
}

function limparDestinosProducao(producaoId, opcoes) {
  assertCanWriteProducao('Limpeza de destinos da producao');
  const cfg = opcoes || {};
  const preservarManuais = !!cfg.preservarManuais;

  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(ABA_PRODUCAO_DESTINOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUCAO_DESTINOS);
  }

  ensureSchema(sheet, PRODUCAO_DESTINOS_SCHEMA);

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idxProducao = headers.indexOf('producao_id');
  const idxReceitaSaida = headers.indexOf('receita_saida_id');
  const idxAtivo = headers.indexOf('ativo');

  if (idxProducao === -1 || idxAtivo === -1 || data.length <= 1) return;

  const producaoIdAlvo = String(producaoId || '').trim();
  const linhasAlvo = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxProducao] || '').trim() !== producaoIdAlvo) continue;
    if (String(data[i][idxAtivo]).toLowerCase() === 'false') continue;

    if (preservarManuais && idxReceitaSaida !== -1) {
      const receitaSaidaId = String(data[i][idxReceitaSaida] || '').trim();
      if (!receitaSaidaId) continue;
    }

    linhasAlvo.push(i + 1);
  }

  inativarLinhasEmBlocosProducao(sheet, idxAtivo, linhasAlvo);
}

function salvarDestinosProducao(producaoId, receitaId, destinos, opcoes) {
  assertCanWriteProducao('Salvamento de destinos da producao');
  const cfg = opcoes || {};
  const preservarManuais = !!cfg.preservarManuais;

  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(ABA_PRODUCAO_DESTINOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUCAO_DESTINOS);
  }

  ensureSchema(sheet, PRODUCAO_DESTINOS_SCHEMA);
  limparDestinosProducao(producaoId, { preservarManuais });

  const linhas = Array.isArray(destinos) ? destinos : [];
  linhas.forEach(d => {
    const quantidadePrevista = parseNumeroBR(d.quantidade);
    if (!quantidadePrevista || quantidadePrevista <= 0) return;

    const estoqueId = String(d.estoque_id || '').trim();
    const quantidadeProduzidaBase = Object.prototype.hasOwnProperty.call(d, 'quantidade_produzida')
      ? parseNumeroBR(d.quantidade_produzida)
      : quantidadePrevista;
    const quantidadeProduzida = quantidadeProduzidaBase < 0 ? 0 : quantidadeProduzidaBase;
    const status = !estoqueId
      ? 'Sem vinculo'
      : (quantidadeProduzida > 0 ? 'Pendente' : 'Sem quantidade');

    const novo = {
      id: gerarId('PDS'),
      producao_id: producaoId,
      receita_id: receitaId || d.receita_id || '',
      receita_saida_id: d.receita_saida_id || '',
      estoque_id: estoqueId,
      nome_saida_snapshot: d.nome_saida || d.nome_item || '',
      unidade: d.unidade || '',
      quantidade_prevista: quantidadePrevista,
      quantidade_produzida: quantidadeProduzida,
      status,
      criado_em: new Date(),
      ativo: true
    };

    insert(ABA_PRODUCAO_DESTINOS, novo, PRODUCAO_DESTINOS_SCHEMA);
  });

  invalidarCachesRelacionadosAba(ABA_PRODUCAO);
  return true;
}

function listarDestinosProducao(producaoId) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_DESTINOS);
  if (!sheet) return [];

  const rows = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.producao_id === producaoId);

  if (rows.length === 0) return [];

  const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
  const estoqueRows = sheetEstoque ? rowsToObjects(sheetEstoque) : [];
  const estoqueMap = {};
  estoqueRows.forEach(i => {
    estoqueMap[i.ID] = i;
  });

  return rows.map(r => {
    const estoqueId = String(r.estoque_id || '').trim();
    const vinculado = !!estoqueId;
    const estoqueItem = estoqueMap[estoqueId] || {};

    const quantidadePrevista = parseNumeroBR(r.quantidade_prevista);
    const quantidadeProduzida = parseNumeroBR(r.quantidade_produzida);
    const saldoEstoque = parseNumeroBR(estoqueItem.quantidade);

    let status = String(r.status || '').trim();
    if (!vinculado) {
      status = 'Sem vinculo';
    } else if (!status || status === 'Sem vinculo') {
      status = quantidadeProduzida > 0 ? 'Pendente' : 'Sem quantidade';
    }

    return {
      id: r.id,
      producao_id: r.producao_id,
      receita_id: r.receita_id || '',
      receita_saida_id: r.receita_saida_id || '',
      estoque_id: estoqueId,
      nome_saida: r.nome_saida_snapshot || estoqueItem.item || '',
      unidade: r.unidade || estoqueItem.unidade || '',
      quantidade_prevista: quantidadePrevista,
      quantidade_produzida: quantidadeProduzida,
      saldo_estoque: saldoEstoque,
      vinculado,
      status
    };
  }).sort((a, b) => {
    const ta = String(a.nome_saida || '').toLowerCase();
    const tb = String(b.nome_saida || '').toLowerCase();
    return ta.localeCompare(tb);
  });
}

function vincularDestinoProducaoAoEstoque(producaoId, destinoId, estoqueId) {
  assertCanWriteProducao('Vinculacao de destino da producao ao estoque');
  if (!producaoId || !destinoId) {
    throw new Error('Destino invalido');
  }

  const estoqueIdNorm = String(estoqueId || '').trim();
  if (!estoqueIdNorm) {
    throw new Error('Selecione um item do estoque');
  }

  const sheetDestinos = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_DESTINOS);
  if (!sheetDestinos) {
    throw new Error('Aba de destinos nao encontrada');
  }

  const destino = rowsToObjects(sheetDestinos)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .find(i => i.id === destinoId && i.producao_id === producaoId);
  if (!destino) {
    throw new Error('Destino nao encontrado');
  }

  const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
  if (!sheetEstoque) {
    throw new Error('Aba ESTOQUE nao encontrada');
  }

  const estoque = rowsToObjects(sheetEstoque)
    .find(i => i.ID === estoqueIdNorm && String(i.ativo).toLowerCase() === 'true');
  if (!estoque) {
    throw new Error('Item de estoque nao encontrado');
  }

  const quantidadeProduzida = parseNumeroBR(destino.quantidade_produzida);
  const status = quantidadeProduzida > 0 ? 'Pendente' : 'Sem quantidade';

  updateById(
    ABA_PRODUCAO_DESTINOS,
    'id',
    destinoId,
    {
      estoque_id: estoqueIdNorm,
      nome_saida_snapshot: destino.nome_saida_snapshot || estoque.item || estoqueIdNorm,
      unidade: destino.unidade || estoque.unidade || '',
      status
    },
    PRODUCAO_DESTINOS_SCHEMA
  );

  return {
    destinos: listarDestinosProducao(producaoId)
  };
}

function atualizarDestinoProducao(producaoId, destinoId, payload) {
  assertCanWriteProducao('Atualizacao de destino da producao');
  if (!producaoId || !destinoId) {
    throw new Error('Destino invalido');
  }

  const dados = payload || {};
  const sheetDestinos = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_DESTINOS);
  if (!sheetDestinos) {
    throw new Error('Aba de destinos nao encontrada');
  }

  const destino = rowsToObjects(sheetDestinos)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .find(i => i.id === destinoId && i.producao_id === producaoId);
  if (!destino) {
    throw new Error('Destino nao encontrado');
  }

  let estoqueId = String(destino.estoque_id || '').trim();
  if (Object.prototype.hasOwnProperty.call(dados, 'estoque_id')) {
    estoqueId = String(dados.estoque_id || '').trim();
  }

  let unidade = String(dados.unidade || destino.unidade || '').trim();
  let nomeSaida = String(dados.nome_saida || destino.nome_saida_snapshot || '').trim();
  let quantidadeProduzida = Object.prototype.hasOwnProperty.call(dados, 'quantidade_produzida')
    ? parseNumeroBR(dados.quantidade_produzida)
    : parseNumeroBR(destino.quantidade_produzida);

  if (quantidadeProduzida < 0) {
    throw new Error('Quantidade produzida invalida');
  }

  if (estoqueId) {
    const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
    const estoque = sheetEstoque
      ? rowsToObjects(sheetEstoque).find(i => i.ID === estoqueId && String(i.ativo).toLowerCase() === 'true')
      : null;
    if (!estoque) {
      throw new Error('Item de estoque nao encontrado');
    }
    if (!unidade) unidade = estoque.unidade || '';
    if (!nomeSaida) nomeSaida = estoque.item || '';
  }

  const status = !estoqueId
    ? 'Sem vinculo'
    : (quantidadeProduzida > 0 ? 'Pendente' : 'Sem quantidade');

  updateById(
    ABA_PRODUCAO_DESTINOS,
    'id',
    destinoId,
    {
      estoque_id: estoqueId,
      nome_saida_snapshot: nomeSaida,
      unidade,
      quantidade_produzida: quantidadeProduzida,
      status
    },
    PRODUCAO_DESTINOS_SCHEMA
  );

  return {
    destinos: listarDestinosProducao(producaoId)
  };
}

function adicionarDestinoManualProducao(producaoId, payload) {
  assertCanWriteProducao('Adicao manual de destino da producao');
  if (!producaoId) {
    throw new Error('Producao invalida');
  }

  const dados = payload || {};
  const nomeSaida = String(dados.nome_saida || dados.nome_item || '').trim();
  const quantidadePrevista = parseNumeroBR(dados.quantidade_prevista || dados.quantidade || dados.quantidade_produzida);
  const estoqueId = String(dados.estoque_id || '').trim();

  if (!nomeSaida) {
    throw new Error('Informe o nome da saida');
  }
  if (!quantidadePrevista || quantidadePrevista <= 0) {
    throw new Error('Quantidade prevista invalida');
  }

  let quantidadeProduzida = Object.prototype.hasOwnProperty.call(dados, 'quantidade_produzida')
    ? parseNumeroBR(dados.quantidade_produzida)
    : quantidadePrevista;
  if (quantidadeProduzida < 0) {
    throw new Error('Quantidade produzida invalida');
  }

  let unidade = String(dados.unidade || '').trim();
  if (estoqueId) {
    const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
    const estoque = sheetEstoque
      ? rowsToObjects(sheetEstoque).find(i => i.ID === estoqueId && String(i.ativo).toLowerCase() === 'true')
      : null;
    if (!estoque) {
      throw new Error('Item de estoque nao encontrado');
    }
    if (!unidade) unidade = estoque.unidade || '';
  }

  const status = !estoqueId
    ? 'Sem vinculo'
    : (quantidadeProduzida > 0 ? 'Pendente' : 'Sem quantidade');

  const novo = {
    id: gerarId('PDS'),
    producao_id: producaoId,
    receita_id: '',
    receita_saida_id: '',
    estoque_id: estoqueId,
    nome_saida_snapshot: nomeSaida,
    unidade,
    quantidade_prevista: quantidadePrevista,
    quantidade_produzida: quantidadeProduzida,
    status,
    criado_em: new Date(),
    ativo: true
  };

  insert(ABA_PRODUCAO_DESTINOS, novo, PRODUCAO_DESTINOS_SCHEMA);

  return {
    item: novo,
    destinos: listarDestinosProducao(producaoId)
  };
}

function concluirDestinosProducao(producaoId) {
  assertCanWriteProducao('Conclusao de destinos da producao');
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_DESTINOS);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idxProducao = headers.indexOf('producao_id');
  const idxEstoque = headers.indexOf('estoque_id');
  const idxQtd = headers.indexOf('quantidade_produzida');
  const idxStatus = headers.indexOf('status');
  const idxAtivo = headers.indexOf('ativo');

  if (idxProducao === -1 || idxEstoque === -1 || idxQtd === -1 || idxStatus === -1) return;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idxProducao] !== producaoId) continue;
    if (idxAtivo !== -1 && String(row[idxAtivo]).toLowerCase() !== 'true') continue;

    const estoqueId = String(row[idxEstoque] || '').trim();
    const qtd = parseNumeroBR(row[idxQtd]);
    const status = !estoqueId
      ? 'Sem vinculo'
      : (qtd > 0 ? 'Concluido' : 'Sem quantidade');
    sheet.getRange(i + 1, idxStatus + 1).setValue(status);
  }
  invalidarCachesRelacionadosAba(ABA_PRODUCAO);
}

function recalcularMateriaisPrevistosFromVinculos(producaoId) {
  const vinculos = listarVinculosMateriaisProducao(producaoId);
  const itens = [];

  (Array.isArray(vinculos) ? vinculos : []).forEach(vinculo => {
    const reservas = Array.isArray(vinculo?.reservas) ? vinculo.reservas : [];
    reservas.forEach(reserva => {
      const quantidade = parseNumeroBR(reserva?.quantidade_restante);
      if (!reserva?.estoque_id || quantidade <= 0) return;
      itens.push({
        estoque_id: String(reserva.estoque_id || '').trim(),
        item: String(reserva.item || vinculo.item || reserva.estoque_id).trim(),
        unidade: String(reserva.unidade || vinculo.unidade || '').trim(),
        quantidade,
        valor_unit: parseNumeroBR(reserva.valor_unit)
      });
    });
  });

  const agrupado = agruparMateriaisPorEstoque(itens);
  salvarMateriaisPrevistosSnapshot(producaoId, agrupado.itens || []);

  return {
    itens: agrupado.itens || [],
    custoPrevisto: parseNumeroBR(agrupado.custoPrevisto)
  };
}

function listarVinculosMateriaisProducao(producaoId, opcoes) {
  const opts = opcoes || {};
  const contexto = opts.context || null;
  const prodId = String(producaoId || '').trim();
  if (!prodId) return [];

  const necessidades = listarNecessidadesEntradaRows_(prodId, { context: contexto });
  if (necessidades.length === 0) return [];

  const reservas = listarReservasEntradaRows_(prodId, { context: contexto });
  const reservasPorNecessidade = {};
  reservas.forEach(reserva => {
    const necessidadeId = String(reserva.necessidade_id || '').trim();
    if (!necessidadeId) return;
    if (!Array.isArray(reservasPorNecessidade[necessidadeId])) {
      reservasPorNecessidade[necessidadeId] = [];
    }
    reservasPorNecessidade[necessidadeId].push(reserva);
  });

  const estoqueLista = listarEstoqueAtivoComReservaProducao({ context: contexto });
  const estoqueMap = {};
  estoqueLista.forEach(item => {
    estoqueMap[String(item?.ID || '').trim()] = item;
  });

  return necessidades
    .map(necessidade => calcularResumoNecessidadeEntrada_(
      necessidade,
      reservasPorNecessidade[String(necessidade.id || '').trim()] || [],
      estoqueMap
    ))
    .sort((a, b) => {
      const serieA = parseNumeroBR(a?.serie_item);
      const serieB = parseNumeroBR(b?.serie_item);
      const ordemSerie = serieA - serieB;
      if (ordemSerie !== 0) return ordemSerie;
      const ta = `${a.tipo_item || ''} ${a.item || ''}`.toLowerCase();
      const tb = `${b.tipo_item || ''} ${b.item || ''}`.toLowerCase();
      return ta.localeCompare(tb);
    });
}

function salvarReservaEntradaProducao(producaoId, necessidadeId, payload) {
  assertCanWriteProducao('Salvamento de reserva da producao');
  const prodId = String(producaoId || '').trim();
  const necessidade = obterNecessidadeEntradaProducao_(prodId, necessidadeId);
  if (!prodId || !necessidade) {
    throw new Error('Necessidade invalida');
  }

  const dados = payload || {};
  const reservaId = String(dados.reserva_id || '').trim();
  const estoqueId = String(dados.estoque_id || '').trim();
  if (!estoqueId) {
    throw new Error('Selecione um item do estoque');
  }

  const reservaAtual = reservaId ? obterReservaEntradaProducao_(prodId, reservaId) : null;
  if (reservaId && !reservaAtual) {
    throw new Error('Reserva nao encontrada');
  }

  const estoqueLista = listarEstoqueAtivoComReservaProducao({
    excluirReservaId: reservaId
  });
  const estoque = estoqueLista.find(i => String(i?.ID || '').trim() === estoqueId);
  if (!estoque) {
    throw new Error('Item de estoque nao encontrado');
  }
  if (!itemEstoqueCompativelComNecessidade_(necessidade, estoque)) {
    throw new Error('Item de estoque nao compativel com a necessidade selecionada.');
  }

  const modo = String(necessidade.modo_atendimento || '').trim().toUpperCase();
  const quantidadeConsumidaAtual = parseNumeroBR(reservaAtual?.quantidade_consumida);
  const quantidadeDisponivel = parseNumeroBR(estoque.quantidade_disponivel);
  let quantidadeReservada = parseNumeroBR(dados.quantidade_reservada);
  let reservaExclusiva = false;

  const reservasNecessidade = listarReservasEntradaRows_(prodId)
    .filter(i => String(i.necessidade_id || '').trim() === String(necessidade.id || '').trim())
    .filter(i => String(i.id || '').trim() !== reservaId);

  if (modo === 'PECA_UNICA') {
    if (reservasNecessidade.length > 0) {
      throw new Error('Esta necessidade de madeira aceita apenas uma peca vinculada.');
    }
    quantidadeReservada = parseNumeroBR(estoque.quantidade);
    reservaExclusiva = true;
    if (quantidadeDisponivel <= 0) {
      throw new Error('A peca selecionada nao esta disponivel para reserva.');
    }
  } else {
    if (!quantidadeReservada || quantidadeReservada <= 0) {
      throw new Error('Quantidade reservada invalida.');
    }
    const limiteDisponivel = quantidadeDisponivel + Math.max(parseNumeroBR(reservaAtual?.quantidade_reservada) - quantidadeConsumidaAtual, 0);
    if (quantidadeReservada > limiteDisponivel + 0.000001) {
      throw new Error('Quantidade reservada maior que o saldo disponivel no estoque.');
    }
    const quantidadeMinima = quantidadeConsumidaAtual;
    if (quantidadeReservada + 0.000001 < quantidadeMinima) {
      throw new Error('Nao e permitido reservar menos do que ja foi baixado nesta reserva.');
    }
  }

  const payloadPersistencia = {
    producao_id: prodId,
    necessidade_id: String(necessidade.id || '').trim(),
    estoque_id: estoqueId,
    quantidade_reservada: quantidadeReservada,
    quantidade_consumida: quantidadeConsumidaAtual,
    reserva_exclusiva: reservaExclusiva,
    item_snapshot: String(estoque.item || necessidade.nome_item || estoqueId).trim(),
    unidade_snapshot: String(estoque.unidade || necessidade.unidade || '').trim(),
    valor_unit_snapshot: parseNumeroBR(estoque.custo_unitario || estoque.valor_unit),
    criado_em: reservaAtual?.criado_em || new Date(),
    ativo: true
  };

  if (reservaAtual) {
    updateById(
      ABA_PRODUCAO_RESERVAS_ENTRADA,
      'id',
      reservaId,
      payloadPersistencia,
      PRODUCAO_RESERVAS_ENTRADA_SCHEMA
    );
  } else {
    insert(ABA_PRODUCAO_RESERVAS_ENTRADA, {
      id: gerarId('PRE'),
      ...payloadPersistencia
    }, PRODUCAO_RESERVAS_ENTRADA_SCHEMA);
  }

  const materiais = recalcularMateriaisPrevistosFromVinculos(prodId);
  return {
    vinculos: listarVinculosMateriaisProducao(prodId),
    materiaisPrevistos: materiais.itens || [],
    custoPrevisto: parseNumeroBR(materiais.custoPrevisto)
  };
}

function removerReservaEntradaProducao(producaoId, necessidadeId, reservaId) {
  assertCanWriteProducao('Remocao de reserva da producao');
  const prodId = String(producaoId || '').trim();
  const necessidade = obterNecessidadeEntradaProducao_(prodId, necessidadeId);
  const reserva = obterReservaEntradaProducao_(prodId, reservaId);
  if (!prodId || !necessidade || !reserva) {
    throw new Error('Reserva invalida');
  }
  if (parseNumeroBR(reserva.quantidade_consumida) > 0) {
    throw new Error('Nao e permitido remover reserva que ja possui baixa registrada.');
  }

  updateById(
    ABA_PRODUCAO_RESERVAS_ENTRADA,
    'id',
    String(reserva.id || '').trim(),
    { ativo: false },
    PRODUCAO_RESERVAS_ENTRADA_SCHEMA
  );

  const materiais = recalcularMateriaisPrevistosFromVinculos(prodId);
  return {
    vinculos: listarVinculosMateriaisProducao(prodId),
    materiaisPrevistos: materiais.itens || [],
    custoPrevisto: parseNumeroBR(materiais.custoPrevisto)
  };
}

function vincularItemProducaoAoEstoque(producaoId, vinculoId, estoqueId, quantidadeReservada) {
  const prodId = String(producaoId || '').trim();
  const vinculoIdNorm = String(vinculoId || '').trim();
  const estadoAtual = listarVinculosMateriaisProducao(prodId)
    .find(i => String(i.id || '').trim() === vinculoIdNorm);
  if (!estadoAtual) {
    throw new Error('Necessidade nao encontrada');
  }

  const quantidade = String(estadoAtual.modo_atendimento || '').trim().toUpperCase() === 'PECA_UNICA'
    ? ''
    : (
      Object.prototype.hasOwnProperty.call(arguments, 3)
        ? quantidadeReservada
        : estadoAtual.quantidade_pendente
    );

  return salvarReservaEntradaProducao(prodId, vinculoIdNorm, {
    estoque_id: estoqueId,
    quantidade_reservada: quantidade
  });
}

function vincularPendenciasEntradaProducao(producaoId, vinculacoes) {
  assertCanWriteProducao('Vinculacao de pendencias da producao');
  const prodId = String(producaoId || '').trim();
  const itens = Array.isArray(vinculacoes) ? vinculacoes : [];
  if (!prodId || itens.length === 0) {
    throw new Error('Nenhum vinculo informado');
  }

  itens.forEach(v => {
    const necessidadeId = String(v?.necessidade_id || v?.vinculo_id || '').trim();
    const estoqueId = String(v?.estoque_id || '').trim();
    if (!necessidadeId || !estoqueId) {
      throw new Error('Vinculo invalido');
    }
    salvarReservaEntradaProducao(prodId, necessidadeId, {
      estoque_id: estoqueId,
      quantidade_reservada: v?.quantidade_reservada
    });
  });

  const materiais = recalcularMateriaisPrevistosFromVinculos(prodId);
  return {
    vinculos: listarVinculosMateriaisProducao(prodId),
    materiaisPrevistos: materiais.itens || [],
    custoPrevisto: parseNumeroBR(materiais.custoPrevisto)
  };
}

function adicionarItemManualProducao(producaoId, payload) {
  assertCanWriteProducao('Adicao manual de item na producao');
  if (!producaoId) {
    throw new Error('Producao invalida');
  }

  const dados = payload || {};
  const nomeItem = String(dados.nome_item || dados.nome || '').trim();
  const quantidadePrevista = parseNumeroBR(dados.quantidade_prevista || dados.quantidade);
  const tipoItem = String(dados.tipo_item || 'MANUAL').toUpperCase();
  const estoqueId = String(dados.estoque_id || '').trim();

  if (!nomeItem) {
    throw new Error('Informe o nome do item');
  }
  if (!quantidadePrevista || quantidadePrevista <= 0) {
    throw new Error('Quantidade invalida');
  }

  const unidade = String(dados.unidade || '').trim();
  const novo = {
    id: gerarId('PNE'),
    producao_id: producaoId,
    receita_id: '',
    receita_entrada_id: '',
    tipo_item: tipoItem,
    categoria: String(dados.categoria || '').trim(),
    nome_item: nomeItem,
    unidade,
    modo_atendimento: 'QUANTIDADE',
    quantidade_prevista: quantidadePrevista,
    quantidade_baixada: 0,
    valor_unit_referencia: parseNumeroBR(dados.valor_unit_snapshot),
    comprimento_min_cm: 0,
    largura_min_cm: 0,
    espessura_min_cm: 0,
    serie_item: '',
    origem_manual: true,
    criado_em: new Date(),
    ativo: true
  };

  insert(ABA_PRODUCAO_NECESSIDADES_ENTRADA, novo, PRODUCAO_NECESSIDADES_ENTRADA_SCHEMA);
  if (estoqueId) {
    salvarReservaEntradaProducao(producaoId, novo.id, {
      estoque_id: estoqueId,
      quantidade_reservada: quantidadePrevista
    });
  }

  const materiais = recalcularMateriaisPrevistosFromVinculos(producaoId);
  return {
    item: novo,
    vinculos: listarVinculosMateriaisProducao(producaoId),
    materiaisPrevistos: materiais.itens || [],
    custoPrevisto: parseNumeroBR(materiais.custoPrevisto)
  };
}

function regerarMateriaisEObjetosDeVinculo(producaoId, produtoId, receitaId, qtdPlanejada) {
  const itensDetalhados = gerarNecessidadesEntradaProducao(produtoId, receitaId, qtdPlanejada);

  salvarVinculosMateriaisProducao(
    producaoId,
    receitaId,
    itensDetalhados,
    { preservarManuais: true }
  );
  const agrupado = recalcularMateriaisPrevistosFromVinculos(producaoId);

  return {
    itensDetalhados,
    itensAgregados: agrupado.itens || [],
    custoPrevisto: parseNumeroBR(agrupado.custoPrevisto)
  };
}

function regenerarVinculosProducaoExistentes() {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO);
  if (!sheet) {
    return { atualizadas: 0, erros: [] };
  }

  const ordens = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.producao_id && i.produto_id && i.receita_id);

  let atualizadas = 0;
  const erros = [];

  ordens.forEach(o => {
    try {
      regerarMateriaisEObjetosDeVinculo(
        o.producao_id,
        o.produto_id,
        o.receita_id,
        normalizarQuantidadeInteiraProducao(o.qtd_planejada)
      );
      atualizadas++;
    } catch (err) {
      erros.push({
        producao_id: o.producao_id,
        erro: err && err.message ? err.message : 'Erro ao regerar vinculos'
      });
    }
  });

  return { atualizadas, erros };
}

function gerarMateriaisPrevistosReceita(produtoId, receitaId, qtdPlanejada) {
  if (!produtoId || !receitaId) {
    throw new Error('Modelo nao informado');
  }

  const resp = explodirReceita(produtoId, receitaId, qtdPlanejada);
  return resp || { itens: [], custoPrevisto: 0 };
}

function adicionarMaterialExtraProducao(producaoId, estoqueId, quantidade) {
  assertCanWriteProducao('Adicao de material extra na producao');
  if (!producaoId || !estoqueId) {
    throw new Error('Producao ou estoque invalido');
  }

  const qtd = parseNumeroBR(quantidade);
  if (!qtd || qtd <= 0) {
    throw new Error('Quantidade invalida');
  }

  const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
  if (!sheetEstoque) {
    throw new Error('Aba ESTOQUE nao encontrada');
  }

  const estoqueRows = rowsToObjects(sheetEstoque);
  const estoqueItem = estoqueRows.find(i => i.ID === estoqueId) || {};
  const unidade = estoqueItem.unidade || '';

  const novo = {
    id: gerarId('PM'),
    producao_id: producaoId,
    estoque_id: estoqueId,
    quantidade: qtd,
    unidade,
    criado_em: new Date(),
    ativo: true
  };

  insert(ABA_PRODUCAO_MATERIAIS, novo, PRODUCAO_MATERIAIS_SCHEMA);

  return {
    ...novo,
    criado_em: formatDateSafe(novo.criado_em, 'yyyy-MM-dd HH:mm')
  };
}

function aplicarBaixaMateriaisExtras(producaoId, itensValidos) {
  assertCanWriteProducao('Baixa de materiais extras da producao');
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_MATERIAIS);
  if (!sheet) return;

  const rows = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.producao_id === producaoId);

  const extrasPorEstoque = {};
  rows.forEach(r => {
    if (!extrasPorEstoque[r.estoque_id]) {
      extrasPorEstoque[r.estoque_id] = [];
    }
    extrasPorEstoque[r.estoque_id].push(r);
  });

  (Array.isArray(itensValidos) ? itensValidos : []).forEach(item => {
    let restante = parseNumeroBR(item.quantidade);
    if (restante <= 0) return;

    const lista = extrasPorEstoque[item.estoque_id] || [];
    for (let i = 0; i < lista.length && restante > 0; i++) {
      const extra = lista[i];
      const qtdAtual = parseNumeroBR(extra.quantidade);
      if (!extra.id || qtdAtual <= 0) continue;

      if (qtdAtual <= restante) {
        restante -= qtdAtual;
        updateById(
          ABA_PRODUCAO_MATERIAIS,
          'id',
          extra.id,
          { quantidade: 0, ativo: false },
          PRODUCAO_MATERIAIS_SCHEMA
        );
        continue;
      }

      const novoQtd = qtdAtual - restante;
      restante = 0;
      updateById(
        ABA_PRODUCAO_MATERIAIS,
        'id',
        extra.id,
        { quantidade: novoQtd },
        PRODUCAO_MATERIAIS_SCHEMA
      );
    }
  });
}

function atualizarConsumoVinculosProducao(producaoId, itensConsumidos) {
  assertCanWriteProducao('Atualizacao de consumo dos vinculos da producao');
  const prodId = String(producaoId || '').trim();
  const itens = Array.isArray(itensConsumidos) ? itensConsumidos : [];
  if (!prodId || itens.length === 0) return;

  const reservas = listarReservasEntradaRows_(prodId);
  const reservasMap = {};
  reservas.forEach(reserva => {
    reservasMap[String(reserva.id || '').trim()] = reserva;
  });

  const necessidades = listarNecessidadesEntradaRows_(prodId);
  const necessidadesMap = {};
  necessidades.forEach(necessidade => {
    necessidadesMap[String(necessidade.id || '').trim()] = necessidade;
  });

  const consumoPorNecessidade = {};

  itens.forEach(item => {
    const reservaId = String(item?.reserva_id || '').trim();
    const necessidadeId = String(item?.necessidade_id || '').trim();
    const quantidade = parseNumeroBR(item?.quantidade);
    if (quantidade <= 0) return;

    const reserva = reservaId ? reservasMap[reservaId] : null;
    const necessidade = necessidadesMap[necessidadeId] || (reserva ? necessidadesMap[String(reserva.necessidade_id || '').trim()] : null);
    if (!necessidade) return;

    if (reserva) {
      const consumidaAtual = parseNumeroBR(reserva.quantidade_consumida);
      updateById(
        ABA_PRODUCAO_RESERVAS_ENTRADA,
        'id',
        String(reserva.id || '').trim(),
        {
          quantidade_consumida: consumidaAtual + quantidade
        },
        PRODUCAO_RESERVAS_ENTRADA_SCHEMA
      );
      reserva.quantidade_consumida = consumidaAtual + quantidade;
    }

    const necId = String(necessidade.id || '').trim();
    consumoPorNecessidade[necId] = (consumoPorNecessidade[necId] || 0) + quantidade;
  });

  Object.keys(consumoPorNecessidade).forEach(necessidadeId => {
    const necessidade = necessidadesMap[necessidadeId];
    if (!necessidade) return;

    const modo = String(necessidade.modo_atendimento || '').trim().toUpperCase();
    const baixadaAtual = parseNumeroBR(necessidade.quantidade_baixada);
    const prevista = parseNumeroBR(necessidade.quantidade_prevista);
    const acrescimo = parseNumeroBR(consumoPorNecessidade[necessidadeId]);
    const novaBaixada = modo === 'PECA_UNICA'
      ? prevista
      : Math.min(prevista, baixadaAtual + acrescimo);

    updateById(
      ABA_PRODUCAO_NECESSIDADES_ENTRADA,
      'id',
      necessidadeId,
      {
        quantidade_baixada: novaBaixada
      },
      PRODUCAO_NECESSIDADES_ENTRADA_SCHEMA
    );
  });

  invalidarCachesRelacionadosAba(ABA_PRODUCAO);
}

function atualizarEtapasProducao(producaoId, etapas) {
  assertCanWriteProducao('Atualizacao de etapas da producao');
  if (!Array.isArray(etapas)) return false;

  etapas.forEach(et => {
    const dataFeito = et.data_feito
      ? new Date(et.data_feito)
      : '';
    const dataFeitoValida = dataFeito instanceof Date && !isNaN(dataFeito)
      ? dataFeito
      : '';

    if (et.id) {
      updateById(
        ABA_PRODUCAO_ETAPAS,
        'id',
        et.id,
        {
          nome_etapa: et.nome_etapa,
          feito: et.feito,
          ordem: et.ordem,
          data_feito: dataFeitoValida
        },
        PRODUCAO_ETAPAS_SCHEMA
      );
      return;
    }

    const novo = {
      id: gerarId('EPR'),
      producao_id: producaoId,
      nome_etapa: et.nome_etapa || '',
      feito: !!et.feito,
      ordem: et.ordem || '',
      data_feito: dataFeitoValida,
      ativo: true
    };
    insert(ABA_PRODUCAO_ETAPAS, novo, PRODUCAO_ETAPAS_SCHEMA);
  });

  return true;
}

function deletarEtapaProducao(id) {
  assertCanWriteProducao('Exclusao de etapa da producao');
  return updateById(
    ABA_PRODUCAO_ETAPAS,
    'id',
    id,
    { ativo: false },
    PRODUCAO_ETAPAS_SCHEMA
  );
}

function obterMateriaisPrevistosProducao(producaoId) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO);
  if (!sheet) return { itens: [], custoPrevisto: 0 };

  const rows = rowsToObjects(sheet);
  const ordem = rows.find(i => i.producao_id === producaoId);
  if (!ordem) return { itens: [], custoPrevisto: 0 };
  if (!ordem.receita_id) return { itens: [], custoPrevisto: 0 };

  let baseItens = listarMateriaisPrevistosSnapshot(producaoId);
  let custoPrevisto = 0;

  if (!baseItens || baseItens.length === 0) {
    const vinculos = listarVinculosMateriaisProducao(producaoId);

    if (Array.isArray(vinculos) && vinculos.length > 0) {
      const recalculado = recalcularMateriaisPrevistosFromVinculos(producaoId);
      baseItens = recalculado.itens || [];
      custoPrevisto = parseNumeroBR(recalculado.custoPrevisto);
    } else {
      const base = regerarMateriaisEObjetosDeVinculo(
        producaoId,
        ordem.produto_id,
        ordem.receita_id,
        normalizarQuantidadeInteiraProducao(ordem.qtd_planejada)
      );
      baseItens = base.itensAgregados || [];
      custoPrevisto = parseNumeroBR(base.custoPrevisto);
    }
  } else {
    custoPrevisto = baseItens.reduce((acc, i) => {
      return acc + (parseNumeroBR(i.quantidade) * parseNumeroBR(i.valor_unit));
    }, 0);
  }

  const extras = listarMateriaisExtrasProducao(producaoId);

  if (!extras || extras.length === 0) {
    return {
      itens: baseItens,
      custoPrevisto
    };
  }

  const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
  const estoqueRows = sheetEstoque ? rowsToObjects(sheetEstoque) : [];
  const estoqueMap = {};
  estoqueRows.forEach(i => {
    estoqueMap[i.ID] = i;
  });

  const itensMap = {};
  (baseItens || []).forEach(i => {
    itensMap[i.estoque_id] = { ...i };
  });

  extras.forEach(e => {
    const estoqueItem = estoqueMap[e.estoque_id] || {};
    const valorUnit = parseNumeroBR(estoqueItem.valor_unit);
    const unidade = estoqueItem.unidade || e.unidade || '';
    const nomeItem = estoqueItem.item || e.estoque_id || '';
    const qtd = parseNumeroBR(e.quantidade);

    if (!itensMap[e.estoque_id]) {
      itensMap[e.estoque_id] = {
        estoque_id: e.estoque_id,
        item: nomeItem,
        unidade,
        quantidade: 0,
        valor_unit: valorUnit
      };
    }

    itensMap[e.estoque_id].quantidade += qtd;
    custoPrevisto += qtd * valorUnit;
  });

  return {
    itens: Object.values(itensMap),
    custoPrevisto
  };
}

function normalizarTextoProducao(valor) {
  return String(valor || '').trim().toLowerCase();
}

function normalizarTextoSemAcentoProducao(valor) {
  const bruto = String(valor || '').trim();
  if (!bruto) return '';
  const semAcento = (typeof bruto.normalize === 'function')
    ? bruto.normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    : bruto;
  return semAcento
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function listarPendenciasVinculoEntradaProducao(producaoId, estoqueAtivoMap) {
  const mapaAtivos = estoqueAtivoMap || {};
  const vinculos = listarVinculosMateriaisProducao(producaoId);

  return (Array.isArray(vinculos) ? vinculos : [])
    .filter(v => {
      const reservas = Array.isArray(v.reservas) ? v.reservas : [];
      if (reservas.length === 0) return true;
      return reservas.some(reserva => {
        const estoqueId = String(reserva.estoque_id || '').trim();
        return !estoqueId || !mapaAtivos[estoqueId];
      });
    })
    .map(v => ({
      id: v.id,
      producao_id: v.producao_id,
      receita_entrada_id: v.receita_entrada_id || '',
      necessidade_id: v.id,
      estoque_id: String(v.reservas?.[0]?.estoque_id || '').trim(),
      tipo_item: v.tipo_item || '',
      categoria: v.categoria || '',
      origem_item: v.origem_item || '',
      item: v.item || v.origem_item || v.id,
      unidade: v.unidade || '',
      quantidade_prevista: parseNumeroBR(v.quantidade_prevista)
    }));
}

function agruparSaidasReceitaParaEstoque(saidas) {
  const map = {};

  (Array.isArray(saidas) ? saidas : []).forEach(s => {
    const nomeSaida = String(s.nome_saida || '').trim();
    const produtoRefId = String(s.produto_ref_id || '').trim();
    const tipoItem = normalizarTipoSaidaProducao(s.tipo_item);
    const categoria = normalizarCategoriaSaidaProducao(tipoItem, s.categoria);
    const unidade = String(s.unidade || '').trim();
    const quantidade = parseNumeroBR(s.quantidade);

    if (!nomeSaida || quantidade <= 0) return;

    const chave = `${normalizarTextoProducao(tipoItem)}||${normalizarTextoProducao(categoria)}||${normalizarTextoProducao(nomeSaida)}||${normalizarTextoProducao(unidade)}||${normalizarTextoProducao(produtoRefId)}`;
    if (!map[chave]) {
      map[chave] = {
        nome_saida: nomeSaida,
        produto_ref_id: produtoRefId,
        tipo_item: tipoItem,
        categoria,
        unidade,
        quantidade: 0
      };
    }

    map[chave].quantidade += quantidade;
    if (!map[chave].categoria && categoria) {
      map[chave].categoria = categoria;
    }
    if (!map[chave].unidade && unidade) {
      map[chave].unidade = unidade;
    }
  });

  return Object.values(map).map(i => ({
    nome_saida: i.nome_saida,
    produto_ref_id: i.produto_ref_id,
    tipo_item: i.tipo_item,
    categoria: i.categoria,
    unidade: i.unidade,
    quantidade: parseNumeroBR(i.quantidade)
  }));
}

function normalizarTipoSaidaProducao(tipo) {
  const t = String(tipo || '').trim().toUpperCase();
  return t || 'PRODUTO';
}

function normalizarCategoriaSaidaProducao(tipo, categoria) {
  const tipoNorm = normalizarTipoSaidaProducao(tipo);
  const categoriaInformada = String(categoria || '').trim();

  let validacoes = null;
  try {
    validacoes = typeof obterValidacoes === 'function' ? obterValidacoes() : null;
  } catch (error) {
    validacoes = null;
  }

  const mapaCategorias = validacoes?.categoriasPorTipo || {};
  const categoriasTipo = Array.isArray(mapaCategorias[tipoNorm]) ? mapaCategorias[tipoNorm] : [];

  if (categoriaInformada) {
    if (categoriasTipo.length > 0) {
      const encontrada = categoriasTipo.find(v =>
        String(v || '').trim().toUpperCase() === categoriaInformada.toUpperCase()
      );
      return encontrada || '';
    }

    const categoriasGerais = Array.isArray(validacoes?.categorias) ? validacoes.categorias : [];
    if (categoriasGerais.length > 0) {
      const encontrada = categoriasGerais.find(v =>
        String(v || '').trim().toUpperCase() === categoriaInformada.toUpperCase()
      );
      return encontrada || '';
    }

    return categoriaInformada;
  }

  if (categoriasTipo.length > 0) {
    if (tipoNorm === 'PRODUTO') {
      const peca = categoriasTipo.find(v => String(v || '').trim().toUpperCase() === 'PECA');
      if (peca) return peca;
    }
    return String(categoriasTipo[0] || '').trim();
  }

  return tipoNorm === 'PRODUTO' ? 'PECA' : '';
}

function localizarItemEstoqueProdutoPorSaida(estoqueRows, nomeSaida, unidadeSaida, tipoSaida, categoriaSaida) {
  const tipoNorm = normalizarTextoProducao(normalizarTipoSaidaProducao(tipoSaida));
  const categoriaNorm = normalizarTextoProducao(normalizarCategoriaSaidaProducao(tipoSaida, categoriaSaida));
  const nomeNorm = normalizarTextoProducao(nomeSaida);
  const unidadeNorm = normalizarTextoProducao(unidadeSaida);
  const lista = Array.isArray(estoqueRows) ? estoqueRows : [];

  let candidatos = lista.filter(i =>
    String(i.ativo).toLowerCase() === 'true' &&
    normalizarTextoProducao(i.tipo) === tipoNorm &&
    normalizarTextoProducao(i.item) === nomeNorm
  );

  if (categoriaNorm) {
    candidatos = candidatos.filter(i => normalizarTextoProducao(i.categoria) === categoriaNorm);
  }

  if (candidatos.length === 0) return null;
  if (!unidadeNorm) return candidatos[0];

  const exato = candidatos.find(i => normalizarTextoProducao(i.unidade) === unidadeNorm);
  return exato || null;
}

function obterCustoUnitarioEstoqueItem(item) {
  const tipo = String(item?.tipo || '').trim().toUpperCase();
  const custoBruto = String(item?.custo_unitario ?? '').trim();
  const custo = parseNumeroBR(custoBruto);

  if (tipo === 'PRODUTO') {
    if (custoBruto === '') return 0;
    return custo;
  }

  if (custoBruto !== '') return custo;
  return parseNumeroBR(item?.valor_unit);
}

function calcularCustoUnitarioSaidas(totalConsumo, saidasAgrupadas) {
  const totalConsumoSeguro = parseNumeroBR(totalConsumo);
  const totalSaidas = (Array.isArray(saidasAgrupadas) ? saidasAgrupadas : [])
    .reduce((acc, s) => acc + parseNumeroBR(s?.quantidade), 0);

  if (totalSaidas <= 0 || totalConsumoSeguro <= 0) {
    return 0;
  }
  return Number((totalConsumoSeguro / totalSaidas).toFixed(6));
}

function lancarSaidasProducaoNoEstoque(producaoId, saidasAgrupadas, estoqueRows, estoqueAtivoMap, custoUnitarioSaida) {
  const saidas = Array.isArray(saidasAgrupadas) ? saidasAgrupadas : [];
  const linhasEstoque = Array.isArray(estoqueRows) ? estoqueRows : [];
  const mapaAtivos = estoqueAtivoMap || {};
  const atualizados = [];
  const lotesCriados = [];
  const custoSaida = parseNumeroBR(custoUnitarioSaida);

  saidas.forEach(s => {
    const nomeSaida = String(s.nome_saida || '').trim();
    const produtoRefId = String(s.produto_ref_id || '').trim();
    const tipoSaida = normalizarTipoSaidaProducao(s.tipo_item);
    const categoriaSaida = normalizarCategoriaSaidaProducao(tipoSaida, s.categoria);
    const quantidade = parseNumeroBR(s.quantidade);
    const unidade = String(s.unidade || 'UN').trim() || 'UN';

    if (!nomeSaida || quantidade <= 0) return;

    const itemExistente = localizarItemEstoqueProdutoPorSaida(
      linhasEstoque,
      nomeSaida,
      unidade,
      tipoSaida,
      categoriaSaida
    );
    let estoqueId = '';
    let novoSaldo = quantidade;
    let custoUnitarioFinal = custoSaida;

    if (!itemExistente) {
      const novoItem = {
        ID: gerarId('EST'),
        tipo: tipoSaida,
        item: nomeSaida,
        unidade,
        valor_unit: custoSaida,
        custo_unitario: custoSaida,
        preco_venda: '',
        ativo: true,
        criado_em: new Date(),
        quantidade,
        comprimento_cm: '',
        largura_cm: '',
        espessura_cm: '',
        categoria: categoriaSaida,
        fornecedor: '',
        pago_por: '',
        potencia: '',
        voltagem: '',
        comprado_em: '',
        data_pagamento: '',
        forma_pagamento: '',
        parcelas: 1,
        vida_util_mes: '',
        observacao: `Gerado na OP ${producaoId}`,
        origem_tipo: 'PRODUCAO',
        origem_id: producaoId,
        op_id: producaoId
      };

      insert(ABA_ESTOQUE, novoItem, ESTOQUE_SCHEMA);
      linhasEstoque.push(novoItem);
      mapaAtivos[novoItem.ID] = novoItem;
      estoqueId = novoItem.ID;

      atualizados.push({
        ID: novoItem.ID,
        quantidade: parseNumeroBR(novoItem.quantidade)
      });
    } else {
      const saldoAtual = parseNumeroBR(itemExistente.quantidade);
      novoSaldo = saldoAtual + quantidade;
      const custoAtual = obterCustoUnitarioEstoqueItem(itemExistente);
      custoUnitarioFinal = novoSaldo > 0
        ? Number((((saldoAtual * custoAtual) + (quantidade * custoSaida)) / novoSaldo).toFixed(6))
        : 0;

      updateById(
        ABA_ESTOQUE,
        'ID',
        itemExistente.ID,
        {
          quantidade: novoSaldo,
          custo_unitario: custoUnitarioFinal,
          valor_unit: custoUnitarioFinal,
          categoria: itemExistente.categoria || categoriaSaida,
          origem_tipo: 'PRODUCAO',
          origem_id: producaoId,
          op_id: producaoId
        },
        ESTOQUE_SCHEMA
      );

      itemExistente.quantidade = novoSaldo;
      itemExistente.custo_unitario = custoUnitarioFinal;
      itemExistente.valor_unit = custoUnitarioFinal;
      if (!String(itemExistente.categoria || '').trim() && categoriaSaida) {
        itemExistente.categoria = categoriaSaida;
      }
      mapaAtivos[itemExistente.ID] = itemExistente;
      estoqueId = itemExistente.ID;

      atualizados.push({
        ID: itemExistente.ID,
        quantidade: novoSaldo
      });
    }

    const custoTotalLote = Number((quantidade * custoSaida).toFixed(6));
    const lote = {
      id: gerarId('LOT'),
      producao_id: producaoId,
      estoque_id: estoqueId,
      nome_saida: nomeSaida,
      produto_ref_id: produtoRefId,
      unidade,
      quantidade,
      custo_unitario: custoSaida,
      custo_total: custoTotalLote,
      criado_em: new Date(),
      ativo: true
    };
    insert(ABA_PRODUCAO_SAIDAS_LOTES, lote, PRODUCAO_SAIDAS_LOTES_SCHEMA);
    lotesCriados.push({
      ...lote,
      criado_em: formatDateSafe(lote.criado_em, 'yyyy-MM-dd HH:mm')
    });
  });

  return { atualizados, lotesCriados };
}

function atualizarQuantidadeEstoqueComSoftDeleteProducao_(estoqueId, novoSaldo) {
  const saldoFinal = Math.max(parseNumeroBR(novoSaldo), 0);
  const payload = {
    quantidade: saldoFinal
  };
  if (saldoFinal <= 0.000001) {
    payload.quantidade = 0;
    payload.ativo = false;
  }
  updateById(
    ABA_ESTOQUE,
    'ID',
    estoqueId,
    payload,
    ESTOQUE_SCHEMA
  );
  return payload;
}

function normalizarSaidasOverrideConsumo_(saidas, qtdProduzidaLote) {
  const qtdLote = normalizarQuantidadeInteiraProducao(qtdProduzidaLote);
  const lista = (Array.isArray(saidas) ? saidas : [])
    .map(s => ({
      receita_saida_id: String(s?.receita_saida_id || '').trim(),
      nome_saida: String(s?.nome_saida || '').trim(),
      produto_ref_id: String(s?.produto_ref_id || '').trim(),
      tipo_item: normalizarTipoSaidaProducao(s?.tipo_item),
      categoria: normalizarCategoriaSaidaProducao(s?.tipo_item, s?.categoria),
      unidade: String(s?.unidade || '').trim(),
      quantidade: parseNumeroBR(s?.quantidade)
    }))
    .filter(s => s.nome_saida && s.quantidade > 0);

  if (lista.length === 0 && qtdLote > 0) {
    return null;
  }

  return agruparSaidasReceitaParaEstoque(lista);
}

function consumirEstoque(producaoId, itensParaBaixar, opcoes) {
  assertCanWriteProducao('Consumo de estoque da producao');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (!producaoId) {
      throw new Error('Producao invalida');
    }

    const cfg = opcoes || {};
    const bypassEntradas = !!cfg.bypass_entradas;

    const sheetProducao = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO);
    if (!sheetProducao) {
      throw new Error('Aba PRODUCAO nao encontrada');
    }

    const producaoRows = rowsToObjects(sheetProducao);
    const ordem = producaoRows.find(i => i.producao_id === producaoId);
    if (!ordem) {
      throw new Error('Producao nao encontrada');
    }

    if (!ordem.receita_id) {
      throw new Error('Modelo nao informado');
    }

    const qtdPlanejada = normalizarQuantidadeInteiraProducao(ordem.qtd_planejada);
    if (qtdPlanejada <= 0) {
      throw new Error('Quantidade planejada invalida');
    }

    let qtdProduzidaAcumuladaAtual = normalizarQuantidadeInteiraProducao(ordem.qtd_produzida_acumulada);
    if (qtdProduzidaAcumuladaAtual <= 0 && String(ordem.estoque_atualizado).toLowerCase() === 'true') {
      qtdProduzidaAcumuladaAtual = qtdPlanejada;
    }
    if (qtdProduzidaAcumuladaAtual < 0) {
      qtdProduzidaAcumuladaAtual = 0;
    }

    const qtdRestanteAtual = Math.max(0, qtdPlanejada - qtdProduzidaAcumuladaAtual);
    if (qtdRestanteAtual <= 0) {
      throw new Error('Quantidade planejada ja foi totalmente produzida');
    }

    let qtdProduzidaLote = parseNumeroBR(cfg.qtd_produzida);
    if (!qtdProduzidaLote || qtdProduzidaLote <= 0) {
      qtdProduzidaLote = qtdRestanteAtual;
    }
    if (Math.abs(qtdProduzidaLote - Math.round(qtdProduzidaLote)) > 0.000001) {
      throw new Error('Quantidade do lote deve ser inteira.');
    }
    qtdProduzidaLote = normalizarQuantidadeInteiraProducao(qtdProduzidaLote);
    if (qtdProduzidaLote > qtdRestanteAtual) {
      throw new Error(`Quantidade do lote maior que o restante da OP (${qtdRestanteAtual}).`);
    }

    const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
    if (!sheetEstoque) {
      throw new Error('Aba ESTOQUE nao encontrada');
    }

    const estoqueRows = rowsToObjects(sheetEstoque);
    const estoqueMapAtivo = {};
    estoqueRows
      .filter(i => String(i.ativo).toLowerCase() === 'true')
      .forEach(i => {
        estoqueMapAtivo[String(i.ID || '').trim()] = i;
      });

    const vinculosMateriais = listarVinculosMateriaisProducao(producaoId);
    const vinculosMap = {};
    (Array.isArray(vinculosMateriais) ? vinculosMateriais : []).forEach(v => {
      vinculosMap[String(v.id || '').trim()] = v;
    });

    const itensEntrada = bypassEntradas ? [] : (Array.isArray(itensParaBaixar) ? itensParaBaixar : []);
    const itensValidos = [];

    itensEntrada.forEach(item => {
      const necessidadeId = String(item?.necessidade_id || '').trim();
      const estoqueId = String(item?.estoque_id || '').trim();
      const quantidade = parseNumeroBR(item?.quantidade);
      const reservaId = String(item?.reserva_id || '').trim();

      if (!necessidadeId || !estoqueId || quantidade <= 0) return;

      let vinculo = vinculosMap[necessidadeId];
      if (!vinculo) {
        throw new Error(`Necessidade de entrada nao encontrada: ${necessidadeId}`);
      }

      let reserva = (Array.isArray(vinculo.reservas) ? vinculo.reservas : [])
        .find(r => String(r.id || '').trim() === reservaId);
      if (!reserva) {
        salvarReservaEntradaProducao(producaoId, necessidadeId, {
          estoque_id: estoqueId,
          quantidade_reservada: quantidade
        });
        vinculo = (listarVinculosMateriaisProducao(producaoId) || [])
          .find(v => String(v.id || '').trim() === necessidadeId) || vinculo;
        reserva = (Array.isArray(vinculo.reservas) ? vinculo.reservas : [])
          .find(r => String(r.estoque_id || '').trim() === estoqueId);
      }

      if (!reserva) {
        throw new Error(`Reserva nao encontrada para a necessidade ${necessidadeId}`);
      }

      const estoqueItem = estoqueMapAtivo[estoqueId];
      if (!estoqueItem) {
        throw new Error(`Item de estoque nao encontrado: ${estoqueId}`);
      }

      const quantidadeReservadaRestante = Math.max(
        parseNumeroBR(reserva.quantidade_reservada) - parseNumeroBR(reserva.quantidade_consumida),
        0
      );
      if (quantidade > quantidadeReservadaRestante + 0.000001) {
        throw new Error(`Quantidade maior que a reserva disponivel para ${estoqueItem.item || estoqueId}`);
      }

      const saldoFisico = parseNumeroBR(estoqueItem.quantidade);
      if ((saldoFisico + 0.000001) < quantidade) {
        throw new Error(`Saldo insuficiente para ${estoqueItem.item || estoqueId}`);
      }

      itensValidos.push({
        necessidade_id: necessidadeId,
        reserva_id: String(reserva.id || '').trim(),
        estoque_id: estoqueId,
        quantidade: Number(quantidade.toFixed(6))
      });
    });

    if (!bypassEntradas && vinculosMateriais.length > 0 && itensValidos.length === 0) {
      throw new Error('Nenhum item de entrada informado para baixa.');
    }

    const consumoRegistrado = [];
    const estoqueAtualizados = [];

    if (!bypassEntradas) {
      const itensPorEstoque = {};
      itensValidos.forEach(i => {
        itensPorEstoque[i.estoque_id] = (itensPorEstoque[i.estoque_id] || 0) + parseNumeroBR(i.quantidade);
      });

      Object.keys(itensPorEstoque).forEach(estoqueId => {
        const quantidadeBaixada = parseNumeroBR(itensPorEstoque[estoqueId]);
        const estoqueItem = estoqueMapAtivo[estoqueId];
        const saldo = parseNumeroBR(estoqueItem.quantidade);
        const novoSaldo = Math.max(0, Number((saldo - quantidadeBaixada).toFixed(6)));
        const payloadEstoque = atualizarQuantidadeEstoqueComSoftDeleteProducao_(estoqueId, novoSaldo);

        estoqueMapAtivo[estoqueId].quantidade = novoSaldo;
        if (payloadEstoque.ativo === false) {
          estoqueMapAtivo[estoqueId].ativo = false;
        }

        estoqueAtualizados.push({
          ID: estoqueId,
          quantidade: novoSaldo
        });
      });

      itensValidos.forEach(i => {
        const estoqueItem = estoqueMapAtivo[i.estoque_id];

        const valorUnit = parseNumeroBR(estoqueItem.custo_unitario || estoqueItem.valor_unit);
        const total = Number((i.quantidade * valorUnit).toFixed(6));

        const consumo = {
          id: gerarId('PC'),
          producao_id: producaoId,
          estoque_id: i.estoque_id,
          quantidade_consumida: i.quantidade,
          valor_unit_snapshot: valorUnit,
          total_snapshot: total,
          criado_em: new Date(),
          ativo: true
        };

        insert(ABA_PRODUCAO_CONSUMO, consumo, PRODUCAO_CONSUMO_SCHEMA);
        consumoRegistrado.push(consumo);
      });

      aplicarBaixaMateriaisExtras(producaoId, itensValidos);
      atualizarConsumoVinculosProducao(producaoId, itensValidos);
    }

    const saidasOverride = normalizarSaidasOverrideConsumo_(cfg.saidas || cfg.saidas_override, qtdProduzidaLote);
    const saidasAgrupadas = saidasOverride || agruparSaidasReceitaParaEstoque(
      (explodirSaidasReceitaDetalhada(
        ordem.produto_id,
        ordem.receita_id,
        qtdProduzidaLote
      )?.itens) || []
    );
    if (saidasAgrupadas.length === 0) {
      throw new Error('Modelo sem saidas configuradas para lancamento no estoque');
    }
    const custoTotalEntradas = consumoRegistrado
      .reduce((acc, c) => acc + parseNumeroBR(c?.total_snapshot), 0);
    const custoUnitarioSaida = calcularCustoUnitarioSaidas(custoTotalEntradas, saidasAgrupadas);

    const resultadoSaidas = lancarSaidasProducaoNoEstoque(
      producaoId,
      saidasAgrupadas,
      estoqueRows,
      estoqueMapAtivo,
      custoUnitarioSaida
    );
    estoqueAtualizados.push(...(resultadoSaidas.atualizados || []));

    const qtdProduzidaAcumuladaNova = qtdProduzidaAcumuladaAtual + qtdProduzidaLote;
    const qtdRestanteNova = Math.max(0, qtdPlanejada - qtdProduzidaAcumuladaNova);
    const concluiuOP = qtdRestanteNova <= 0;
    const dataAtualizacao = new Date();
    const statusAtual = String(ordem.status || '').trim();
    const statusAtualNorm = normalizarTextoSemAcentoProducao(statusAtual);
    const statusNovo = concluiuOP
      ? 'Concluido'
      : (
        statusAtualNorm === 'EM PLANEJAMENTO' ||
        statusAtualNorm === 'CONCLUIDO' ||
        statusAtualNorm === 'CONCLUIDA' ||
        !statusAtual
          ? 'Em producao'
          : statusAtual
      );

    const payloadAtualizacaoOP = {
      qtd_produzida_acumulada: qtdProduzidaAcumuladaNova,
      qtd_restante: qtdRestanteNova,
      estoque_atualizado: concluiuOP,
      data_estoque_atualizado: concluiuOP ? dataAtualizacao : '',
      data_ultima_movimentacao_estoque: dataAtualizacao,
      status: statusNovo
    };
    if (concluiuOP) {
      payloadAtualizacaoOP.data_conclusao = ordem.data_conclusao || dataAtualizacao;
    }

    updateById(
      ABA_PRODUCAO,
      'producao_id',
      producaoId,
      payloadAtualizacaoOP,
      PRODUCAO_SCHEMA
    );

    const estoqueAtualizadosMap = {};
    estoqueAtualizados.forEach(i => {
      if (!i || !i.ID) return;
      estoqueAtualizadosMap[i.ID] = parseNumeroBR(i.quantidade);
    });
    const estoqueAtualizadosUnicos = Object.keys(estoqueAtualizadosMap).map(id => ({
      ID: id,
      quantidade: estoqueAtualizadosMap[id]
    }));

    return {
      estoqueAtualizados: estoqueAtualizadosUnicos,
      consumoRegistrado,
      lote: {
        qtd_produzida: qtdProduzidaLote,
        bypass_entradas: bypassEntradas
      },
      producaoAtualizada: {
        producao_id: producaoId,
        status: statusNovo,
        estoque_atualizado: concluiuOP,
        qtd_produzida_acumulada: qtdProduzidaAcumuladaNova,
        qtd_restante: qtdRestanteNova,
        data_conclusao: concluiuOP
          ? formatDateSafe(payloadAtualizacaoOP.data_conclusao, 'yyyy-MM-dd')
          : formatDateSafe(ordem.data_conclusao, 'yyyy-MM-dd'),
        data_estoque_atualizado: concluiuOP
          ? formatDateSafe(dataAtualizacao, 'yyyy-MM-dd')
          : '',
        data_ultima_movimentacao_estoque: formatDateSafe(dataAtualizacao, 'yyyy-MM-dd')
      },
      saidasLancadas: saidasAgrupadas,
      saidasLotes: resultadoSaidas.lotesCriados || []
    };
  } finally {
    lock.releaseLock();
  }
}

function listarSaidasLotesProducao(producaoId) {
  const prodId = String(producaoId || '').trim();
  if (!prodId) return [];

  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_SAIDAS_LOTES);
  if (!sheet) return [];
  const produtosIndices = listarProdutosAtivosIndicesProducao();
  const produtosPorId = produtosIndices.porId || {};
  const produtosPorNome = produtosIndices.porNome || {};

  return rowsToObjects(sheet)
    .filter(i =>
      String(i.ativo).toLowerCase() === 'true' &&
      String(i.producao_id || '').trim() === prodId
    )
    .map(i => {
      const produtoRefId = String(i.produto_ref_id || '').trim();
      const produtoResolvido = resolverProdutoSaidaProducao(
        produtoRefId,
        i.nome_saida,
        { porId: produtosPorId, porNome: produtosPorNome }
      );
      const produtoVinculado = produtoResolvido.produto;
      return {
        ...i,
        produto_ref_id: produtoRefId,
        quantidade: parseNumeroBR(i.quantidade),
        custo_unitario: parseNumeroBR(i.custo_unitario),
        custo_total: parseNumeroBR(i.custo_total),
        preco_venda_produto: produtoVinculado?.preco_venda || '',
        produto_id_vinculado: produtoResolvido.vinculoPorId ? (produtoVinculado?.produto_id || '') : '',
        vinculo_produto_por_id: produtoResolvido.vinculoPorId,
        permite_editar_preco_venda: produtoResolvido.vinculoPorId,
        criado_em: i.criado_em
          ? formatDateSafe(i.criado_em, 'yyyy-MM-dd HH:mm')
          : ''
      };
    })
    .sort((a, b) => {
      const da = new Date(a.criado_em || 0).getTime() || 0;
      const db = new Date(b.criado_em || 0).getTime() || 0;
      return db - da;
    });
}

function recalcularCustoMedioSaidasNoEstoque(estoqueId) {
  const estId = String(estoqueId || '').trim();
  if (!estId) return null;

  const sheetLotes = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_SAIDAS_LOTES);
  if (!sheetLotes) return null;

  const lotes = rowsToObjects(sheetLotes)
    .filter(i =>
      String(i.ativo).toLowerCase() === 'true' &&
      String(i.estoque_id || '').trim() === estId
    );

  const totalQtd = lotes.reduce((acc, l) => acc + parseNumeroBR(l.quantidade), 0);
  const totalCusto = lotes.reduce((acc, l) => acc + parseNumeroBR(l.custo_total), 0);
  if (totalQtd <= 0) return null;

  const custoMedio = Number((totalCusto / totalQtd).toFixed(6));
  updateById(
    ABA_ESTOQUE,
    'ID',
    estId,
    {
      custo_unitario: custoMedio,
      valor_unit: custoMedio
    },
    ESTOQUE_SCHEMA
  );

  return {
    ID: estId,
    custo_unitario: custoMedio,
    valor_unit: custoMedio
  };
}

function atualizarCustoLoteSaidaProducao(producaoId, loteId, custoUnitario) {
  assertCanWriteProducao('Atualizacao de custo de lote de saida');
  const prodId = String(producaoId || '').trim();
  const lotId = String(loteId || '').trim();
  const novoCusto = parseNumeroBR(custoUnitario);

  if (!prodId) {
    throw new Error('Producao invalida.');
  }
  if (!lotId) {
    throw new Error('Lote de saida invalido.');
  }
  if (novoCusto < 0) {
    throw new Error('Custo unitario nao pode ser negativo.');
  }

  const sheetProducao = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO);
  if (!sheetProducao) {
    throw new Error('Aba PRODUCAO nao encontrada');
  }
  const ordem = rowsToObjects(sheetProducao).find(i => String(i.producao_id || '').trim() === prodId);
  if (!ordem) {
    throw new Error('Producao nao encontrada');
  }
  if (String(ordem.estoque_atualizado).toLowerCase() !== 'true') {
    throw new Error('Custo de saida so pode ser editado apos atualizar o estoque da OP.');
  }

  const sheetLotes = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_SAIDAS_LOTES);
  if (!sheetLotes) {
    throw new Error('Aba PRODUCAO_SAIDAS_LOTES nao encontrada');
  }

  const lote = rowsToObjects(sheetLotes).find(i =>
    String(i.id || '').trim() === lotId &&
    String(i.producao_id || '').trim() === prodId &&
    String(i.ativo).toLowerCase() === 'true'
  );
  if (!lote) {
    throw new Error('Lote de saida nao encontrado para a OP informada.');
  }

  const quantidade = parseNumeroBR(lote.quantidade);
  const novoTotal = Number((quantidade * novoCusto).toFixed(6));
  updateById(
    ABA_PRODUCAO_SAIDAS_LOTES,
    'id',
    lotId,
    {
      custo_unitario: novoCusto,
      custo_total: novoTotal
    },
    PRODUCAO_SAIDAS_LOTES_SCHEMA
  );

  const estoqueAtualizado = recalcularCustoMedioSaidasNoEstoque(lote.estoque_id);
  const saidasLotes = listarSaidasLotesProducao(prodId);
  const loteAtualizado = saidasLotes.find(i => String(i.id || '').trim() === lotId) || null;

  return {
    ok: true,
    loteAtualizado,
    estoqueAtualizado,
    saidasLotes
  };
}

function atualizarPrecoVendaProdutoSaidaProducao(producaoId, nomeSaida, precoVendaInput, produtoRefIdInput) {
  assertCanWriteProducao('Atualizacao de preco de venda da saida da producao');
  const prodId = String(producaoId || '').trim();
  const nome = String(nomeSaida || '').trim();
  const produtoRefId = String(produtoRefIdInput || '').trim();
  const novoPreco = parseNumeroBR(precoVendaInput);

  if (!prodId) {
    throw new Error('Producao invalida.');
  }
  if (!nome) {
    throw new Error('Saida invalida.');
  }
  if (!isFinite(novoPreco) || novoPreco <= 0) {
    throw new Error('Preco de venda deve ser maior que zero.');
  }
  if (!produtoRefId) {
    throw new Error('Saida sem produto vinculado por ID. Ajuste o modelo na aba Produtos.');
  }

  const sheetProducao = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO);
  if (!sheetProducao) {
    throw new Error('Aba PRODUCAO nao encontrada');
  }
  const ordem = rowsToObjects(sheetProducao).find(i =>
    String(i.producao_id || '').trim() === prodId &&
    String(i.ativo).toLowerCase() === 'true'
  );
  if (!ordem) {
    throw new Error('Producao nao encontrada');
  }

  const nomeNorm = normalizarTextoSemAcentoProducao(nome);
  const lotesDaOp = listarSaidasLotesProducao(prodId);
  const pertenceLote = lotesDaOp.some(l => {
    const refIdLote = String(l?.produto_ref_id || '').trim();
    if (refIdLote !== produtoRefId) return false;
    if (!nomeNorm) return true;
    return normalizarTextoSemAcentoProducao(l?.nome_saida) === nomeNorm;
  });

  let pertencePrevista = false;
  if (!pertenceLote && ordem.receita_id) {
    try {
      const saidasPrevistas = agruparSaidasReceitaParaEstoque(
        (explodirSaidasReceitaDetalhada(
          ordem.produto_id,
          ordem.receita_id,
          normalizarQuantidadeInteiraProducao(ordem.qtd_planejada)
        )?.itens) || []
      );
      pertencePrevista = saidasPrevistas.some(s => {
        const refIdPrev = String(s?.produto_ref_id || '').trim();
        if (refIdPrev !== produtoRefId) return false;
        if (!nomeNorm) return true;
        return normalizarTextoSemAcentoProducao(s?.nome_saida) === nomeNorm;
      });
    } catch (error) {
      pertencePrevista = false;
    }
  }

  if (!pertenceLote && !pertencePrevista) {
    throw new Error('Produto vinculado nao pertence a esta saida da OP.');
  }

  const produto = obterProdutoAtivoPorIdSaidaProducao(produtoRefId);
  if (!produto || !produto.produto_id) {
    throw new Error('Produto vinculado nao encontrado ou inativo.');
  }

  const precoVenda = Number(novoPreco.toFixed(2));
  const ok = updateById(
    ABA_PRODUTOS,
    'produto_id',
    produto.produto_id,
    { preco_venda: precoVenda },
    PRODUTOS_SCHEMA
  );
  if (!ok) {
    throw new Error('Falha ao atualizar preco do produto vinculado.');
  }

  return {
    ok: true,
    produtoCriado: false,
    produtoAtualizado: {
      produto_id: produto.produto_id,
      nome_produto: produto.nome_produto || nome,
      preco_venda: precoVenda
    },
    saidasLotes: listarSaidasLotesProducao(prodId)
  };
}
