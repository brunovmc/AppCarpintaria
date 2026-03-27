const ABA_PRODUCAO = 'PRODUCAO';
const ABA_PRODUCAO_ETAPAS = 'PRODUCAO_ETAPAS';
const ABA_PRODUCAO_CONSUMO = 'PRODUCAO_CONSUMO';
const ABA_PRODUCAO_MATERIAIS = 'PRODUCAO_MATERIAIS';
const ABA_PRODUCAO_MATERIAIS_PREVISTOS = 'PRODUCAO_MATERIAIS_PREVISTOS';
const ABA_PRODUCAO_VINCULOS = 'PRODUCAO_VINCULOS_MATERIAIS';
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
  'status',
  'data_inicio',
  'data_prevista_termino',
  'data_conclusao',
  'observacao',
  'estoque_atualizado',
  'data_estoque_atualizado',
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

function listarProdutosAtivosMapeadosPorNomeProducao() {
  const produtosSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS);
  const produtosRows = produtosSheet ? rowsToObjects(produtosSheet) : [];
  const mapa = {};

  produtosRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(p => {
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

      const atual = mapa[nomeNorm];
      if (!atual) {
        mapa[nomeNorm] = proximo;
        return;
      }

      const precoAtual = parseNumeroBR(atual.preco_venda);
      const precoProximo = parseNumeroBR(proximo.preco_venda);
      if (precoAtual <= 0 && precoProximo > 0) {
        mapa[nomeNorm] = proximo;
      }
    });

  return mapa;
}

function obterProdutoAtivoPorNomeSaidaProducao(nomeSaida, mapaProdutosPorNome) {
  const nomeNorm = normalizarTextoSemAcentoProducao(nomeSaida);
  if (!nomeNorm) return null;

  const mapa = (mapaProdutosPorNome && typeof mapaProdutosPorNome === 'object')
    ? mapaProdutosPorNome
    : listarProdutosAtivosMapeadosPorNomeProducao();

  return mapa[nomeNorm] || null;
}

function listarProducao(forcarRecarregar) {
  if (!forcarRecarregar) {
    const cached = lerCacheProducao();
    if (Array.isArray(cached)) {
      return cached;
    }
  }

  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO);
  if (!sheet) {
    salvarCacheProducao([]);
    return [];
  }

  const produtosSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS);
  const produtos = produtosSheet ? rowsToObjects(produtosSheet) : [];
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

  const receitasSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  const receitasRows = receitasSheet ? rowsToObjects(receitasSheet) : [];
  const receitasMap = {};
  receitasRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(r => {
      receitasMap[r.receita_id] = r;
    });

  const saidasReceitaSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS_SAIDAS);
  const saidasReceitaRows = saidasReceitaSheet ? rowsToObjects(saidasReceitaSheet) : [];
  const saidasReceitaMap = {};
  saidasReceitaRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(s => {
      const receitaId = String(s.receita_id || '').trim();
      if (!receitaId) return;
      if (!saidasReceitaMap[receitaId]) saidasReceitaMap[receitaId] = [];
      saidasReceitaMap[receitaId].push({
        nome_saida: String(s.nome_saida || '').trim(),
        tipo_item: String(s.tipo_item || '').trim().toUpperCase() || 'PRODUTO',
        categoria: String(s.categoria || '').trim(),
        unidade: String(s.unidade || '').trim(),
        quantidade_base: parseNumeroBR(s.quantidade)
      });
    });

  const etapasSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_ETAPAS);
  const etapasRows = etapasSheet ? rowsToObjects(etapasSheet) : [];
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

  const rows = rowsToObjects(sheet);
  const lista = rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .map(i => {
      const prod = produtosMap[i.produto_id] || {};
      const receita = receitasMap[i.receita_id] || {};
      const qtdPlanejada = parseNumeroBR(i.qtd_planejada);
      const saidasBaseReceita = Array.isArray(saidasReceitaMap[i.receita_id])
        ? saidasReceitaMap[i.receita_id]
        : [];
      const saidasPrevistasBase = agruparSaidasReceitaParaEstoque(
        saidasBaseReceita
          .map(s => {
            const quantidade = parseNumeroBR(s.quantidade_base) * qtdPlanejada;
            if (quantidade <= 0) return null;
            return {
              nome_saida: s.nome_saida || prod.nome_produto || 'Saida',
              tipo_item: s.tipo_item || 'PRODUTO',
              categoria: s.categoria || '',
              unidade: s.unidade || '',
              quantidade
            };
          })
          .filter(v => !!v)
      );
      const saidasPrevistas = saidasPrevistasBase.map(s => {
        const produtoVinculado = obterProdutoAtivoPorNomeSaidaProducao(
          s.nome_saida,
          produtosAtivosPorNome
        );
        return {
          ...s,
          produto_id_vinculado: produtoVinculado?.produto_id || '',
          preco_venda_produto: produtoVinculado?.preco_venda || '',
          permite_editar_preco_venda: true
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
        saidas_previstas: saidasPrevistas,
        saidas_previstas_tipos: saidasPrevistas.length,
        saidas_previstas_total: parseNumeroBR(totalSaidasPrevistas),
        valor_previsto_venda_pecas: Number(valorPrevistoVendaPecas.toFixed(2)),
        estoque_atualizado: String(i.estoque_atualizado).toLowerCase() === 'true',
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
        etapas: etapasMap[i.producao_id] || []
      };
    });

  salvarCacheProducao(lista);
  return lista;
}

function criarProducao(payload) {
  assertCanWriteProducao('Criacao de producao');
  const novo = {
    ...payload,
    producao_id: gerarId('OP'),
    status: payload.status || 'Em planejamento',
    qtd_planejada: parseNumeroBR(payload.qtd_planejada),
    estoque_atualizado: false,
    data_estoque_atualizado: '',
    ativo: true,
    criado_em: new Date()
  };

  const detalhado = explodirReceitaDetalhada(
    novo.produto_id,
    novo.receita_id,
    novo.qtd_planejada
  );

  insert(ABA_PRODUCAO, novo, PRODUCAO_SCHEMA);
  salvarVinculosMateriaisProducao(novo.producao_id, novo.receita_id, detalhado.itens || []);
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

  const dadosAtualizados = {
    ...atual,
    ...payload
  };

  const mudouReceita = Object.prototype.hasOwnProperty.call(payload, 'receita_id') &&
    payload.receita_id !== atual.receita_id;
  const mudouProduto = Object.prototype.hasOwnProperty.call(payload, 'produto_id') &&
    payload.produto_id !== atual.produto_id;
  const mudouQtd = Object.prototype.hasOwnProperty.call(payload, 'qtd_planejada') &&
    parseNumeroBR(payload.qtd_planejada) !== parseNumeroBR(atual.qtd_planejada);

  let materiaisRegerados = null;
  if (mudouReceita || mudouProduto || mudouQtd) {
    const detalhado = explodirReceitaDetalhada(
      dadosAtualizados.produto_id,
      dadosAtualizados.receita_id,
      parseNumeroBR(dadosAtualizados.qtd_planejada)
    );
    materiaisRegerados = {
      detalhado: detalhado.itens || []
    };
  }

  const ok = updateById(
    ABA_PRODUCAO,
    'producao_id',
    id,
    payload,
    PRODUCAO_SCHEMA
  );

  if (ok && materiaisRegerados) {
    salvarVinculosMateriaisProducao(
      id,
      dadosAtualizados.receita_id,
      materiaisRegerados.detalhado,
      { preservarManuais: true }
    );
    recalcularMateriaisPrevistosFromVinculos(id);
  }

  return ok;
}

function deletarProducao(id) {
  assertCanWriteProducao('Exclusao de producao');
  return updateById(
    ABA_PRODUCAO,
    'producao_id',
    id,
    { ativo: false },
    PRODUCAO_SCHEMA
  );
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

function limparMateriaisPrevistosSnapshot(producaoId) {
  assertCanWriteProducao('Limpeza de snapshot de materiais previstos');
  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(ABA_PRODUCAO_MATERIAIS_PREVISTOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUCAO_MATERIAIS_PREVISTOS);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idCol = headers.indexOf('producao_id');

  if (idCol !== -1 && data.length > 1) {
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][idCol] === producaoId) {
        sheet.deleteRow(i + 1);
      }
    }
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
  assertCanWriteProducao('Limpeza de vinculos de materiais da producao');
  const cfg = opcoes || {};
  const preservarManuais = !!cfg.preservarManuais;

  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(ABA_PRODUCAO_VINCULOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUCAO_VINCULOS);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idCol = headers.indexOf('producao_id');
  const receitaEntradaCol = headers.indexOf('receita_entrada_id');

  if (idCol !== -1 && data.length > 1) {
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][idCol] !== producaoId) continue;

      if (preservarManuais && receitaEntradaCol !== -1) {
        const receitaEntradaId = String(data[i][receitaEntradaCol] || '').trim();
        if (!receitaEntradaId) continue;
      }

      sheet.deleteRow(i + 1);
    }
  }
}

function salvarVinculosMateriaisProducao(producaoId, receitaId, vinculos, opcoes) {
  assertCanWriteProducao('Salvamento de vinculos de materiais da producao');
  const cfg = opcoes || {};
  const preservarManuais = !!cfg.preservarManuais;

  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(ABA_PRODUCAO_VINCULOS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUCAO_VINCULOS);
  }

  ensureSchema(sheet, PRODUCAO_VINCULOS_SCHEMA);
  limparVinculosMateriaisProducao(producaoId, { preservarManuais });

  const linhasValidas = Array.isArray(vinculos) ? vinculos : [];
  linhasValidas.forEach(v => {
    const quantidadePrevista = parseNumeroBR(v.quantidade);
    if (!quantidadePrevista || quantidadePrevista <= 0) return;

    const estoqueId = String(v.estoque_id || '').trim();
    const statusInicial = estoqueId ? 'Pendente' : 'Sem vinculo';

    const novo = {
      id: gerarId('PVM'),
      producao_id: producaoId,
      receita_id: receitaId || v.receita_id || '',
      receita_entrada_id: v.receita_entrada_id || '',
      estoque_id: estoqueId,
      tipo_item: v.tipo_item || '',
      origem_item: v.origem_item || '',
      unidade: v.unidade || '',
      quantidade_prevista: quantidadePrevista,
      quantidade_consumida: parseNumeroBR(v.quantidade_consumida),
      item_snapshot: v.item || '',
      valor_unit_snapshot: parseNumeroBR(v.valor_unit),
      status: statusInicial,
      criado_em: new Date(),
      ativo: true
    };

    insert(ABA_PRODUCAO_VINCULOS, novo, PRODUCAO_VINCULOS_SCHEMA);
  });

  invalidarCachesRelacionadosAba(ABA_PRODUCAO);
  return true;
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

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idxProducao = headers.indexOf('producao_id');
  const idxReceitaSaida = headers.indexOf('receita_saida_id');

  if (idxProducao === -1 || data.length <= 1) return;

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][idxProducao] !== producaoId) continue;

    if (preservarManuais && idxReceitaSaida !== -1) {
      const receitaSaidaId = String(data[i][idxReceitaSaida] || '').trim();
      if (!receitaSaidaId) continue;
    }

    sheet.deleteRow(i + 1);
  }
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
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_VINCULOS);
  if (!sheet) {
    salvarMateriaisPrevistosSnapshot(producaoId, []);
    return { itens: [], custoPrevisto: 0 };
  }

  const rows = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.producao_id === producaoId)
    .filter(i => String(i.estoque_id || '').trim() !== '');

  const itens = rows.map(r => ({
    estoque_id: String(r.estoque_id || '').trim(),
    item: r.item_snapshot || r.origem_item || r.estoque_id || '',
    unidade: r.unidade || '',
    quantidade: parseNumeroBR(r.quantidade_prevista),
    valor_unit: parseNumeroBR(r.valor_unit_snapshot)
  })).filter(i => i.estoque_id && i.quantidade > 0);

  const agrupado = agruparMateriaisPorEstoque(itens);
  salvarMateriaisPrevistosSnapshot(producaoId, agrupado.itens || []);

  return {
    itens: agrupado.itens || [],
    custoPrevisto: parseNumeroBR(agrupado.custoPrevisto)
  };
}

function listarVinculosMateriaisProducao(producaoId) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_VINCULOS);
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

  const lista = rows.map(r => {
    const estoqueId = String(r.estoque_id || '').trim();
    const vinculado = !!estoqueId;
    const estoqueItem = estoqueMap[r.estoque_id] || {};
    const quantidadePrevista = parseNumeroBR(r.quantidade_prevista);
    const quantidadeConsumida = parseNumeroBR(r.quantidade_consumida);
    const quantidadePendente = Math.max(quantidadePrevista - quantidadeConsumida, 0);
    const saldoEstoque = parseNumeroBR(estoqueItem.quantidade);

    let status = String(r.status || '').trim();
    if (!vinculado) {
      status = 'Sem vinculo';
    } else if (!status || status === 'Sem vinculo') {
      status = quantidadeConsumida <= 0
        ? 'Pendente'
        : (quantidadePendente <= 0 ? 'Concluido' : 'Parcial');
    }

    return {
      id: r.id,
      producao_id: r.producao_id,
      receita_id: r.receita_id || '',
      receita_entrada_id: r.receita_entrada_id || '',
      estoque_id: estoqueId,
      tipo_item: r.tipo_item || '',
      origem_item: r.origem_item || '',
      item: r.item_snapshot || estoqueItem.item || r.origem_item || r.estoque_id,
      unidade: r.unidade || estoqueItem.unidade || '',
      quantidade_prevista: quantidadePrevista,
      quantidade_consumida: quantidadeConsumida,
      quantidade_pendente: quantidadePendente,
      valor_unit: parseNumeroBR(r.valor_unit_snapshot || estoqueItem.valor_unit),
      saldo_estoque: saldoEstoque,
      saldo_suficiente: vinculado ? (saldoEstoque >= quantidadePendente) : false,
      vinculado,
      status
    };
  });

  return lista.sort((a, b) => {
    const ta = `${a.tipo_item || ''} ${a.item || ''}`.toLowerCase();
    const tb = `${b.tipo_item || ''} ${b.item || ''}`.toLowerCase();
    return ta.localeCompare(tb);
  });
}

function vincularItemProducaoAoEstoque(producaoId, vinculoId, estoqueId) {
  assertCanWriteProducao('Vinculacao de item da producao ao estoque');
  if (!producaoId || !vinculoId) {
    throw new Error('Vinculo invalido');
  }
  const estoqueIdNorm = String(estoqueId || '').trim();
  if (!estoqueIdNorm) {
    throw new Error('Selecione um item do estoque');
  }

  const sheetVinculos = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_VINCULOS);
  if (!sheetVinculos) {
    throw new Error('Aba de vinculos nao encontrada');
  }

  const vinculo = rowsToObjects(sheetVinculos)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .find(i => i.id === vinculoId && i.producao_id === producaoId);
  if (!vinculo) {
    throw new Error('Vinculo nao encontrado');
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

  updateById(
    ABA_PRODUCAO_VINCULOS,
    'id',
    vinculoId,
    {
      estoque_id: estoqueIdNorm,
      item_snapshot: estoque.item || vinculo.item_snapshot || vinculo.origem_item || estoqueIdNorm,
      unidade: estoque.unidade || vinculo.unidade || '',
      valor_unit_snapshot: parseNumeroBR(estoque.valor_unit),
      status: 'Pendente'
    },
    PRODUCAO_VINCULOS_SCHEMA
  );

  const materiais = recalcularMateriaisPrevistosFromVinculos(producaoId);
  return {
    vinculos: listarVinculosMateriaisProducao(producaoId),
    materiaisPrevistos: materiais.itens || [],
    custoPrevisto: parseNumeroBR(materiais.custoPrevisto)
  };
}

function vincularPendenciasEntradaProducao(producaoId, vinculacoes) {
  assertCanWriteProducao('Vinculacao de pendencias da producao');
  if (!producaoId) {
    throw new Error('Producao invalida');
  }

  const itens = Array.isArray(vinculacoes) ? vinculacoes : [];
  if (itens.length === 0) {
    throw new Error('Nenhum vinculo informado');
  }

  const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
  if (!sheetEstoque) {
    throw new Error('Aba ESTOQUE nao encontrada');
  }

  const estoqueAtivos = rowsToObjects(sheetEstoque)
    .filter(i => String(i.ativo).toLowerCase() === 'true');
  const estoqueMap = {};
  estoqueAtivos.forEach(i => {
    estoqueMap[i.ID] = i;
  });

  const sheetVinculos = getDataSpreadsheet().getSheetByName(ABA_PRODUCAO_VINCULOS);
  if (!sheetVinculos) {
    throw new Error('Aba de vinculos nao encontrada');
  }

  const vinculosAtivos = rowsToObjects(sheetVinculos)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.producao_id === producaoId);
  const vinculosMap = {};
  vinculosAtivos.forEach(v => {
    vinculosMap[v.id] = v;
  });

  itens.forEach(v => {
    const vinculoId = String(v?.vinculo_id || '').trim();
    const estoqueId = String(v?.estoque_id || '').trim();

    if (!vinculoId) {
      throw new Error('Vinculo invalido');
    }
    if (!estoqueId) {
      throw new Error('Selecione um item do estoque');
    }

    const vinculo = vinculosMap[vinculoId];
    if (!vinculo) {
      throw new Error(`Vinculo nao encontrado: ${vinculoId}`);
    }

    const estoque = estoqueMap[estoqueId];
    if (!estoque) {
      throw new Error(`Item de estoque nao encontrado: ${estoqueId}`);
    }

    updateById(
      ABA_PRODUCAO_VINCULOS,
      'id',
      vinculoId,
      {
        estoque_id: estoqueId,
        item_snapshot: estoque.item || vinculo.item_snapshot || vinculo.origem_item || estoqueId,
        unidade: estoque.unidade || vinculo.unidade || '',
        valor_unit_snapshot: parseNumeroBR(estoque.valor_unit),
        status: 'Pendente'
      },
      PRODUCAO_VINCULOS_SCHEMA
    );
  });

  const materiais = recalcularMateriaisPrevistosFromVinculos(producaoId);
  return {
    vinculos: listarVinculosMateriaisProducao(producaoId),
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

  let itemSnapshot = nomeItem;
  let unidade = String(dados.unidade || '').trim();
  let valorUnitSnapshot = parseNumeroBR(dados.valor_unit_snapshot);
  let status = estoqueId ? 'Pendente' : 'Sem vinculo';

  if (estoqueId) {
    const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
    const estoque = sheetEstoque
      ? rowsToObjects(sheetEstoque).find(i => i.ID === estoqueId && String(i.ativo).toLowerCase() === 'true')
      : null;
    if (!estoque) {
      throw new Error('Item de estoque nao encontrado');
    }
    itemSnapshot = estoque.item || nomeItem;
    unidade = estoque.unidade || unidade;
    valorUnitSnapshot = parseNumeroBR(estoque.valor_unit);
  }

  const novo = {
    id: gerarId('PVM'),
    producao_id: producaoId,
    receita_id: '',
    receita_entrada_id: '',
    estoque_id: estoqueId,
    tipo_item: tipoItem,
    origem_item: nomeItem,
    unidade: unidade || '',
    quantidade_prevista: quantidadePrevista,
    quantidade_consumida: 0,
    item_snapshot: itemSnapshot,
    valor_unit_snapshot: valorUnitSnapshot,
    status,
    criado_em: new Date(),
    ativo: true
  };

  insert(ABA_PRODUCAO_VINCULOS, novo, PRODUCAO_VINCULOS_SCHEMA);
  const materiais = recalcularMateriaisPrevistosFromVinculos(producaoId);

  return {
    item: novo,
    vinculos: listarVinculosMateriaisProducao(producaoId),
    materiaisPrevistos: materiais.itens || [],
    custoPrevisto: parseNumeroBR(materiais.custoPrevisto)
  };
}

function regerarMateriaisEObjetosDeVinculo(producaoId, produtoId, receitaId, qtdPlanejada) {
  const detalhado = explodirReceitaDetalhada(produtoId, receitaId, qtdPlanejada);
  const itensDetalhados = Array.isArray(detalhado && detalhado.itens) ? detalhado.itens : [];

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
        parseNumeroBR(o.qtd_planejada)
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
    throw new Error('Receita nao informada');
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
  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(ABA_PRODUCAO_VINCULOS);
  if (!sheet) return;

  ensureSchema(sheet, PRODUCAO_VINCULOS_SCHEMA);

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idxProducao = headers.indexOf('producao_id');
  const idxEstoque = headers.indexOf('estoque_id');
  const idxPrevista = headers.indexOf('quantidade_prevista');
  const idxConsumida = headers.indexOf('quantidade_consumida');
  const idxStatus = headers.indexOf('status');
  const idxAtivo = headers.indexOf('ativo');

  if (
    idxProducao === -1 ||
    idxEstoque === -1 ||
    idxPrevista === -1 ||
    idxConsumida === -1 ||
    idxStatus === -1
  ) {
    return;
  }

  const consumoMap = {};
  (Array.isArray(itensConsumidos) ? itensConsumidos : []).forEach(i => {
    const estoqueId = i && i.estoque_id ? String(i.estoque_id) : '';
    const qtd = parseNumeroBR(i ? i.quantidade : 0);
    if (!estoqueId || qtd <= 0) return;
    consumoMap[estoqueId] = (consumoMap[estoqueId] || 0) + qtd;
  });

  if (Object.keys(consumoMap).length === 0) return;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idxProducao] !== producaoId) continue;
    if (idxAtivo !== -1 && String(row[idxAtivo]).toLowerCase() !== 'true') continue;

    const estoqueId = String(row[idxEstoque] || '');
    let restante = consumoMap[estoqueId] || 0;
    if (restante <= 0) continue;

    const qtdPrevista = parseNumeroBR(row[idxPrevista]);
    const qtdConsumidaAtual = parseNumeroBR(row[idxConsumida]);
    const qtdPendente = Math.max(qtdPrevista - qtdConsumidaAtual, 0);
    if (qtdPendente <= 0) continue;

    const qtdAplicada = Math.min(qtdPendente, restante);
    const novaConsumida = qtdConsumidaAtual + qtdAplicada;
    consumoMap[estoqueId] = restante - qtdAplicada;

    const novoStatus = novaConsumida <= 0
      ? 'Pendente'
      : (novaConsumida >= qtdPrevista ? 'Concluido' : 'Parcial');

    sheet.getRange(i + 1, idxConsumida + 1).setValue(novaConsumida);
    sheet.getRange(i + 1, idxStatus + 1).setValue(novoStatus);
  }
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
        ordem.qtd_planejada
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
      const estoqueId = String(v.estoque_id || '').trim();
      if (!estoqueId) return true;
      return !mapaAtivos[estoqueId];
    })
    .map(v => ({
      id: v.id,
      producao_id: v.producao_id,
      receita_entrada_id: v.receita_entrada_id || '',
      estoque_id: String(v.estoque_id || '').trim(),
      tipo_item: v.tipo_item || '',
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
    const tipoItem = normalizarTipoSaidaProducao(s.tipo_item);
    const categoria = normalizarCategoriaSaidaProducao(tipoItem, s.categoria);
    const unidade = String(s.unidade || '').trim();
    const quantidade = parseNumeroBR(s.quantidade);

    if (!nomeSaida || quantidade <= 0) return;

    const chave = `${normalizarTextoProducao(tipoItem)}||${normalizarTextoProducao(categoria)}||${normalizarTextoProducao(nomeSaida)}||${normalizarTextoProducao(unidade)}`;
    if (!map[chave]) {
      map[chave] = {
        nome_saida: nomeSaida,
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

function consumirEstoque(producaoId, itensParaBaixar) {
  assertCanWriteProducao('Consumo de estoque da producao');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (!producaoId) {
      throw new Error('Producao invalida');
    }

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
      throw new Error('Receita nao informada');
    }

    const status = String(ordem.status || '');
    if (status !== 'Concluido' && status !== 'Concluida') {
      throw new Error('Atualizacao de estoque liberada apenas quando status for Concluido');
    }

    if (String(ordem.estoque_atualizado).toLowerCase() === 'true') {
      throw new Error('Estoque ja atualizado para esta producao');
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
        estoqueMapAtivo[i.ID] = i;
      });

    const pendenciasVinculo = listarPendenciasVinculoEntradaProducao(producaoId, estoqueMapAtivo);
    if (pendenciasVinculo.length > 0) {
      const materiais = recalcularMateriaisPrevistosFromVinculos(producaoId);
      return {
        requer_vinculos: true,
        pendencias_vinculo: pendenciasVinculo,
        vinculos: listarVinculosMateriaisProducao(producaoId),
        materiaisPrevistos: materiais.itens || [],
        custoPrevisto: parseNumeroBR(materiais.custoPrevisto)
      };
    }

    const itens = Array.isArray(itensParaBaixar) ? itensParaBaixar : [];
    const itensMap = {};
    itens.forEach(i => {
      const estoqueId = i && i.estoque_id ? String(i.estoque_id) : '';
      const qtd = parseNumeroBR(i ? i.quantidade : 0);
      if (!estoqueId || qtd <= 0) return;
      itensMap[estoqueId] = (itensMap[estoqueId] || 0) + qtd;
    });

    const itensValidos = Object.keys(itensMap).map(estoqueId => ({
      estoque_id: estoqueId,
      quantidade: itensMap[estoqueId]
    }));

    if (itensValidos.length === 0) {
      throw new Error('Nenhum item para baixar');
    }

    // Obriga informar consumo para todos os materiais previstos da ordem.
    const previstosResp = obterMateriaisPrevistosProducao(producaoId);
    const previstos = Array.isArray(previstosResp && previstosResp.itens)
      ? previstosResp.itens
      : [];
    const previstosMap = {};
    previstos.forEach(i => {
      const estoqueId = i && i.estoque_id ? String(i.estoque_id) : '';
      const qtd = parseNumeroBR(i ? i.quantidade : 0);
      if (!estoqueId || qtd <= 0) return;
      previstosMap[estoqueId] = (previstosMap[estoqueId] || 0) + qtd;
    });

    const faltantes = Object.keys(previstosMap).filter(estoqueId => !itensMap[estoqueId]);
    if (faltantes.length > 0) {
      const nomes = faltantes.map(id => {
        const it = estoqueMapAtivo[id];
        return it ? (it.item || id) : id;
      });
      throw new Error(`Consumo incompleto. Informe todos os materiais previstos: ${nomes.join(', ')}`);
    }

    const naoPrevistos = Object.keys(itensMap).filter(estoqueId => !previstosMap[estoqueId]);
    if (naoPrevistos.length > 0) {
      const nomes = naoPrevistos.map(id => {
        const it = estoqueMapAtivo[id];
        return it ? (it.item || id) : id;
      });
      throw new Error(`Itens nao previstos informados para baixa: ${nomes.join(', ')}`);
    }

    itensValidos.forEach(i => {
      const estoqueItem = estoqueMapAtivo[i.estoque_id];
      if (!estoqueItem) {
        throw new Error(`Item de estoque nao encontrado: ${i.estoque_id}`);
      }
      const saldo = parseNumeroBR(estoqueItem.quantidade);
      if (saldo < i.quantidade) {
        throw new Error(`Saldo insuficiente para ${estoqueItem.item || i.estoque_id}`);
      }
    });

    const consumoRegistrado = [];
    const estoqueAtualizados = [];

    itensValidos.forEach(i => {
      const estoqueItem = estoqueMapAtivo[i.estoque_id];
      const saldo = parseNumeroBR(estoqueItem.quantidade);
      const novoSaldo = saldo - i.quantidade;

      updateById(
        ABA_ESTOQUE,
        'ID',
        i.estoque_id,
        { quantidade: novoSaldo },
        ESTOQUE_SCHEMA
      );

      estoqueMapAtivo[i.estoque_id].quantidade = novoSaldo;

      estoqueAtualizados.push({
        ID: i.estoque_id,
        quantidade: novoSaldo
      });

      const valorUnit = parseNumeroBR(estoqueItem.custo_unitario || estoqueItem.valor_unit);
      const total = i.quantidade * valorUnit;

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

    const saidasExplodidas = explodirSaidasReceitaDetalhada(
      ordem.produto_id,
      ordem.receita_id,
      parseNumeroBR(ordem.qtd_planejada)
    );
    const saidasAgrupadas = agruparSaidasReceitaParaEstoque(saidasExplodidas.itens || []);
    if (saidasAgrupadas.length === 0) {
      throw new Error('Receita sem saidas configuradas para lancamento no estoque');
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

    const dataAtualizacao = new Date();
    updateById(
      ABA_PRODUCAO,
      'producao_id',
      producaoId,
      {
        estoque_atualizado: true,
        data_estoque_atualizado: dataAtualizacao
      },
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
      producaoAtualizada: {
        producao_id: producaoId,
        estoque_atualizado: true,
        data_estoque_atualizado: formatDateSafe(dataAtualizacao, 'yyyy-MM-dd')
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
  const produtosPorNome = listarProdutosAtivosMapeadosPorNomeProducao();

  return rowsToObjects(sheet)
    .filter(i =>
      String(i.ativo).toLowerCase() === 'true' &&
      String(i.producao_id || '').trim() === prodId
    )
    .map(i => {
      const produtoVinculado = obterProdutoAtivoPorNomeSaidaProducao(
        i.nome_saida,
        produtosPorNome
      );
      return {
        ...i,
        quantidade: parseNumeroBR(i.quantidade),
        custo_unitario: parseNumeroBR(i.custo_unitario),
        custo_total: parseNumeroBR(i.custo_total),
        preco_venda_produto: produtoVinculado?.preco_venda || '',
        produto_id_vinculado: produtoVinculado?.produto_id || '',
        permite_editar_preco_venda: true,
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

function atualizarPrecoVendaProdutoSaidaProducao(producaoId, nomeSaida, precoVendaInput) {
  assertCanWriteProducao('Atualizacao de preco de venda da saida da producao');
  const prodId = String(producaoId || '').trim();
  const nome = String(nomeSaida || '').trim();
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

  const produtosPorNome = listarProdutosAtivosMapeadosPorNomeProducao();
  const precoVenda = Number(novoPreco.toFixed(2));
  let produto = obterProdutoAtivoPorNomeSaidaProducao(nome, produtosPorNome);
  let produtoCriado = false;

  if (!produto || !produto.produto_id) {
    const novoProduto = {
      produto_id: gerarId('PRD'),
      nome_produto: nome,
      unidade_produto: 'UN',
      preco_venda: precoVenda,
      ativo: true,
      criado_em: new Date()
    };
    insert(ABA_PRODUTOS, novoProduto, PRODUTOS_SCHEMA);
    produtoCriado = true;
    produto = {
      produto_id: novoProduto.produto_id,
      nome_produto: novoProduto.nome_produto,
      preco_venda
    };
  } else {
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
  }

  return {
    ok: true,
    produtoCriado,
    produtoAtualizado: {
      produto_id: produto.produto_id,
      nome_produto: produto.nome_produto || nome,
      preco_venda: precoVenda
    },
    saidasLotes: listarSaidasLotesProducao(prodId)
  };
}
