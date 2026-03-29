const ABA_PRODUTOS = 'PRODUTOS';
const ABA_PRODUTOS_COMPONENTES = 'PRODUTOS_COMPONENTES';
const ABA_PRODUTOS_ETAPAS = 'PRODUTOS_ETAPAS';
const ABA_PRODUTOS_RECEITAS = 'PRODUTOS_RECEITAS';
const ABA_PRODUTOS_RECEITAS_ENTRADAS = 'PRODUTOS_RECEITAS_ENTRADAS';
const ABA_PRODUTOS_RECEITAS_SAIDAS = 'PRODUTOS_RECEITAS_SAIDAS';
const PRODUTOS_CACHE_SCOPE = 'PRODUTOS_LISTA_ATIVOS';
const PRODUTOS_CACHE_TTL_SEC = 120;

const PRODUTOS_SCHEMA = [
  'produto_id',
  'nome_produto',
  'unidade_produto',
  'preco_venda',
  'ativo',
  'criado_em'
];

const PRODUTOS_COMPONENTES_SCHEMA = [
  'id',
  'produto_id',
  'tipo_componente',
  'ref_id',
  'quantidade',
  'unidade',
  'observacao',
  'ativo'
];

const PRODUTOS_ETAPAS_SCHEMA = [
  'id',
  'produto_id',
  'nome_etapa',
  'ordem',
  'ativo'
];

const PRODUTOS_RECEITAS_SCHEMA = [
  'receita_id',
  'produto_id',
  'nome_receita',
  'descricao',
  'parent_receita_id',
  'ativo',
  'criado_em'
];

const PRODUTOS_RECEITAS_ENTRADAS_SCHEMA = [
  'id',
  'receita_id',
  'tipo_item',
  'nome_item',
  'estoque_ref_id',
  'produto_ref_id',
  'receita_ref_id',
  'categoria',
  'unidade',
  'qtd_pecas',
  'comprimento_cm',
  'largura_cm',
  'espessura_cm',
  'custo_manual',
  'observacao',
  'ativo'
];

const PRODUTOS_RECEITAS_SAIDAS_SCHEMA = [
  'id',
  'receita_id',
  'nome_saida',
  'produto_ref_id',
  'tipo_item',
  'categoria',
  'unidade',
  'quantidade',
  'ativo'
];

function normalizarTipoSaidaReceitaProduto(tipo) {
  const t = String(tipo || '').trim().toUpperCase();
  return t || 'PRODUTO';
}

function normalizarCategoriaSaidaReceitaProduto(tipo, categoria) {
  const tipoNorm = normalizarTipoSaidaReceitaProduto(tipo);
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

function normalizarChaveProdutoNomeSaida(valor) {
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

function listarProdutosAtivosIndicesSaidaReceita() {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS);
  const rows = sheet ? rowsToObjects(sheet) : [];
  const porId = {};
  const porNome = {};

  rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(p => {
      const produtoId = String(p.produto_id || '').trim();
      if (!produtoId) return;

      const registro = {
        produto_id: produtoId,
        nome_produto: String(p.nome_produto || '').trim(),
        preco_venda: p.preco_venda
      };

      porId[produtoId] = registro;

      const chaveNome = normalizarChaveProdutoNomeSaida(registro.nome_produto);
      if (!chaveNome) return;
      if (!Array.isArray(porNome[chaveNome])) {
        porNome[chaveNome] = [];
      }
      porNome[chaveNome].push(registro);
    });

  return { porId, porNome };
}

function resolverProdutoRefIdSaidaReceita(tipoItem, nomeSaida, produtoRefIdInput, indicesProdutos, opcoes) {
  const cfg = opcoes || {};
  const strict = cfg.strict === true;
  const produtoPadraoId = String(cfg.produtoPadraoId || '').trim();
  const nomeProdutoPadrao = String(cfg.nomeProdutoPadrao || '').trim();
  const totalSaidasProduto = Number(cfg.totalSaidasProduto || 0);
  const tipoNorm = normalizarTipoSaidaReceitaProduto(tipoItem);
  if (tipoNorm !== 'PRODUTO') return '';

  const produtoRefId = String(produtoRefIdInput || '').trim();
  const indices = indicesProdutos && typeof indicesProdutos === 'object'
    ? indicesProdutos
    : { porId: {}, porNome: {} };

  if (produtoRefId) {
    if (indices.porId && indices.porId[produtoRefId]) {
      return produtoRefId;
    }
    if (strict) {
      throw new Error('Produto vinculado na saida nao encontrado ou inativo.');
    }
    return '';
  }

  const chaveNome = normalizarChaveProdutoNomeSaida(nomeSaida);
  if (!chaveNome) return '';

  const candidatos = Array.isArray(indices.porNome?.[chaveNome])
    ? indices.porNome[chaveNome]
    : [];
  if (candidatos.length === 1) {
    return String(candidatos[0].produto_id || '').trim();
  }

  if (produtoPadraoId) {
    const chaveProdutoPadrao = normalizarChaveProdutoNomeSaida(nomeProdutoPadrao);
    if ((totalSaidasProduto === 1) || (chaveProdutoPadrao && chaveProdutoPadrao === chaveNome)) {
      if (indices.porId && indices.porId[produtoPadraoId]) {
        return produtoPadraoId;
      }
    }
  }

  return '';
}

function obterContextoProdutoPorReceitaId(receitaId) {
  const recId = String(receitaId || '').trim();
  if (!recId) {
    return {
      produto_id: '',
      nome_produto: ''
    };
  }

  const sheetReceitas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  const receita = sheetReceitas
    ? rowsToObjects(sheetReceitas).find(r =>
      String(r.receita_id || '').trim() === recId &&
      String(r.ativo).toLowerCase() === 'true'
    )
    : null;
  const produtoId = String(receita?.produto_id || '').trim();
  if (!produtoId) {
    return {
      produto_id: '',
      nome_produto: ''
    };
  }

  const sheetProdutos = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS);
  const produto = sheetProdutos
    ? rowsToObjects(sheetProdutos).find(p =>
      String(p.produto_id || '').trim() === produtoId &&
      String(p.ativo).toLowerCase() === 'true'
    )
    : null;

  return {
    produto_id: produtoId,
    nome_produto: String(produto?.nome_produto || '').trim()
  };
}

function obterValorUnitarioEstoqueParaCustoProduto(itemEstoque) {
  const tipo = String(itemEstoque?.tipo || '').trim().toUpperCase();
  const custoBruto = String(itemEstoque?.custo_unitario ?? '').trim();
  const custo = parseNumeroBR(custoBruto);

  if (tipo === 'PRODUTO') {
    if (custoBruto === '') return 0;
    return custo;
  }

  if (custoBruto !== '') return custo;
  return parseNumeroBR(itemEstoque?.valor_unit);
}

function normalizarPrecoVendaProdutoEntrada(valor) {
  const raw = String(valor ?? '').trim();
  if (raw === '') return '';

  const n = parseNumeroBR(valor);
  if (!isFinite(n) || n <= 0) {
    throw new Error('Preco de venda invalido. Informe valor maior que zero ou deixe em branco.');
  }
  return Number(n.toFixed(2));
}

function normalizarPrecoVendaProdutoSaida(valor) {
  const raw = String(valor ?? '').trim();
  if (raw === '') return '';

  const n = parseNumeroBR(valor);
  if (!isFinite(n) || n <= 0) return '';
  return Number(n.toFixed(2));
}

function assertCanWriteProdutos(acao) {
  assertCanWrite(acao || 'Operacao de produtos');
}

function lerCacheProdutos() {
  return appCacheGetJson(PRODUTOS_CACHE_SCOPE);
}

function salvarCacheProdutos(lista) {
  appCachePutJson(PRODUTOS_CACHE_SCOPE, Array.isArray(lista) ? lista : [], PRODUTOS_CACHE_TTL_SEC);
}

function limparCacheProdutos() {
  return appCacheRemove(PRODUTOS_CACHE_SCOPE);
}

function recarregarCacheProdutos() {
  limparCacheProdutos();
  const dados = listarProdutos(true);
  return {
    ok: true,
    scope: PRODUTOS_CACHE_SCOPE,
    ttl_segundos: PRODUTOS_CACHE_TTL_SEC,
    total_itens: Array.isArray(dados) ? dados.length : 0
  };
}

function getResumoModeloProduto(produtoId) {
  const modelo = obterModeloProduto(produtoId);
  const entradasTotal = Array.isArray(modelo?.entradas) ? modelo.entradas.length : 0;
  const saidasTotal = Array.isArray(modelo?.saidas) ? modelo.saidas.length : 0;
  const etapasTotal = Array.isArray(modelo?.etapas) ? modelo.etapas.length : 0;
  const modeloNome = String(modelo?.dados?.nome_receita || 'Modelo principal').trim() || 'Modelo principal';

  return {
    modelo_nome: modeloNome,
    modelo_entradas_total: entradasTotal,
    modelo_saidas_total: saidasTotal,
    modelo_etapas_total: etapasTotal,
    modelo_resumo: `${entradasTotal} entradas, ${saidasTotal} saidas, ${etapasTotal} etapas`
  };
}

function listarProdutos(forcarRecarregar) {
  if (!forcarRecarregar) {
    const cached = lerCacheProdutos();
    if (Array.isArray(cached)) {
      return cached;
    }
  }

  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS);
  if (!sheet) {
    salvarCacheProdutos([]);
    return [];
  }

  const produtosAtivos = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .map(i => ({ ...i }));

  const receitasSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  const receitasAtivas = receitasSheet
    ? rowsToObjects(receitasSheet).filter(i => String(i.ativo).toLowerCase() === 'true')
    : [];
  const receitasPorProduto = {};
  receitasAtivas.forEach(r => {
    const produtoId = String(r.produto_id || '').trim();
    if (!produtoId) return;
    if (!Array.isArray(receitasPorProduto[produtoId])) {
      receitasPorProduto[produtoId] = [];
    }
    receitasPorProduto[produtoId].push(r);
  });
  const receitaPrincipalPorProduto = {};
  Object.keys(receitasPorProduto).forEach(produtoId => {
    const receitas = receitasPorProduto[produtoId] || [];
    const semParent = receitas.find(r => !String(r.parent_receita_id || '').trim());
    receitaPrincipalPorProduto[produtoId] = semParent || receitas[0] || null;
  });

  const entradasSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS_ENTRADAS);
  const entradasAtivas = entradasSheet
    ? rowsToObjects(entradasSheet).filter(i => String(i.ativo).toLowerCase() === 'true')
    : [];
  const entradasPorReceita = {};
  entradasAtivas.forEach(e => {
    const receitaId = String(e.receita_id || '').trim();
    if (!receitaId) return;
    if (!Array.isArray(entradasPorReceita[receitaId])) {
      entradasPorReceita[receitaId] = [];
    }
    entradasPorReceita[receitaId].push(e);
  });

  const saidasSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS_SAIDAS);
  const saidasAtivas = saidasSheet
    ? rowsToObjects(saidasSheet).filter(i => String(i.ativo).toLowerCase() === 'true')
    : [];
  const saidasPorReceita = {};
  saidasAtivas.forEach(s => {
    const receitaId = String(s.receita_id || '').trim();
    if (!receitaId) return;
    if (!Array.isArray(saidasPorReceita[receitaId])) {
      saidasPorReceita[receitaId] = [];
    }
    saidasPorReceita[receitaId].push(s);
  });

  const etapasSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_ETAPAS);
  const etapasAtivas = etapasSheet
    ? rowsToObjects(etapasSheet).filter(i => String(i.ativo).toLowerCase() === 'true')
    : [];
  const etapasPorProduto = {};
  etapasAtivas.forEach(et => {
    const produtoId = String(et.produto_id || '').trim();
    if (!produtoId) return;
    if (!Array.isArray(etapasPorProduto[produtoId])) {
      etapasPorProduto[produtoId] = [];
    }
    etapasPorProduto[produtoId].push(et);
  });

  const estoqueSheet = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
  const estoqueMap = {};
  if (estoqueSheet) {
    rowsToObjects(estoqueSheet).forEach(item => {
      const estoqueId = String(item.ID || '').trim();
      if (!estoqueId) return;
      estoqueMap[estoqueId] = item;
    });
  }

  const custoProdutoCache = {};
  function calcularCustoProdutoListagem(produtoId) {
    const id = String(produtoId || '').trim();
    if (!id) return 0;
    if (Object.prototype.hasOwnProperty.call(custoProdutoCache, id)) {
      return custoProdutoCache[id];
    }

    const receitaPrincipal = receitaPrincipalPorProduto[id];
    if (!receitaPrincipal || !receitaPrincipal.receita_id) {
      throw new Error('Produto sem receita de producao');
    }

    const receitaId = String(receitaPrincipal.receita_id || '').trim();
    const entradas = Array.isArray(entradasPorReceita[receitaId]) ? entradasPorReceita[receitaId] : [];
    let custoPrevisto = 0;

    entradas.forEach(e => {
      const tipo = String(e.tipo_item || '').toUpperCase();
      const qtdBase = parseNumeroBR(e.qtd_pecas);
      if (!qtdBase || qtdBase <= 0) return;

      let valorUnit = parseNumeroBR(e.custo_manual);
      let quantidade = 0;

      if (tipo === 'MADEIRA') {
        const comp = parseNumeroBR(e.comprimento_cm);
        const larg = parseNumeroBR(e.largura_cm);
        const esp = parseNumeroBR(e.espessura_cm);
        const volumeM3 = (comp > 0 && larg > 0 && esp > 0)
          ? ((comp * larg * esp) / 1000000)
          : 0;
        const unidadeEntrada = String(e.unidade || '').trim().toUpperCase();
        const qtdInteira = Math.abs(qtdBase - Math.round(qtdBase)) < 0.000001;
        const modoLegadoQtdPecas = unidadeEntrada !== 'M3' && volumeM3 > 0 && qtdInteira;
        quantidade = modoLegadoQtdPecas
          ? (qtdBase * volumeM3)
          : qtdBase;
      } else {
        quantidade = qtdBase;
      }

      const estoqueId = String(e.estoque_ref_id || '').trim();
      if (estoqueId) {
        const estoqueItem = estoqueMap[estoqueId] || null;
        if (estoqueItem) {
          const valorEstoque = obterValorUnitarioEstoqueParaCustoProduto(estoqueItem);
          if (valorEstoque > 0 || !valorUnit) {
            valorUnit = valorEstoque;
          }
        }
      }

      const total = parseNumeroBR(quantidade) * parseNumeroBR(valorUnit);
      custoPrevisto += total;
    });

    const totalNormalizado = parseNumeroBR(custoPrevisto);
    custoProdutoCache[id] = totalNormalizado;
    return totalNormalizado;
  }

  const lista = produtosAtivos.map(produto => {
    const produtoId = String(produto.produto_id || '').trim();
    const receitaPrincipal = receitaPrincipalPorProduto[produtoId] || null;
    const receitaId = String(receitaPrincipal?.receita_id || '').trim();
    const entradasTotal = receitaId && Array.isArray(entradasPorReceita[receitaId])
      ? entradasPorReceita[receitaId].length
      : 0;
    const saidasTotal = receitaId && Array.isArray(saidasPorReceita[receitaId])
      ? saidasPorReceita[receitaId].length
      : 0;
    const etapasTotal = Array.isArray(etapasPorProduto[produtoId])
      ? etapasPorProduto[produtoId].length
      : 0;
    const modeloNome = String(receitaPrincipal?.nome_receita || 'Modelo principal').trim() || 'Modelo principal';

    let custoPrevisto = 0;
    let custoErro = '';
    try {
      custoPrevisto = calcularCustoProdutoListagem(produtoId);
    } catch (err) {
      custoPrevisto = 0;
      custoErro = err && err.message ? err.message : 'Erro ao calcular custo';
    }

    const precoVenda = normalizarPrecoVendaProdutoSaida(produto.preco_venda);
    return {
      ...produto,
      modelo_nome: modeloNome,
      modelo_entradas_total: entradasTotal,
      modelo_saidas_total: saidasTotal,
      modelo_etapas_total: etapasTotal,
      modelo_resumo: `${entradasTotal} entradas, ${saidasTotal} saidas, ${etapasTotal} etapas`,
      preco_venda: precoVenda,
      produto_vendavel: precoVenda !== '',
      custo_previsto: custoPrevisto,
      custo_erro: custoErro,
      criado_em: produto.criado_em
        ? Utilities.formatDate(new Date(produto.criado_em), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
        : ''
    };
  });

  salvarCacheProdutos(lista);
  return lista;
}

function criarProduto(payload) {
  assertCanWriteProdutos('Criacao de produto');
  const nome = String(payload?.nome_produto || '').trim();
  if (!nome) {
    throw new Error('Nome do produto obrigatorio');
  }

  const novo = {
    nome_produto: nome,
    unidade_produto: String(payload?.unidade_produto || 'UN').trim() || 'UN',
    preco_venda: normalizarPrecoVendaProdutoEntrada(payload?.preco_venda),
    produto_id: gerarId('PRD'),
    ativo: true,
    criado_em: new Date()
  };

  insert(ABA_PRODUTOS, novo, PRODUTOS_SCHEMA);
  const precoVenda = normalizarPrecoVendaProdutoSaida(novo.preco_venda);

  return {
    ...novo,
    preco_venda: precoVenda,
    produto_vendavel: precoVenda !== '',
    criado_em: Utilities.formatDate(novo.criado_em, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
  };
}

function obterProduto(produtoId) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS);
  if (!sheet) return null;

  const rows = rowsToObjects(sheet);
  return rows.find(i => i.produto_id === produtoId) || null;
}

function atualizarProduto(produtoId, payload) {
  assertCanWriteProdutos('Atualizacao de produto');
  return updateById(
    ABA_PRODUTOS,
    'produto_id',
    produtoId,
    payload,
    PRODUTOS_SCHEMA
  );
}

function deletarProduto(produtoId) {
  assertCanWriteProdutos('Exclusao de produto');
  return updateById(
    ABA_PRODUTOS,
    'produto_id',
    produtoId,
    { ativo: false },
    PRODUTOS_SCHEMA
  );
}

function listarComponentesProduto(produtoId) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_COMPONENTES);
  if (!sheet) return [];

  const rows = rowsToObjects(sheet);
  return rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.produto_id === produtoId);
}

function listarComposicaoProduto(produtoId) {
  const rows = listarComponentesProduto(produtoId);
  return rows.map(i => ({
    id: i.id,
    produto_id: i.produto_id,
    tipo_item: String(i.tipo_componente || '').toUpperCase(),
    item_id: i.ref_id,
    quantidade: parseNumeroBR(i.quantidade),
    unidade: i.unidade || '',
    observacao: i.observacao || ''
  }));
}

function validarCicloComposicao(produtoId, linhas) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_COMPONENTES);
  const rows = sheet ? rowsToObjects(sheet) : [];

  const ativos = rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.produto_id !== produtoId);

  const novas = (Array.isArray(linhas) ? linhas : []).map(l => ({
    produto_id: produtoId,
    tipo_componente: String(l.tipo_item || l.tipo_componente || '').toUpperCase(),
    ref_id: l.item_id || l.ref_id || ''
  }));

  const all = ativos.concat(novas);
  const mapa = {};

  all.forEach(c => {
    if (c.tipo_componente !== 'PRODUTO') return;
    if (!mapa[c.produto_id]) mapa[c.produto_id] = [];
    mapa[c.produto_id].push(c.ref_id);
  });

  const visitados = {};
  function dfs(id) {
    if (visitados[id]) {
      throw new Error('Loop detectado na composicao do produto');
    }
    visitados[id] = true;
    (mapa[id] || []).forEach(dfs);
    visitados[id] = false;
  }

  dfs(produtoId);
  return true;
}

function criarComponenteProduto(payload) {
  assertCanWriteProdutos('Criacao de componente de produto');
  const novo = {
    ...payload,
    id: gerarId('CMP'),
    ativo: true
  };

  insert(ABA_PRODUTOS_COMPONENTES, novo, PRODUTOS_COMPONENTES_SCHEMA);
  return novo;
}

function atualizarComponenteProduto(id, payload) {
  assertCanWriteProdutos('Atualizacao de componente de produto');
  return updateById(
    ABA_PRODUTOS_COMPONENTES,
    'id',
    id,
    payload,
    PRODUTOS_COMPONENTES_SCHEMA
  );
}

function deletarComponenteProduto(id) {
  assertCanWriteProdutos('Exclusao de componente de produto');
  return updateById(
    ABA_PRODUTOS_COMPONENTES,
    'id',
    id,
    { ativo: false },
    PRODUTOS_COMPONENTES_SCHEMA
  );
}

function salvarComposicaoProduto(produtoId, linhas) {
  assertCanWriteProdutos('Salvamento de composicao de produto');
  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(ABA_PRODUTOS_COMPONENTES);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUTOS_COMPONENTES);
  }

  ensureSchema(sheet, PRODUTOS_COMPONENTES_SCHEMA);

  validarCicloComposicao(produtoId, linhas);

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idCol = headers.indexOf('produto_id');

  if (idCol !== -1 && data.length > 1) {
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][idCol] === produtoId) {
        sheet.deleteRow(i + 1);
      }
    }
  }

  const linhasValidas = Array.isArray(linhas) ? linhas : [];
  linhasValidas.forEach(l => {
    const novo = {
      id: gerarId('CMP'),
      produto_id: produtoId,
      tipo_componente: String(l.tipo_item || l.tipo_componente || '').toUpperCase(),
      ref_id: l.item_id || l.ref_id || '',
      quantidade: parseNumeroBR(l.quantidade),
      unidade: l.unidade || '',
      observacao: l.observacao || '',
      ativo: true
    };
    insert(ABA_PRODUTOS_COMPONENTES, novo, PRODUTOS_COMPONENTES_SCHEMA);
  });

  invalidarCachesRelacionadosAba(ABA_PRODUTOS);
  return true;
}

function adicionarMaterialProduto(produtoId, estoqueId, quantidade) {
  assertCanWriteProdutos('Adicao de material em produto');
  if (!produtoId || !estoqueId) {
    throw new Error('Produto ou estoque invalido');
  }

  const qtd = parseNumeroBR(quantidade);
  if (!qtd || qtd <= 0) {
    throw new Error('Quantidade invalida');
  }

  const sheetEstoque = getDataSpreadsheet().getSheetByName('ESTOQUE');
  if (!sheetEstoque) {
    throw new Error('Aba ESTOQUE nao encontrada');
  }

  const estoqueRows = rowsToObjects(sheetEstoque);
  const estoqueItem = estoqueRows.find(i => i.ID === estoqueId) || {};
  const unidade = estoqueItem.unidade || '';

  const componente = {
    id: gerarId('CMP'),
    produto_id: produtoId,
    tipo_componente: 'ESTOQUE',
    ref_id: estoqueId,
    quantidade: qtd,
    unidade,
    observacao: '',
    ativo: true
  };

  insert(ABA_PRODUTOS_COMPONENTES, componente, PRODUTOS_COMPONENTES_SCHEMA);

  return {
    componente: {
      id: componente.id,
      produto_id: produtoId,
      tipo_item: 'ESTOQUE',
      item_id: estoqueId,
      quantidade: qtd,
      unidade,
      observacao: ''
    },
    custoPrevisto: calcularCustoProduto(produtoId)
  };
}

function listarEtapasProduto(produtoId) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_ETAPAS);
  if (!sheet) return [];

  const rows = rowsToObjects(sheet);
  return rows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.produto_id === produtoId)
    .sort((a, b) => parseNumeroBR(a.ordem) - parseNumeroBR(b.ordem));
}

function criarEtapaProduto(payload) {
  assertCanWriteProdutos('Criacao de etapa de produto');
  const novo = {
    ...payload,
    id: gerarId('ETP'),
    ativo: true
  };

  insert(ABA_PRODUTOS_ETAPAS, novo, PRODUTOS_ETAPAS_SCHEMA);
  return novo;
}

function atualizarEtapaProduto(id, payload) {
  assertCanWriteProdutos('Atualizacao de etapa de produto');
  return updateById(
    ABA_PRODUTOS_ETAPAS,
    'id',
    id,
    payload,
    PRODUTOS_ETAPAS_SCHEMA
  );
}

function deletarEtapaProduto(id) {
  assertCanWriteProdutos('Exclusao de etapa de produto');
  return updateById(
    ABA_PRODUTOS_ETAPAS,
    'id',
    id,
    { ativo: false },
    PRODUTOS_ETAPAS_SCHEMA
  );
}

function limparEtapasProduto(produtoId) {
  assertCanWriteProdutos('Limpeza de etapas de produto');
  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(ABA_PRODUTOS_ETAPAS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUTOS_ETAPAS);
  }

  ensureSchema(sheet, PRODUTOS_ETAPAS_SCHEMA);

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idxProduto = headers.indexOf('produto_id');

  if (idxProduto === -1 || data.length <= 1) return;

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idxProduto] || '') === String(produtoId || '')) {
      sheet.deleteRow(i + 1);
    }
  }
}

function salvarEtapasProduto(produtoId, etapas) {
  assertCanWriteProdutos('Salvamento de etapas de produto');
  if (!produtoId) throw new Error('Produto invalido');

  limparEtapasProduto(produtoId);

  const linhas = Array.isArray(etapas) ? etapas : [];
  let ordemSeq = 1;

  linhas.forEach(et => {
    const nomeEtapa = String(et?.nome_etapa || '').trim();
    if (!nomeEtapa) return;

    const ordemInformada = parseNumeroBR(et?.ordem);
    const ordem = ordemInformada > 0 ? ordemInformada : ordemSeq;

    const novo = {
      id: gerarId('ETP'),
      produto_id: produtoId,
      nome_etapa: nomeEtapa,
      ordem,
      ativo: true
    };

    insert(ABA_PRODUTOS_ETAPAS, novo, PRODUTOS_ETAPAS_SCHEMA);
    ordemSeq += 1;
  });

  invalidarCachesRelacionadosAba(ABA_PRODUTOS);
  return true;
}

function listarReceitasProduto(produtoId) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (!sheet) return [];

  const receitas = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.produto_id === produtoId);

  const sheetEntradas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS_ENTRADAS);
  const sheetSaidas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS_SAIDAS);

  const entradas = sheetEntradas ? rowsToObjects(sheetEntradas) : [];
  const saidas = sheetSaidas ? rowsToObjects(sheetSaidas) : [];

  const entradasAtivas = entradas.filter(i => String(i.ativo).toLowerCase() === 'true');
  const saidasAtivas = saidas
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .map(i => {
      const tipoItem = normalizarTipoSaidaReceitaProduto(i.tipo_item);
      const categoria = normalizarCategoriaSaidaReceitaProduto(tipoItem, i.categoria);
      return {
        ...i,
        produto_ref_id: String(i.produto_ref_id || '').trim(),
        tipo_item: tipoItem,
        categoria
      };
    });

  return receitas.map(r => ({
    ...r,
    criado_em: r.criado_em
      ? Utilities.formatDate(new Date(r.criado_em), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
      : '',
    entradas: entradasAtivas.filter(e => e.receita_id === r.receita_id),
    saidas: saidasAtivas.filter(s => s.receita_id === r.receita_id)
  }));
}

function criarReceitaProduto(produtoId, payload) {
  assertCanWriteProdutos('Criacao de receita de produto');
  if (!produtoId) {
    throw new Error('Produto invalido');
  }

  const receitasAtuais = listarReceitasProduto(produtoId);
  if (receitasAtuais.length > 0) {
    return receitasAtuais[0];
  }

  const novo = {
    receita_id: gerarId('REC'),
    produto_id: produtoId,
    nome_receita: payload?.nome_receita || 'Modelo principal',
    descricao: payload?.descricao || '',
    parent_receita_id: payload?.parent_receita_id || '',
    ativo: true,
    criado_em: new Date()
  };

  insert(ABA_PRODUTOS_RECEITAS, novo, PRODUTOS_RECEITAS_SCHEMA);

  return {
    ...novo,
    criado_em: Utilities.formatDate(novo.criado_em, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
  };
}

function atualizarReceitaProduto(receitaId, payload) {
  assertCanWriteProdutos('Atualizacao de receita de produto');
  return updateById(
    ABA_PRODUTOS_RECEITAS,
    'receita_id',
    receitaId,
    payload,
    PRODUTOS_RECEITAS_SCHEMA
  );
}

function inativarLinhasReceita(sheetName, receitaId) {
  assertCanWriteProdutos('Inativacao de linhas de receita');
  const sheet = getDataSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idCol = headers.indexOf('receita_id');
  const ativoCol = headers.indexOf('ativo');

  if (idCol === -1 || ativoCol === -1) return;

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === receitaId) {
      sheet.getRange(i + 1, ativoCol + 1).setValue(false);
    }
  }
}

function deletarReceitaProduto(receitaId) {
  assertCanWriteProdutos('Exclusao de receita de produto');
  const ok = updateById(
    ABA_PRODUTOS_RECEITAS,
    'receita_id',
    receitaId,
    { ativo: false },
    PRODUTOS_RECEITAS_SCHEMA
  );

  inativarLinhasReceita(ABA_PRODUTOS_RECEITAS_ENTRADAS, receitaId);
  inativarLinhasReceita(ABA_PRODUTOS_RECEITAS_SAIDAS, receitaId);

  invalidarCachesRelacionadosAba(ABA_PRODUTOS);
  return ok;
}

function limparLinhasReceita(sheetName, receitaId) {
  assertCanWriteProdutos('Limpeza de linhas de receita');
  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idCol = headers.indexOf('receita_id');

  if (idCol !== -1 && data.length > 1) {
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][idCol] === receitaId) {
        sheet.deleteRow(i + 1);
      }
    }
  }
}

function salvarEntradasReceita(receitaId, linhas) {
  assertCanWriteProdutos('Salvamento de entradas da receita');
  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(ABA_PRODUTOS_RECEITAS_ENTRADAS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUTOS_RECEITAS_ENTRADAS);
  }

  ensureSchema(sheet, PRODUTOS_RECEITAS_ENTRADAS_SCHEMA);
  limparLinhasReceita(ABA_PRODUTOS_RECEITAS_ENTRADAS, receitaId);

  const linhasValidas = Array.isArray(linhas) ? linhas : [];
  linhasValidas.forEach(l => {
    const tipoItem = String(l.tipo_item || '').toUpperCase();
    const nomeItem = String(l.nome_item || '').trim();
    const comprimentoCm = parseNumeroBR(l.comprimento_cm);
    const larguraCm = parseNumeroBR(l.largura_cm);
    const espessuraCm = parseNumeroBR(l.espessura_cm);
    const volumeM3 = (comprimentoCm > 0 && larguraCm > 0 && espessuraCm > 0)
      ? ((comprimentoCm * larguraCm * espessuraCm) / 1000000)
      : 0;
    const qtdPecas = tipoItem === 'MADEIRA'
      ? volumeM3
      : parseNumeroBR(l.qtd_pecas);

    if (!tipoItem || !nomeItem || !qtdPecas || qtdPecas <= 0) {
      return;
    }

    const novo = {
      id: gerarId('REN'),
      receita_id: receitaId,
      tipo_item: tipoItem,
      nome_item: nomeItem,
      // Modelo nao deve vincular estoque/produto; vinculo real acontece na OP.
      estoque_ref_id: '',
      produto_ref_id: '',
      receita_ref_id: '',
      categoria: l.categoria || '',
      unidade: tipoItem === 'MADEIRA' ? 'M3' : (l.unidade || ''),
      qtd_pecas: qtdPecas,
      comprimento_cm: comprimentoCm,
      largura_cm: larguraCm,
      espessura_cm: espessuraCm,
      custo_manual: parseNumeroBR(l.custo_manual),
      observacao: l.observacao || '',
      ativo: true
    };
    insert(ABA_PRODUTOS_RECEITAS_ENTRADAS, novo, PRODUTOS_RECEITAS_ENTRADAS_SCHEMA);
  });

  return true;
}

function salvarSaidasReceita(receitaId, linhas) {
  assertCanWriteProdutos('Salvamento de saidas da receita');
  const ss = getDataSpreadsheet();
  let sheet = ss.getSheetByName(ABA_PRODUTOS_RECEITAS_SAIDAS);
  if (!sheet) {
    sheet = ss.insertSheet(ABA_PRODUTOS_RECEITAS_SAIDAS);
  }

  ensureSchema(sheet, PRODUTOS_RECEITAS_SAIDAS_SCHEMA);
  limparLinhasReceita(ABA_PRODUTOS_RECEITAS_SAIDAS, receitaId);

  const linhasValidas = Array.isArray(linhas) ? linhas : [];
  const indicesProdutos = listarProdutosAtivosIndicesSaidaReceita();
  const contextoProduto = obterContextoProdutoPorReceitaId(receitaId);
  const totalSaidasProduto = linhasValidas
    .map(l => normalizarTipoSaidaReceitaProduto(l?.tipo_item))
    .filter(tipo => tipo === 'PRODUTO')
    .length;
  linhasValidas.forEach(l => {
    const nomeSaida = String(l.nome_saida || '').trim();
    const unidade = String(l.unidade || '').trim();
    const quantidade = parseNumeroBR(l.quantidade);
    const tipoItem = normalizarTipoSaidaReceitaProduto(l.tipo_item);
    const categoria = normalizarCategoriaSaidaReceitaProduto(tipoItem, l.categoria);
    const produtoRefId = resolverProdutoRefIdSaidaReceita(
      tipoItem,
      nomeSaida,
      l.produto_ref_id,
      indicesProdutos,
      {
        strict: true,
        produtoPadraoId: contextoProduto.produto_id,
        nomeProdutoPadrao: contextoProduto.nome_produto,
        totalSaidasProduto
      }
    );

    if (!nomeSaida || !unidade || quantidade <= 0) {
      return;
    }

    const novo = {
      id: gerarId('RSA'),
      receita_id: receitaId,
      nome_saida: nomeSaida,
      produto_ref_id: produtoRefId,
      tipo_item: tipoItem,
      categoria,
      unidade,
      quantidade,
      ativo: true
    };
    insert(ABA_PRODUTOS_RECEITAS_SAIDAS, novo, PRODUTOS_RECEITAS_SAIDAS_SCHEMA);
  });

  return true;
}

function migrarProdutoRefIdSaidasReceitaPorNome() {
  assertCanWriteProdutos('Migracao de produto_ref_id das saidas de receita');
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS_SAIDAS);
  if (!sheet) {
    return {
      ok: true,
      total_linhas: 0,
      candidatos: 0,
      atualizados: 0
    };
  }

  ensureSchema(sheet, PRODUTOS_RECEITAS_SAIDAS_SCHEMA);
  const data = sheet.getDataRange().getValues();
  if (!Array.isArray(data) || data.length < 2) {
    return {
      ok: true,
      total_linhas: 0,
      candidatos: 0,
      atualizados: 0
    };
  }

  const headers = data[0] || [];
  const idxNome = headers.indexOf('nome_saida');
  const idxTipo = headers.indexOf('tipo_item');
  const idxRef = headers.indexOf('produto_ref_id');
  const idxAtivo = headers.indexOf('ativo');

  if (idxNome < 0 || idxTipo < 0 || idxRef < 0) {
    throw new Error('Colunas obrigatorias nao encontradas em PRODUTOS_RECEITAS_SAIDAS.');
  }

  const indicesProdutos = listarProdutosAtivosIndicesSaidaReceita();
  const sheetReceitas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  const receitasAtivas = sheetReceitas
    ? rowsToObjects(sheetReceitas).filter(r => String(r.ativo).toLowerCase() === 'true')
    : [];
  const produtoPorReceita = {};
  receitasAtivas.forEach(r => {
    const receitaId = String(r.receita_id || '').trim();
    const produtoId = String(r.produto_id || '').trim();
    if (!receitaId || !produtoId) return;
    produtoPorReceita[receitaId] = produtoId;
  });

  const totalSaidasProdutoPorReceita = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i] || [];
    const ativo = idxAtivo >= 0 ? String(row[idxAtivo]).toLowerCase() === 'true' : true;
    if (!ativo) continue;
    const tipo = normalizarTipoSaidaReceitaProduto(row[idxTipo]);
    if (tipo !== 'PRODUTO') continue;
    const receitaId = String(row[headers.indexOf('receita_id')] || '').trim();
    if (!receitaId) continue;
    totalSaidasProdutoPorReceita[receitaId] = Number(totalSaidasProdutoPorReceita[receitaId] || 0) + 1;
  }
  let candidatos = 0;
  let atualizados = 0;
  const idxReceita = headers.indexOf('receita_id');

  for (let i = 1; i < data.length; i++) {
    const row = data[i] || [];
    const ativo = idxAtivo >= 0 ? String(row[idxAtivo]).toLowerCase() === 'true' : true;
    if (!ativo) continue;

    const tipo = normalizarTipoSaidaReceitaProduto(row[idxTipo]);
    if (tipo !== 'PRODUTO') continue;

    const refAtual = String(row[idxRef] || '').trim();
    if (refAtual) continue;

    candidatos += 1;
    const nomeSaida = String(row[idxNome] || '').trim();
    const receitaId = idxReceita >= 0 ? String(row[idxReceita] || '').trim() : '';
    const produtoPadraoId = String(produtoPorReceita[receitaId] || '').trim();
    const nomeProdutoPadrao = String(indicesProdutos.porId?.[produtoPadraoId]?.nome_produto || '').trim();
    const refCalculado = resolverProdutoRefIdSaidaReceita(
      tipo,
      nomeSaida,
      '',
      indicesProdutos,
      {
        strict: false,
        produtoPadraoId,
        nomeProdutoPadrao,
        totalSaidasProduto: Number(totalSaidasProdutoPorReceita[receitaId] || 0)
      }
    );
    if (!refCalculado) continue;

    sheet.getRange(i + 1, idxRef + 1).setValue(refCalculado);
    atualizados += 1;
  }

  if (atualizados > 0) {
    invalidarCachesRelacionadosAba(ABA_PRODUTOS);
  }

  return {
    ok: true,
    total_linhas: data.length - 1,
    candidatos,
    atualizados
  };
}

function duplicarReceitaProduto(receitaId) {
  const _ = receitaId;
  throw new Error('Duplicacao desativada: cada produto possui apenas um modelo');
}

function salvarReceitaCompleta(receitaId, dados, entradas, saidas) {
  assertCanWriteProdutos('Salvamento completo da receita');
  atualizarReceitaProduto(receitaId, dados || {});
  salvarEntradasReceita(receitaId, entradas || []);
  salvarSaidasReceita(receitaId, saidas || []);
  invalidarCachesRelacionadosAba(ABA_PRODUTOS);
  return true;
}

function inativarReceitasSecundariasProduto(produtoId, receitaPrincipalId) {
  assertCanWriteProdutos('Inativacao de receitas secundarias');
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (!sheet) return true;

  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];
  const idxProduto = headers.indexOf('produto_id');
  const idxReceita = headers.indexOf('receita_id');
  const idxAtivo = headers.indexOf('ativo');

  if (idxProduto === -1 || idxReceita === -1 || idxAtivo === -1) return true;

  for (let i = 1; i < data.length; i++) {
    const rowProdutoId = String(data[i][idxProduto] || '');
    const rowReceitaId = String(data[i][idxReceita] || '');
    if (rowProdutoId !== String(produtoId || '')) continue;
    if (rowReceitaId === String(receitaPrincipalId || '')) continue;
    sheet.getRange(i + 1, idxAtivo + 1).setValue(false);
  }

  return true;
}

function obterModeloProduto(produtoId) {
  if (!produtoId) {
    return {
      receita_id: '',
      dados: { nome_receita: 'Modelo principal', descricao: '' },
      entradas: [],
      saidas: [],
      etapas: []
    };
  }

  const receitas = listarReceitasProduto(produtoId);
  const receita = Array.isArray(receitas) && receitas.length > 0 ? receitas[0] : null;
  const etapas = listarEtapasProduto(produtoId);

  if (!receita) {
    return {
      receita_id: '',
      dados: { nome_receita: 'Modelo principal', descricao: '' },
      entradas: [],
      saidas: [],
      etapas
    };
  }

  return {
    receita_id: receita.receita_id || '',
    dados: {
      nome_receita: receita.nome_receita || 'Modelo principal',
      descricao: receita.descricao || ''
    },
    entradas: Array.isArray(receita.entradas) ? receita.entradas : [],
    saidas: Array.isArray(receita.saidas) ? receita.saidas : [],
    etapas: Array.isArray(etapas) ? etapas : []
  };
}

function salvarProdutoComModelo(produtoId, payloadProduto, dadosReceita, entradas, saidas, etapas) {
  assertCanWriteProdutos('Salvamento de produto com modelo');
  const dadosProduto = {
    nome_produto: String(payloadProduto?.nome_produto || '').trim(),
    unidade_produto: String(payloadProduto?.unidade_produto || 'UN').trim() || 'UN',
    preco_venda: normalizarPrecoVendaProdutoEntrada(payloadProduto?.preco_venda)
  };

  if (!dadosProduto.nome_produto) {
    throw new Error('Nome do produto obrigatorio');
  }

  let id = String(produtoId || '').trim();
  if (id) {
    atualizarProduto(id, dadosProduto);
  } else {
    const criado = criarProduto(dadosProduto);
    id = criado.produto_id;
  }

  const modeloDados = {
    nome_receita: String(dadosReceita?.nome_receita || 'Modelo principal').trim() || 'Modelo principal',
    descricao: String(dadosReceita?.descricao || '').trim(),
    parent_receita_id: ''
  };

  const receitasAtuais = listarReceitasProduto(id);
  const receitaPrincipal = Array.isArray(receitasAtuais) && receitasAtuais.length > 0
    ? receitasAtuais[0]
    : criarReceitaProduto(id, modeloDados);

  atualizarReceitaProduto(receitaPrincipal.receita_id, modeloDados);
  salvarEntradasReceita(receitaPrincipal.receita_id, entradas || []);
  salvarSaidasReceita(receitaPrincipal.receita_id, saidas || []);
  inativarReceitasSecundariasProduto(id, receitaPrincipal.receita_id);
  salvarEtapasProduto(id, etapas || []);

  const produtoAtual = obterProduto(id) || {};
  const resumoModelo = getResumoModeloProduto(id);
  const precoVenda = normalizarPrecoVendaProdutoSaida(
    produtoAtual.preco_venda || dadosProduto.preco_venda
  );

  invalidarCachesRelacionadosAba(ABA_PRODUTOS);
  return {
    ...produtoAtual,
    ...resumoModelo,
    produto_id: id,
    nome_produto: produtoAtual.nome_produto || dadosProduto.nome_produto,
    unidade_produto: produtoAtual.unidade_produto || dadosProduto.unidade_produto,
    preco_venda: precoVenda,
    produto_vendavel: precoVenda !== '',
    criado_em: produtoAtual.criado_em
      ? Utilities.formatDate(new Date(produtoAtual.criado_em), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
      : ''
  };
}

function obterReceitaPadraoProduto(produtoId) {
  const sheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (!sheet) return null;

  const receitas = rowsToObjects(sheet)
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.produto_id === produtoId);

  if (receitas.length === 0) return null;
  const semParent = receitas.find(i => !String(i.parent_receita_id || '').trim());
  return semParent || receitas[0] || null;
}

function escolherSaidaPrincipal(saidas, nomeProduto) {
  if (!Array.isArray(saidas) || saidas.length === 0) return null;

  const nome = String(nomeProduto || '').trim().toLowerCase();
  if (nome) {
    const match = saidas.find(s => String(s.nome_saida || '').trim().toLowerCase() === nome);
    if (match) return match;
  }

  let escolhida = saidas[0];
  let maior = parseNumeroBR(escolhida.quantidade);
  for (let i = 1; i < saidas.length; i++) {
    const qtd = parseNumeroBR(saidas[i].quantidade);
    if (qtd > maior) {
      maior = qtd;
      escolhida = saidas[i];
    }
  }
  return escolhida;
}

function agruparItensExplosaoPorEstoque(itens) {
  const resultado = {};
  let custoPrevisto = 0;

  (Array.isArray(itens) ? itens : []).forEach(i => {
    if (!i || !i.estoque_id) return;

    const estoqueId = String(i.estoque_id);
    const quantidade = parseNumeroBR(i.quantidade);
    if (!quantidade || quantidade <= 0) return;

    const valorUnit = parseNumeroBR(i.valor_unit);

    if (!resultado[estoqueId]) {
      resultado[estoqueId] = {
        estoque_id: estoqueId,
        item: i.item || estoqueId,
        unidade: i.unidade || '',
        quantidade: 0,
        valor_unit: valorUnit
      };
    }

    resultado[estoqueId].quantidade += quantidade;
    custoPrevisto += quantidade * valorUnit;
  });

  return {
    itens: Object.values(resultado),
    custoPrevisto
  };
}

function explodirReceitaDetalhada(produtoId, receitaId, qtdPlanejada) {
  if (!produtoId || !receitaId) return { itens: [], custoPrevisto: 0 };

  const sheetReceitas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (!sheetReceitas) return { itens: [], custoPrevisto: 0 };

  const receitas = rowsToObjects(sheetReceitas)
    .filter(i => String(i.ativo).toLowerCase() === 'true');
  const receita = receitas.find(r => r.receita_id === receitaId && r.produto_id === produtoId);
  if (!receita) {
    throw new Error('Receita nao encontrada');
  }

  const sheetEntradas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS_ENTRADAS);
  const entradas = sheetEntradas ? rowsToObjects(sheetEntradas) : [];

  const entradasAtivas = entradas
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.receita_id === receitaId);

  const produtosSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS);
  const produtosRows = produtosSheet ? rowsToObjects(produtosSheet) : [];
  const produtosMap = {};
  produtosRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(p => {
      produtosMap[p.produto_id] = p;
    });

  const qtdPlanejadaNum = parseNumeroBR(qtdPlanejada);
  if (!qtdPlanejadaNum || qtdPlanejadaNum <= 0) {
    return { itens: [], custoPrevisto: 0 };
  }

  // A OP representa quantidade de modelos, entao cada linha da receita multiplica por qtdPlanejada.
  const fator = qtdPlanejadaNum;

  const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
  const estoqueRows = sheetEstoque ? rowsToObjects(sheetEstoque) : [];
  const estoqueMap = {};
  estoqueRows.forEach(i => {
    estoqueMap[i.ID] = i;
  });

  const itensDetalhados = [];
  let custoPrevisto = 0;

  entradasAtivas.forEach(e => {
    const tipo = String(e.tipo_item || '').toUpperCase();
    const prodId = e.produto_ref_id || '';

    let estoqueId = e.estoque_ref_id || '';
    let itemNome = e.nome_item || '';
    let unidade = e.unidade || '';
    let valorUnit = parseNumeroBR(e.custo_manual);
    let quantidade = 0;

    const qtdBase = parseNumeroBR(e.qtd_pecas);
    if (!qtdBase || qtdBase <= 0) return;

    if (tipo === 'PRODUTO' && prodId) {
      const prod = produtosMap[prodId] || null;
      if (prod) {
        itemNome = prod.nome_produto || itemNome || prodId;
        unidade = prod.unidade_produto || unidade || 'UN';
      }
    }

    if (tipo === 'MADEIRA') {
      const comp = parseNumeroBR(e.comprimento_cm);
      const larg = parseNumeroBR(e.largura_cm);
      const esp = parseNumeroBR(e.espessura_cm);
      const volumeM3 = (comp > 0 && larg > 0 && esp > 0)
        ? ((comp * larg * esp) / 1000000)
        : 0;
      const unidadeEntrada = String(e.unidade || '').trim().toUpperCase();
      const qtdInteira = Math.abs(qtdBase - Math.round(qtdBase)) < 0.000001;
      const modoLegadoQtdPecas = unidadeEntrada !== 'M3' && volumeM3 > 0 && qtdInteira;
      quantidade = modoLegadoQtdPecas
        ? (qtdBase * volumeM3 * fator)
        : (qtdBase * fator);
      unidade = unidade || 'M3';
    } else {
      quantidade = qtdBase * fator;
    }

    const estoqueItem = estoqueId ? (estoqueMap[estoqueId] || null) : null;
    if (estoqueItem) {
      itemNome = estoqueItem.item || itemNome || estoqueId;
      unidade = estoqueItem.unidade || unidade || '';
      const valorEstoque = obterValorUnitarioEstoqueParaCustoProduto(estoqueItem);
      if (valorEstoque > 0 || !valorUnit) {
        valorUnit = valorEstoque;
      }
    }

    if (!quantidade || quantidade <= 0) return;
    if (!itemNome) {
      itemNome = prodId || e.nome_item || `Item ${tipo || 'LIVRE'}`;
    }

    const quantidadeNorm = parseNumeroBR(quantidade);
    const valorUnitNorm = parseNumeroBR(valorUnit);
    const total = quantidadeNorm * valorUnitNorm;

    itensDetalhados.push({
      receita_entrada_id: e.id || '',
      receita_id: receitaId,
      estoque_id: estoqueId,
      tipo_item: tipo,
      origem_item: e.nome_item || itemNome,
      item: itemNome,
      unidade,
      quantidade: quantidadeNorm,
      valor_unit: valorUnitNorm,
      total_previsto: total
    });

    custoPrevisto += total;
  });

  return {
    itens: itensDetalhados,
    custoPrevisto
  };
}

function explodirSaidasReceitaDetalhada(produtoId, receitaId, qtdPlanejada) {
  if (!produtoId || !receitaId) return { itens: [] };

  const sheetReceitas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS);
  if (!sheetReceitas) return { itens: [] };

  const receitas = rowsToObjects(sheetReceitas)
    .filter(i => String(i.ativo).toLowerCase() === 'true');
  const receita = receitas.find(r => r.receita_id === receitaId && r.produto_id === produtoId);
  if (!receita) {
    throw new Error('Receita nao encontrada');
  }

  const sheetSaidas = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_RECEITAS_SAIDAS);
  const saidas = sheetSaidas ? rowsToObjects(sheetSaidas) : [];
  const saidasAtivas = saidas
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .filter(i => i.receita_id === receitaId);

  const produtosSheet = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS);
  const produto = produtosSheet
    ? rowsToObjects(produtosSheet).find(i => i.produto_id === produtoId && String(i.ativo).toLowerCase() === 'true')
    : null;
  const nomeProduto = produto ? (produto.nome_produto || '') : '';
  const indicesProdutos = listarProdutosAtivosIndicesSaidaReceita();

  const qtdPlanejadaNum = parseNumeroBR(qtdPlanejada);
  if (!qtdPlanejadaNum || qtdPlanejadaNum <= 0) {
    return { itens: [] };
  }

  // A OP representa quantidade de modelos, entao cada saida multiplica por qtdPlanejada.
  const fator = qtdPlanejadaNum;
  const totalSaidasProduto = saidasAtivas
    .map(s => normalizarTipoSaidaReceitaProduto(s?.tipo_item))
    .filter(tipo => tipo === 'PRODUTO')
    .length;

  const itens = [];
  saidasAtivas.forEach(s => {
    const qtdBase = parseNumeroBR(s.quantidade);
    if (!qtdBase || qtdBase <= 0) return;

    const nomeSaida = String(s.nome_saida || '').trim() || nomeProduto || 'Saida';
    const tipoItem = normalizarTipoSaidaReceitaProduto(s.tipo_item);
    const categoria = normalizarCategoriaSaidaReceitaProduto(tipoItem, s.categoria);
    const produtoRefId = resolverProdutoRefIdSaidaReceita(
      tipoItem,
      nomeSaida,
      s.produto_ref_id,
      indicesProdutos,
      {
        strict: false,
        produtoPadraoId: String(produtoId || '').trim(),
        nomeProdutoPadrao: String(nomeProduto || '').trim(),
        totalSaidasProduto
      }
    );
    const produtoVinculado = produtoRefId
      ? (indicesProdutos.porId?.[produtoRefId] || null)
      : null;
    const quantidade = parseNumeroBR(qtdBase * fator);
    if (!quantidade || quantidade <= 0) return;

    itens.push({
      receita_saida_id: s.id || '',
      receita_id: receitaId,
      nome_saida: nomeSaida,
      produto_ref_id: produtoRefId,
      tipo_item: tipoItem,
      categoria,
      unidade: s.unidade || '',
      quantidade,
      preco_venda_produto: parseNumeroBR(produtoVinculado?.preco_venda)
    });
  });

  return { itens };
}

function explodirReceita(produtoId, receitaId, qtdPlanejada) {
  const detalhada = explodirReceitaDetalhada(produtoId, receitaId, qtdPlanejada);
  const agrupada = agruparItensExplosaoPorEstoque(detalhada.itens || []);

  return {
    itens: agrupada.itens || [],
    custoPrevisto: parseNumeroBR(agrupada.custoPrevisto)
  };
}

function calcularCustoProduto(produtoId) {
  if (!produtoId) return 0;
  const receita = obterReceitaPadraoProduto(produtoId);
  if (!receita) {
    throw new Error('Produto sem receita de producao');
  }
  const resp = explodirReceita(produtoId, receita.receita_id, 1);
  return parseNumeroBR(resp.custoPrevisto);
}

function explodirBOM(produtoId, qtd) {
  const sheetComponentes = getDataSpreadsheet().getSheetByName(ABA_PRODUTOS_COMPONENTES);
  if (!sheetComponentes) return { itens: [], custoPrevisto: 0 };

  const allComponentes = rowsToObjects(sheetComponentes)
    .filter(i => String(i.ativo).toLowerCase() === 'true');

  const componentesPorProduto = {};
  allComponentes.forEach(c => {
    if (!componentesPorProduto[c.produto_id]) {
      componentesPorProduto[c.produto_id] = [];
    }
    componentesPorProduto[c.produto_id].push(c);
  });

  const sheetEstoque = getDataSpreadsheet().getSheetByName(ABA_ESTOQUE);
  const estoqueRows = sheetEstoque ? rowsToObjects(sheetEstoque) : [];
  const estoqueMap = {};
  estoqueRows
    .filter(i => String(i.ativo).toLowerCase() === 'true')
    .forEach(i => {
      estoqueMap[i.ID] = i;
    });

  const resultado = {};
  let custoPrevisto = 0;
  const visitados = {};

  function adicionarEstoque(refId, qtdTotal, unidadeComp) {
    const estoque = estoqueMap[refId] || {};
    const valorUnit = obterValorUnitarioEstoqueParaCustoProduto(estoque);
    const itemNome = estoque.item || refId || '';
    const unidade = estoque.unidade || unidadeComp || '';

    if (!resultado[refId]) {
      resultado[refId] = {
        estoque_id: refId,
        item: itemNome,
        unidade,
        quantidade: 0,
        valor_unit: valorUnit
      };
    }

    resultado[refId].quantidade += qtdTotal;
    custoPrevisto += qtdTotal * valorUnit;
  }

  function walk(prodId, qtdBase) {
    if (!prodId) return;
    if (visitados[prodId]) {
      throw new Error('Loop detectado na composicao do produto');
    }

    visitados[prodId] = true;
    const componentes = componentesPorProduto[prodId] || [];

    componentes.forEach(c => {
      const qtdComp = parseNumeroBR(c.quantidade) * qtdBase;
      if (c.tipo_componente === 'ESTOQUE') {
        adicionarEstoque(c.ref_id, qtdComp, c.unidade);
        return;
      }

      if (c.tipo_componente === 'PRODUTO') {
        walk(c.ref_id, qtdComp);
      }
    });

    visitados[prodId] = false;
  }

  const qtdBase = parseNumeroBR(qtd);
  if (qtdBase > 0) {
    walk(produtoId, qtdBase);
  }

  return {
    itens: Object.values(resultado),
    custoPrevisto
  };
}
