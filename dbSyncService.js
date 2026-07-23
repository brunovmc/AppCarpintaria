const DB_ENV_PROD = 'prod';
const DB_ENV_DEV = 'dev';
const DB_ENV_USER_PROP_KEY = 'APP_DB_ENV_ACTIVE';
let DB_ENV_EXECUTION_OVERRIDE_ = '';

const DB_SYNC_ARMED_ACTION_KEY = 'DB_SYNC_ARMED_ACTION';
const DB_SYNC_ARMED_UNTIL_MS_KEY = 'DB_SYNC_ARMED_UNTIL_MS';
const DB_SYNC_ARMED_BY_KEY = 'DB_SYNC_ARMED_BY';
const DB_SYNC_ARMED_AT_ISO_KEY = 'DB_SYNC_ARMED_AT_ISO';

const DB_SYNC_WRITE_LOCK_UNTIL_MS_KEY = 'DB_SYNC_WRITE_LOCK_UNTIL_MS';
const DB_SYNC_WRITE_LOCK_ACTION_KEY = 'DB_SYNC_WRITE_LOCK_ACTION';
const DB_SYNC_WRITE_LOCK_BY_KEY = 'DB_SYNC_WRITE_LOCK_BY';

const DB_SYNC_ACTION_DEV_TO_PROD = 'DEV_TO_PROD';
const DB_SYNC_ACTION_PROD_TO_DEV = 'PROD_TO_DEV';
const DB_SYNC_ARM_TTL_MS = 10 * 60 * 1000; // 10 minutos
const DB_SYNC_WRITE_LOCK_TTL_MS = 20 * 60 * 1000; // 20 minutos

const OPTIMISTIC_LEDGER_PROPERTY_PREFIX = 'APP_OPTIMISTIC_LEDGER_V2:';
const OPTIMISTIC_LEDGER_TTL_MS = 15 * 24 * 60 * 60 * 1000; // maior que a retencao da outbox
const OPTIMISTIC_LEDGER_MAX_ENTRIES = 1000;

function normalizarAmbienteBancoDados_(valor) {
  const env = String(valor || '').trim().toLowerCase();
  if (env === DB_ENV_DEV) return DB_ENV_DEV;
  return DB_ENV_PROD;
}

function getSpreadsheetIdsConfigurados_() {
  const prod = String(
    (typeof DATA_SPREADSHEET_ID_PROD === 'string' ? DATA_SPREADSHEET_ID_PROD : '')
      || (typeof DATA_SPREADSHEET_ID === 'string' ? DATA_SPREADSHEET_ID : '')
      || ''
  ).trim();
  const dev = String(typeof DATA_SPREADSHEET_ID_DEV === 'string' ? DATA_SPREADSHEET_ID_DEV : '').trim();

  if (!prod) {
    throw new Error('DATA_SPREADSHEET_ID_PROD nao configurado.');
  }

  return { prod, dev };
}

function getSpreadsheetIdPorAmbiente_(ambiente) {
  const ids = getSpreadsheetIdsConfigurados_();
  const env = normalizarAmbienteBancoDados_(ambiente);
  if (env === DB_ENV_DEV) {
    if (!ids.dev) {
      throw new Error('DATA_SPREADSHEET_ID_DEV nao configurado.');
    }
    return ids.dev;
  }
  return ids.prod;
}

function getUserDbEnvProperties_() {
  return PropertiesService.getUserProperties();
}

function getUserDbEnvironment_() {
  try {
    const raw = getUserDbEnvProperties_().getProperty(DB_ENV_USER_PROP_KEY);
    return normalizarAmbienteBancoDados_(raw);
  } catch (error) {
    return DB_ENV_PROD;
  }
}

function getDbEnvironmentExecutionEffective_() {
  return normalizarAmbienteBancoDados_(DB_ENV_EXECUTION_OVERRIDE_ || getUserDbEnvironment_());
}

function setUserDbEnvironment_(ambiente) {
  const env = normalizarAmbienteBancoDados_(ambiente);
  getUserDbEnvProperties_().setProperty(DB_ENV_USER_PROP_KEY, env);
  return env;
}

function getDataSpreadsheetIdAtivo(opcoes) {
  const opts = opcoes || {};
  const envExplicito = String(
    opts.targetEnv || opts.env || DB_ENV_EXECUTION_OVERRIDE_ || ''
  ).trim().toLowerCase();
  if (envExplicito) {
    return getSpreadsheetIdPorAmbiente_(envExplicito);
  }
  return getSpreadsheetIdPorAmbiente_(getDbEnvironmentExecutionEffective_());
}

/**
 * Fixa o ambiente apenas durante a execucao atual. Isso evita que duas abas do
 * mesmo usuario disputem a UserProperty enquanto uma operacao sensivel roda.
 */
function executarComAmbienteBancoDados_(ambiente, callback) {
  if (typeof callback !== 'function') {
    throw new Error('Callback de banco de dados invalido.');
  }
  const anterior = DB_ENV_EXECUTION_OVERRIDE_;
  DB_ENV_EXECUTION_OVERRIDE_ = normalizarAmbienteBancoDados_(ambiente);
  try {
    return callback(DB_ENV_EXECUTION_OVERRIDE_);
  } finally {
    DB_ENV_EXECUTION_OVERRIDE_ = anterior;
  }
}

/**
 * Executa no ambiente solicitado sem alterar a preferencia persistida do usuario
 * e aplica o mesmo controle de acesso usado pelo seletor DEV/PROD.
 */
function executarComAmbienteBancoDadosAutorizado_(ambiente, callback) {
  if (typeof callback !== 'function') {
    throw new Error('Callback de banco de dados invalido.');
  }
  const ambienteInformado = String(ambiente || '').trim().toLowerCase();
  if (ambienteInformado && ambienteInformado !== DB_ENV_PROD && ambienteInformado !== DB_ENV_DEV) {
    throw new Error('Ambiente de banco de dados invalido.');
  }

  const contexto = getDbEnvironmentContext_() || {};
  if (contexto.can_read !== true) {
    throw new Error('Acesso ao banco de dados nao autorizado.');
  }
  const podeAlternar = contexto.can_toggle === true;
  const ambienteExecucao = String(DB_ENV_EXECUTION_OVERRIDE_ || '').trim().toLowerCase();
  const ambienteContexto = normalizarAmbienteBancoDados_(
    ambienteExecucao || contexto.effective_env || contexto.selected_env || DB_ENV_PROD
  );
  if (ambienteInformado === DB_ENV_DEV && !podeAlternar) {
    throw new Error('Ambiente DEV disponivel apenas para administradores.');
  }

  const ambienteAlvo = podeAlternar
    ? normalizarAmbienteBancoDados_(ambienteInformado || ambienteContexto)
    : DB_ENV_PROD;
  return executarComAmbienteBancoDados_(ambienteAlvo, () => callback(ambienteAlvo));
}

function validarAridadeRpcDados_(nome, args, minimo, maximo) {
  const total = Array.isArray(args) ? args.length : -1;
  if (total < minimo || total > maximo) {
    throw new Error(`Quantidade de argumentos invalida para ${nome}.`);
  }
}

/**
 * Dispatcher fechado para RPCs que ainda nao possuem um argumento de ambiente
 * no contrato original. O override vale por toda a arvore de chamadas feita
 * pelo servico escolhido e nunca altera a preferencia persistida do usuario.
 */
function executarRpcDadosNoAmbiente(nomeOperacao, argumentos, ambiente) {
  const nome = String(nomeOperacao || '').trim();
  if (!Array.isArray(argumentos)) {
    throw new Error('Argumentos da operacao de dados invalidos.');
  }
  const args = argumentos;

  return executarComAmbienteBancoDadosAutorizado_(ambiente, () => {
    switch (nome) {
      case 'listarReceitasProduto':
        validarAridadeRpcDados_(nome, args, 1, 1);
        return listarReceitasProduto(args[0]);
      case 'criarReceitaProduto':
        validarAridadeRpcDados_(nome, args, 2, 2);
        return criarReceitaProduto(args[0], args[1]);
      case 'salvarReceitaCompleta':
        validarAridadeRpcDados_(nome, args, 4, 4);
        return salvarReceitaCompleta(args[0], args[1], args[2], args[3]);
      case 'obterModeloProduto':
        validarAridadeRpcDados_(nome, args, 1, 1);
        return obterModeloProduto(args[0]);
      case 'obterMateriaisPrevistosProducao':
        validarAridadeRpcDados_(nome, args, 1, 1);
        return obterMateriaisPrevistosProducao(args[0]);
      case 'listarVinculosMateriaisProducao':
        validarAridadeRpcDados_(nome, args, 1, 2);
        return listarVinculosMateriaisProducao(args[0], args[1]);
      case 'listarSaidasLotesProducao':
        validarAridadeRpcDados_(nome, args, 1, 1);
        return listarSaidasLotesProducao(args[0]);
      case 'atualizarCustoLoteSaidaProducao':
        validarAridadeRpcDados_(nome, args, 3, 3);
        return atualizarCustoLoteSaidaProducao(args[0], args[1], args[2]);
      case 'atualizarPrecoVendaProdutoSaidaProducao':
        validarAridadeRpcDados_(nome, args, 4, 4);
        return atualizarPrecoVendaProdutoSaidaProducao(args[0], args[1], args[2], args[3]);
      case 'salvarReservaEntradaProducao':
        validarAridadeRpcDados_(nome, args, 3, 3);
        return salvarReservaEntradaProducao(args[0], args[1], args[2]);
      case 'removerReservaEntradaProducao':
        validarAridadeRpcDados_(nome, args, 3, 3);
        return removerReservaEntradaProducao(args[0], args[1], args[2]);
      case 'adicionarItemManualProducao':
        validarAridadeRpcDados_(nome, args, 2, 2);
        return adicionarItemManualProducao(args[0], args[1]);
      case 'consumirEstoque':
        validarAridadeRpcDados_(nome, args, 3, 3);
        return consumirEstoque(args[0], args[1], args[2]);
      case 'vincularPendenciasEntradaProducao':
        validarAridadeRpcDados_(nome, args, 2, 2);
        return vincularPendenciasEntradaProducao(args[0], args[1]);
      case 'atualizarEtapasProducao':
        validarAridadeRpcDados_(nome, args, 2, 2);
        return atualizarEtapasProducao(args[0], args[1]);
      case 'deletarEtapaProducao':
        validarAridadeRpcDados_(nome, args, 1, 1);
        return deletarEtapaProducao(args[0]);
      case 'listarEtapasProducao':
        validarAridadeRpcDados_(nome, args, 1, 1);
        return listarEtapasProducao(args[0]);
      case 'baixarEstoqueVenda':
        validarAridadeRpcDados_(nome, args, 1, 1);
        return baixarEstoqueVenda(args[0]);
      default:
        throw new Error('Operacao de dados nao autorizada.');
    }
  });
}

function executarOperacaoOtimistaAutorizada_(nome, args) {
  switch (nome) {
    case 'criarDespesaGeral': return criarDespesaGeral(args[0]);
    case 'atualizarDespesaGeral': return atualizarDespesaGeral(args[0], args[1]);
    case 'deletarDespesaGeral': return deletarDespesaGeral(args[0]);
    case 'registrarPagamentoDespesaGeral': return registrarPagamentoDespesaGeral(args[0], args[1]);
    case 'criarItemCompra': return criarItemCompra(args[0]);
    case 'atualizarItemCompra': return atualizarItemCompra(args[0], args[1]);
    case 'deletarItemCompra': return deletarItemCompra(args[0]);
    case 'registrarPagamentoCompra': return registrarPagamentoCompra(args[0], args[1]);
    case 'adicionarCompraAoEstoque': return adicionarCompraAoEstoque(args[0]);
    case 'criarVenda': return criarVenda(args[0]);
    case 'atualizarVenda': return atualizarVenda(args[0], args[1]);
    case 'deletarVenda': return deletarVenda(args[0]);
    case 'registrarRecebimentoVenda': return registrarRecebimentoVenda(args[0], args[1]);
    case 'criarInvestimento': return criarInvestimento(args[0]);
    case 'atualizarInvestimento': return atualizarInvestimento(args[0], args[1]);
    case 'deletarInvestimento': return deletarInvestimento(args[0]);
    case 'registrarRecebimentoInvestimento': return registrarRecebimentoInvestimento(args[0], args[1]);
    case 'criarItemEstoque': return criarItemEstoque(args[0]);
    case 'atualizarItemEstoque': return atualizarItemEstoque(args[0], args[1]);
    case 'deletarItemEstoque': return deletarItemEstoque(args[0]);
    case 'salvarProdutoComModelo': return salvarProdutoComModelo(
      args[0], args[1], args[2], args[3], args[4], args[5]
    );
    case 'deletarProduto': return deletarProduto(args[0]);
    case 'criarProducao': return criarProducao(args[0]);
    case 'atualizarProducao': return atualizarProducao(args[0], args[1]);
    case 'deletarProducao': return deletarProducao(args[0]);
    default: throw new Error('Operacao otimista nao autorizada.');
  }
}

function chaveLedgerOtimista_(requestId) {
  return `${OPTIMISTIC_LEDGER_PROPERTY_PREFIX}${requestId}`;
}

function lerRegistroLedgerOtimista_(raw) {
  try {
    const item = JSON.parse(String(raw || ''));
    const environment = String(item?.environment || '').trim().toLowerCase();
    const status = String(item?.status || '').trim().toLowerCase();
    const timestamp = Number(item?.completedAt || item?.startedAt || 0);
    if (
      !String(item?.operation || '').trim() ||
      !['prod', 'dev'].includes(environment) ||
      !['started', 'completed'].includes(status) ||
      !Number.isFinite(timestamp) || timestamp <= 0
    ) {
      return null;
    }
    return { ...item, environment, status, timestamp };
  } catch (error) {
    return null;
  }
}

function podarLedgerOtimista_(props, agora, opcoes) {
  const options = opcoes || {};
  const preservar = String(options.preserveKey || '').trim();
  const reservar = Math.max(0, Number(options.reserveSlots || 0));
  const limite = Math.max(1, OPTIMISTIC_LEDGER_MAX_ENTRIES - reservar);
  const todas = props.getProperties() || {};
  const validas = [];

  Object.keys(todas).forEach(chave => {
    if (!chave.startsWith(OPTIMISTIC_LEDGER_PROPERTY_PREFIX)) return;
    const registro = lerRegistroLedgerOtimista_(todas[chave]);
    const expirado = !registro || (agora - registro.timestamp) > OPTIMISTIC_LEDGER_TTL_MS;
    if (expirado) {
      props.deleteProperty(chave);
      return;
    }
    validas.push({ chave, registro });
  });

  validas.sort((a, b) => b.registro.timestamp - a.registro.timestamp);
  const manter = [];
  const preservada = preservar ? validas.find(item => item.chave === preservar) : null;
  if (preservada) manter.push(preservada);
  validas.forEach(item => {
    if (item.chave === preservar || manter.length >= limite) return;
    manter.push(item);
  });
  const chavesMantidas = new Set(manter.map(item => item.chave));
  validas.forEach(item => {
    if (!chavesMantidas.has(item.chave)) props.deleteProperty(item.chave);
  });
}

function iniciarRegistroLedgerOtimista_(props, chave, nome, ambiente, agora) {
  const registro = {
    operation: nome,
    environment: ambiente,
    status: 'started',
    startedAt: agora
  };
  try {
    props.setProperty(chave, JSON.stringify(registro));
  } catch (error) {
    throw new Error('Nao foi possivel preparar a operacao idempotente. Tente novamente.');
  }
  return registro;
}

function concluirRegistroLedgerOtimista_(props, chave, registro) {
  try {
    props.setProperty(chave, JSON.stringify({
      operation: registro.operation,
      environment: registro.environment,
      status: 'completed',
      completedAt: Date.now()
    }));
    return true;
  } catch (error) {
    // A mutacao principal ja terminou. O marcador `started` e mantido para que
    // um eventual retry pare como incerto, em vez de repetir a escrita.
    try { console.warn('Nao foi possivel concluir o marcador de idempotencia.'); } catch (_) { /* sem acao */ }
    return false;
  }
}

/**
 * Executa somente as mutacoes usadas pela outbox do front-end no ambiente em
 * que elas foram criadas. O quarto argumento torna retries do mesmo item
 * idempotentes no servidor, inclusive depois de um reload do navegador.
 */
function executarMutacaoOtimistaNoAmbiente(nomeOperacao, argumentos, ambiente, clientRequestId) {
  const nome = String(nomeOperacao || '').trim();
  const args = Array.isArray(argumentos) ? argumentos : [];
  const requestId = String(clientRequestId || '').trim();
  if (!requestId) {
    throw new Error('Identificador de idempotencia obrigatorio. Recarregue a aplicacao.');
  }
  if (!/^[A-Za-z0-9._:-]{1,160}$/.test(requestId)) {
    throw new Error('Identificador de idempotencia invalido.');
  }

  return executarComAmbienteBancoDadosAutorizado_(ambiente, ambienteEfetivo => {
    const lock = LockService.getUserLock();
    lock.waitLock(30000);
    try {
      const props = PropertiesService.getUserProperties();
      const agora = Date.now();
      const chave = chaveLedgerOtimista_(requestId);
      const rawExistente = props.getProperty(chave);
      if (rawExistente && !lerRegistroLedgerOtimista_(rawExistente)) {
        throw new Error('Registro de idempotencia invalido. Atualize os dados antes de continuar.');
      }
      podarLedgerOtimista_(props, agora, {
        preserveKey: chave,
        reserveSlots: rawExistente ? 0 : 1
      });
      const existente = lerRegistroLedgerOtimista_(props.getProperty(chave));
      if (existente) {
        if (
          String(existente.operation || '') !== nome ||
          normalizarAmbienteBancoDados_(existente.environment) !== ambienteEfetivo
        ) {
          throw new Error('Identificador de idempotencia reutilizado em outra operacao.');
        }
        if (String(existente.status || 'completed') !== 'completed') {
          throw new Error('Uma tentativa anterior ficou com resultado incerto. Atualize os dados antes de continuar.');
        }
        return { ok: true, idempotent_replay: true };
      }

      const inicio = iniciarRegistroLedgerOtimista_(props, chave, nome, ambienteEfetivo, agora);

      let resultado;
      try {
        resultado = executarOperacaoOtimistaAutorizada_(nome, args);
      } catch (error) {
        // Algumas operacoes abrangem varias abas e podem falhar depois de uma
        // escrita parcial. O marcador `started` permanece deliberadamente: o
        // mesmo token nunca e executado outra vez sem reconciliacao manual.
        throw error;
      }
      concluirRegistroLedgerOtimista_(props, chave, inicio);
      return resultado;
    } finally {
      lock.releaseLock();
    }
  });
}

function obterAcessoParaContextoBancoDados_() {
  return (typeof obterContextoUsuario === 'function')
    ? obterContextoUsuario(false)
    : { role: 'denied', can_write: false, can_read: false, email: '' };
}

function getDbEnvironmentContextComAcesso_(acessoInformado) {
  const acesso = acessoInformado && typeof acessoInformado === 'object'
    ? acessoInformado
    : obterAcessoParaContextoBancoDados_();
  const isAdmin = String(acesso?.role || '').toLowerCase() === 'admin' && acesso?.can_write === true;
  const envUsuario = getUserDbEnvironment_();
  const ambienteAtivo = isAdmin ? envUsuario : DB_ENV_PROD;
  const ids = getSpreadsheetIdsConfigurados_();

  return {
    can_read: acesso?.can_read === true,
    can_write: acesso?.can_write === true,
    role: String(acesso?.role || '').trim() || (isAdmin ? 'admin' : 'viewer'),
    email: String(acesso?.email || '').trim(),
    can_toggle: isAdmin,
    selected_env: ambienteAtivo,
    effective_env: ambienteAtivo,
    labels: {
      prod: 'PROD',
      dev: 'DEV'
    },
    ids_configured: {
      prod: !!ids.prod,
      dev: !!ids.dev
    }
  };
}

function getDbEnvironmentContext_() {
  return getDbEnvironmentContextComAcesso_(obterAcessoParaContextoBancoDados_());
}

function obterContextoBancoDados() {
  return getDbEnvironmentContext_();
}

function definirContextoBancoDadosComAcesso_(ambiente, acesso) {
  const ctx = getDbEnvironmentContextComAcesso_(acesso);
  if (!ctx.can_toggle) {
    setUserDbEnvironment_(DB_ENV_PROD);
    return getDbEnvironmentContextComAcesso_(acesso);
  }

  const env = normalizarAmbienteBancoDados_(ambiente);
  if (env === DB_ENV_DEV) {
    const ids = getSpreadsheetIdsConfigurados_();
    if (!ids.dev) {
      throw new Error('Ambiente DEV indisponivel: DATA_SPREADSHEET_ID_DEV nao configurado.');
    }
  }

  setUserDbEnvironment_(env);
  return getDbEnvironmentContextComAcesso_(acesso);
}

function definirContextoBancoDados(ambiente) {
  const acesso = obterAcessoParaContextoBancoDados_();
  return definirContextoBancoDadosComAcesso_(ambiente, acesso);
}

function obterContextoInicialAplicacaoComAcesso_(ambienteSolicitado, acesso) {
  if (!acesso || typeof acesso !== 'object') {
    throw new Error('Contexto de acesso indisponivel para inicializar a aplicacao.');
  }

  if (acesso.can_read === false) {
    return {
      acesso,
      banco_dados: null
    };
  }

  return {
    acesso,
    banco_dados: definirContextoBancoDadosComAcesso_(ambienteSolicitado, acesso)
  };
}

/**
 * Inicializa acesso e ambiente de banco em uma unica chamada do front-end.
 * O ambiente persistido so e alterado depois que a leitura foi autorizada.
 */
function obterContextoInicialAplicacao(ambienteSolicitado, forcarRecarregarAcesso) {
  if (typeof obterContextoUsuario !== 'function') {
    throw new Error('Servico de controle de acesso indisponivel.');
  }

  const acesso = obterContextoUsuario(!!forcarRecarregarAcesso);
  return obterContextoInicialAplicacaoComAcesso_(ambienteSolicitado, acesso);
}

function resetarContextoBancoDadosSessao() {
  setUserDbEnvironment_(DB_ENV_PROD);
  return getDbEnvironmentContext_();
}

function limparChavesSyncArmadas_() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(DB_SYNC_ARMED_ACTION_KEY);
  props.deleteProperty(DB_SYNC_ARMED_UNTIL_MS_KEY);
  props.deleteProperty(DB_SYNC_ARMED_BY_KEY);
  props.deleteProperty(DB_SYNC_ARMED_AT_ISO_KEY);
}

function getArmedSyncStatus_() {
  const props = PropertiesService.getScriptProperties();
  const action = String(props.getProperty(DB_SYNC_ARMED_ACTION_KEY) || '').trim();
  const untilMs = Number(props.getProperty(DB_SYNC_ARMED_UNTIL_MS_KEY) || 0);
  const armedBy = String(props.getProperty(DB_SYNC_ARMED_BY_KEY) || '').trim();
  const armedAtIso = String(props.getProperty(DB_SYNC_ARMED_AT_ISO_KEY) || '').trim();
  const now = Date.now();

  if (!action || !Number.isFinite(untilMs) || untilMs <= now) {
    if (action || untilMs) {
      limparChavesSyncArmadas_();
    }
    return {
      armed: false,
      action: '',
      untilMs: 0,
      armedBy: '',
      armedAtIso: ''
    };
  }

  return {
    armed: true,
    action,
    untilMs,
    armedBy,
    armedAtIso
  };
}

function descreverAcaoSync_(action) {
  if (action === DB_SYNC_ACTION_DEV_TO_PROD) return 'DEV -> PROD';
  if (action === DB_SYNC_ACTION_PROD_TO_DEV) return 'PROD -> DEV';
  return String(action || '').trim() || 'N/A';
}

function assertAdminParaSync_(acao) {
  if (typeof obterContextoUsuario !== 'function') return true;
  const ctx = obterContextoUsuario(false);
  const isAdmin = String(ctx?.role || '').toLowerCase() === 'admin' && ctx?.can_write === true;
  if (isAdmin) return true;
  throw new Error(`${String(acao || 'Operacao').trim() || 'Operacao'} bloqueada. Apenas admin pode executar sync de banco.`);
}

function armarSyncBancoDados_(action) {
  assertAdminParaSync_('Armamento de sync de banco');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const now = Date.now();
    const untilMs = now + DB_SYNC_ARM_TTL_MS;
    const email = (typeof obterEmailUsuarioAtualAcesso === 'function')
      ? String(obterEmailUsuarioAtualAcesso() || '').trim()
      : '';
    const props = PropertiesService.getScriptProperties();
    props.setProperty(DB_SYNC_ARMED_ACTION_KEY, action);
    props.setProperty(DB_SYNC_ARMED_UNTIL_MS_KEY, String(untilMs));
    props.setProperty(DB_SYNC_ARMED_BY_KEY, email || 'admin');
    props.setProperty(DB_SYNC_ARMED_AT_ISO_KEY, new Date(now).toISOString());
    return {
      ok: true,
      armed: true,
      action,
      action_label: descreverAcaoSync_(action),
      armed_by: email || 'admin',
      armed_at: new Date(now),
      expires_at: new Date(untilMs)
    };
  } finally {
    lock.releaseLock();
  }
}

function armSyncDevToProdDB() {
  return armarSyncBancoDados_(DB_SYNC_ACTION_DEV_TO_PROD);
}

function armRevertProdToDevDB() {
  return armarSyncBancoDados_(DB_SYNC_ACTION_PROD_TO_DEV);
}

function limparBloqueioEscritaSyncDB_() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(DB_SYNC_WRITE_LOCK_UNTIL_MS_KEY);
  props.deleteProperty(DB_SYNC_WRITE_LOCK_ACTION_KEY);
  props.deleteProperty(DB_SYNC_WRITE_LOCK_BY_KEY);
}

function setBloqueioEscritaSyncDB_(action) {
  const now = Date.now();
  const untilMs = now + DB_SYNC_WRITE_LOCK_TTL_MS;
  const email = (typeof obterEmailUsuarioAtualAcesso === 'function')
    ? String(obterEmailUsuarioAtualAcesso() || '').trim()
    : '';
  const props = PropertiesService.getScriptProperties();
  props.setProperty(DB_SYNC_WRITE_LOCK_UNTIL_MS_KEY, String(untilMs));
  props.setProperty(DB_SYNC_WRITE_LOCK_ACTION_KEY, String(action || '').trim());
  props.setProperty(DB_SYNC_WRITE_LOCK_BY_KEY, email || 'admin');
  return {
    locked: true,
    action: action,
    by: email || 'admin',
    untilMs
  };
}

function obterStatusBloqueioSyncDB() {
  const props = PropertiesService.getScriptProperties();
  const untilMs = Number(props.getProperty(DB_SYNC_WRITE_LOCK_UNTIL_MS_KEY) || 0);
  const action = String(props.getProperty(DB_SYNC_WRITE_LOCK_ACTION_KEY) || '').trim();
  const by = String(props.getProperty(DB_SYNC_WRITE_LOCK_BY_KEY) || '').trim();
  const now = Date.now();

  if (!Number.isFinite(untilMs) || untilMs <= now) {
    if (untilMs || action || by) {
      limparBloqueioEscritaSyncDB_();
    }
    return {
      blocked: false,
      action: '',
      by: '',
      untilMs: 0
    };
  }

  return {
    blocked: true,
    action,
    by,
    untilMs
  };
}

function assertPodeEscreverNoBancoDados(acao) {
  const status = obterStatusBloqueioSyncDB();
  if (!status.blocked) return true;

  const contexto = String(acao || 'Operacao de escrita').trim() || 'Operacao de escrita';
  const ate = Utilities.formatDate(
    new Date(status.untilMs),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
  const descricaoAcao = descreverAcaoSync_(status.action);
  throw new Error(`${contexto} bloqueada. Sincronizacao manual (${descricaoAcao}) em andamento ate ${ate}.`);
}

function ajustarDimensaoAbaParaSync_(sheet, rowsNeed, colsNeed) {
  const minRows = Math.max(1, Number(rowsNeed) || 1);
  const minCols = Math.max(1, Number(colsNeed) || 1);

  const currentRows = sheet.getMaxRows();
  const currentCols = sheet.getMaxColumns();

  if (currentRows < minRows) {
    sheet.insertRowsAfter(currentRows, minRows - currentRows);
  }
  if (currentCols < minCols) {
    sheet.insertColumnsAfter(currentCols, minCols - currentCols);
  }
}

function copiarDadosPlanilhaCompleta_(sourceSpreadsheetId, targetSpreadsheetId) {
  const sourceId = String(sourceSpreadsheetId || '').trim();
  const targetId = String(targetSpreadsheetId || '').trim();
  if (!sourceId || !targetId) {
    throw new Error('IDs de origem/destino invalidos para sincronizacao.');
  }
  if (sourceId === targetId) {
    throw new Error('Sync invalido: origem e destino apontam para a mesma planilha.');
  }

  const source = SpreadsheetApp.openById(sourceId);
  const target = SpreadsheetApp.openById(targetId);

  const sourceSheets = source.getSheets();
  if (!Array.isArray(sourceSheets) || sourceSheets.length === 0) {
    throw new Error('Planilha de origem sem abas para sincronizar.');
  }

  const resumoAbas = [];
  const copiasTemporarias = [];
  const marker = `__SYNC_TMP_${Date.now()}_${Math.floor(Math.random() * 100000)}_`;

  sourceSheets.forEach((sourceSheet, idx) => {
    const nomeFinal = sourceSheet.getName();
    const linhaMax = sourceSheet.getLastRow();
    const colunaMax = sourceSheet.getLastColumn();
    const clone = sourceSheet.copyTo(target);
    const nomeTemporario = `${marker}${idx + 1}`;
    clone.setName(nomeTemporario);

    copiasTemporarias.push({
      sheet: clone,
      nomeTemporario,
      nomeFinal,
      ordem: idx + 1
    });

    resumoAbas.push({
      aba: nomeFinal,
      linhas: linhaMax,
      colunas: colunaMax
    });
  });

  const idsCopiados = {};
  copiasTemporarias.forEach(item => {
    idsCopiados[item.sheet.getSheetId()] = true;
  });

  const abasRemovidas = [];
  const abasAtuaisDestino = target.getSheets();
  abasAtuaisDestino.forEach(sheet => {
    if (idsCopiados[sheet.getSheetId()]) return;
    abasRemovidas.push(sheet.getName());
    if (target.getSheets().length <= 1) return;
    target.deleteSheet(sheet);
  });

  copiasTemporarias.forEach(item => {
    item.sheet.setName(item.nomeFinal);
    target.setActiveSheet(item.sheet);
    target.moveActiveSheet(item.ordem);
  });

  SpreadsheetApp.flush();

  return {
    total_abas_copiadas: resumoAbas.length,
    abas: resumoAbas,
    abas_removidas: abasRemovidas
  };
}

function limparCachesAposSyncBancoDados_() {
  const envOriginal = getUserDbEnvironment_();
  const ids = getSpreadsheetIdsConfigurados_();
  const ambientes = [DB_ENV_PROD];
  if (ids.dev) ambientes.push(DB_ENV_DEV);

  try {
    ambientes.forEach(env => {
      setUserDbEnvironment_(env);
      if (typeof limparCacheValidacoes === 'function') limparCacheValidacoes();
      if (typeof limparCacheEstoque === 'function') limparCacheEstoque();
      if (typeof limparCacheCompras === 'function') limparCacheCompras();
      if (typeof limparCacheVendas === 'function') limparCacheVendas();
      if (typeof limparCacheInvestimentos === 'function') limparCacheInvestimentos();
      if (typeof limparCacheProdutos === 'function') limparCacheProdutos();
      if (typeof limparCacheProducao === 'function') limparCacheProducao();
      if (typeof limparCacheDespesasGerais === 'function') limparCacheDespesasGerais();
      if (typeof limparCachePagamentos === 'function') limparCachePagamentos();
      if (typeof limparCacheUsuariosAcesso === 'function') limparCacheUsuariosAcesso();
      if (typeof limparCacheDashboardFinanceiro === 'function') limparCacheDashboardFinanceiro();
    });
  } finally {
    setUserDbEnvironment_(envOriginal);
  }
}

function executarSyncBancoDadosArmado_(expectedAction) {
  assertAdminParaSync_('Sync de banco');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const armed = getArmedSyncStatus_();
    if (!armed.armed) {
      throw new Error(`Sync nao autorizado. Execute arm primeiro para ${descreverAcaoSync_(expectedAction)}.`);
    }
    if (armed.action !== expectedAction) {
      throw new Error(`Sync armado para ${descreverAcaoSync_(armed.action)}, mas solicitado ${descreverAcaoSync_(expectedAction)}.`);
    }

    setBloqueioEscritaSyncDB_(expectedAction);

    try {
      const sourceEnv = expectedAction === DB_SYNC_ACTION_DEV_TO_PROD ? DB_ENV_DEV : DB_ENV_PROD;
      const targetEnv = expectedAction === DB_SYNC_ACTION_DEV_TO_PROD ? DB_ENV_PROD : DB_ENV_DEV;
      const sourceId = getSpreadsheetIdPorAmbiente_(sourceEnv);
      const targetId = getSpreadsheetIdPorAmbiente_(targetEnv);
      const startedAt = new Date();

      const copyResult = copiarDadosPlanilhaCompleta_(sourceId, targetId);
      limparCachesAposSyncBancoDados_();

      return {
        ok: true,
        action: expectedAction,
        action_label: descreverAcaoSync_(expectedAction),
        source_env: sourceEnv,
        target_env: targetEnv,
        source_id: sourceId,
        target_id: targetId,
        started_at: startedAt,
        finished_at: new Date(),
        resumo: copyResult
      };
    } finally {
      limparBloqueioEscritaSyncDB_();
      limparChavesSyncArmadas_();
    }
  } finally {
    lock.releaseLock();
  }
}

function syncDevToProdDB() {
  return executarSyncBancoDadosArmado_(DB_SYNC_ACTION_DEV_TO_PROD);
}

function revertProdToDevDB() {
  return executarSyncBancoDadosArmado_(DB_SYNC_ACTION_PROD_TO_DEV);
}
