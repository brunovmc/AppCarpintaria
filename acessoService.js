const ABA_USUARIOS_ACESSO = 'USUARIOS';
const USUARIOS_ACESSO_CACHE_SCOPE = 'USUARIOS_ACESSO_MAPA';
const USUARIOS_ACESSO_CACHE_TTL_SEC = 120;

function normalizarTextoAcesso(valor) {
  return String(valor || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function normalizarChaveCabecalhoAcesso(valor) {
  return normalizarTextoAcesso(valor).replace(/[\s_\-]+/g, '');
}

function obterCampoUsuarioAcesso(row, nomeCampo) {
  const alvo = normalizarChaveCabecalhoAcesso(nomeCampo);
  if (!row || typeof row !== 'object' || !alvo) return '';

  const chave = Object.keys(row).find((k) => normalizarChaveCabecalhoAcesso(k) === alvo);
  if (!chave) return '';
  return row[chave];
}

function normalizarEmailAcesso(valor) {
  return String(valor || '').trim().toLowerCase();
}

function parseBooleanAcesso(valor, fallback) {
  if (typeof valor === 'boolean') return valor;

  const normalizado = normalizarTextoAcesso(valor);
  if (!normalizado) return !!fallback;

  if ([
    '1',
    'true',
    'sim',
    's',
    'yes',
    'y',
    'ativo',
    'on'
  ].includes(normalizado)) {
    return true;
  }

  if ([
    '0',
    'false',
    'nao',
    'n',
    'no',
    'inativo',
    'off'
  ].includes(normalizado)) {
    return false;
  }

  return !!fallback;
}

function normalizarRoleAcesso(valor) {
  const role = normalizarTextoAcesso(valor);
  if (!role) return 'viewer';
  if (['admin', 'administrador', 'owner'].includes(role)) return 'admin';
  if ([
    'viewer',
    'leitor',
    'somenteleitura',
    'somente_leitura',
    'readonly',
    'read_only'
  ].includes(role)) {
    return 'viewer';
  }
  return 'viewer';
}

function lerCacheUsuariosAcesso() {
  return appCacheGetJson(USUARIOS_ACESSO_CACHE_SCOPE);
}

function salvarCacheUsuariosAcesso(payload) {
  appCachePutJson(
    USUARIOS_ACESSO_CACHE_SCOPE,
    payload || { configurado: false, mapa: {}, total_usuarios_ativos: 0 },
    USUARIOS_ACESSO_CACHE_TTL_SEC
  );
}

function limparCacheUsuariosAcesso() {
  return appCacheRemove(USUARIOS_ACESSO_CACHE_SCOPE);
}

function carregarUsuariosAcessoConfig(forcarRecarregar) {
  if (!forcarRecarregar) {
    const cached = lerCacheUsuariosAcesso();
    if (cached && typeof cached === 'object' && cached.mapa && typeof cached.mapa === 'object') {
      return cached;
    }
  }

  const sheet = getDataSpreadsheet({ skipAccessCheck: true }).getSheetByName(ABA_USUARIOS_ACESSO);
  if (!sheet) {
    const vazio = { configurado: false, mapa: {}, total_usuarios_ativos: 0 };
    salvarCacheUsuariosAcesso(vazio);
    return vazio;
  }

  const rows = rowsToObjects(sheet);
  const mapa = {};

  rows.forEach((row) => {
    const email = normalizarEmailAcesso(obterCampoUsuarioAcesso(row, 'email'));
    if (!email) return;

    const ativo = parseBooleanAcesso(obterCampoUsuarioAcesso(row, 'ativo'), true);
    if (!ativo) return;

    const role = normalizarRoleAcesso(obterCampoUsuarioAcesso(row, 'role'));
    mapa[email] = {
      email,
      role,
      ativo: true
    };
  });

  const resultado = {
    configurado: true,
    mapa,
    total_usuarios_ativos: Object.keys(mapa).length
  };

  salvarCacheUsuariosAcesso(resultado);
  return resultado;
}

function obterEmailUsuarioAtualAcesso() {
  try {
    return normalizarEmailAcesso(Session.getActiveUser().getEmail());
  } catch (error) {
    return '';
  }
}

function obterContextoUsuario(forcarRecarregar) {
  const config = carregarUsuariosAcessoConfig(!!forcarRecarregar);
  const email = obterEmailUsuarioAtualAcesso();

  if (!config.configurado || Number(config.total_usuarios_ativos || 0) <= 0) {
    return {
      email,
      role: 'admin',
      ativo: true,
      can_read: true,
      can_write: true,
      read_only: false,
      motivo: '',
      configurado: false
    };
  }

  if (!email) {
    return {
      email: '',
      role: 'viewer',
      ativo: false,
      can_read: false,
      can_write: false,
      read_only: true,
      motivo: 'Nao foi possivel identificar seu email no Apps Script.',
      configurado: true
    };
  }

  const usuario = config.mapa[email];
  if (!usuario) {
    return {
      email,
      role: 'viewer',
      ativo: false,
      can_read: false,
      can_write: false,
      read_only: true,
      motivo: 'Usuario nao cadastrado como ativo na aba USUARIOS.',
      configurado: true
    };
  }

  const canWrite = usuario.role === 'admin';
  return {
    email,
    role: usuario.role,
    ativo: true,
    can_read: true,
    can_write: canWrite,
    read_only: !canWrite,
    motivo: canWrite ? '' : 'Perfil somente leitura.',
    configurado: true
  };
}

function assertCanRead(acao) {
  const contexto = String(acao || 'Operacao de leitura').trim() || 'Operacao de leitura';
  const acesso = obterContextoUsuario(false);
  if (acesso.can_read) return true;
  const motivo = String(acesso.motivo || 'Usuario sem permissao de leitura.').trim();
  throw new Error(`${contexto} bloqueada. ${motivo} Entre em contato com o administrador.`);
}

function assertCanWrite(acao) {
  const contexto = String(acao || 'Operacao de escrita').trim() || 'Operacao de escrita';
  const acesso = obterContextoUsuario(false);
  if (acesso.can_write) return true;
  const motivo = String(acesso.motivo || 'Usuario sem permissao de escrita.').trim();
  throw new Error(`${contexto} bloqueada. ${motivo}`);
}

function recarregarCacheUsuariosAcesso() {
  limparCacheUsuariosAcesso();
  const dados = carregarUsuariosAcessoConfig(true);
  return {
    ok: true,
    scope: USUARIOS_ACESSO_CACHE_SCOPE,
    ttl_segundos: USUARIOS_ACESSO_CACHE_TTL_SEC,
    configurado: !!dados.configurado,
    usuarios_ativos: Number(dados.total_usuarios_ativos || 0)
  };
}
