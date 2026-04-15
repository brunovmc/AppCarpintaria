function gerarId(prefixo) {
  const props = PropertiesService.getScriptProperties();
  const chave = `SEQ_${prefixo}`;
  const lock = LockService.getScriptLock();

  lock.waitLock(10000);
  try {
    let seq = Number(props.getProperty(chave)) || 0;
    seq++;

    props.setProperty(chave, String(seq));

    return `${prefixo}-${String(seq).padStart(4, '0')}`;
  } finally {
    lock.releaseLock();
  }
}

function agora() {
  return new Date();
}

function normalizarTextoNumero(valor) {
  let texto = String(valor ?? '').trim();
  if (texto === '') return '';

  // Remove simbolos comuns mantendo apenas digitos, separadores e sinal.
  texto = texto.replace(/\s+/g, '');
  texto = texto.replace(/[^\d,.\-]/g, '');
  if (!texto || texto === '-') return '';

  const negativo = texto.startsWith('-');
  texto = texto.replace(/-/g, '');
  if (!texto) return '';

  const temVirgula = texto.includes(',');
  const temPonto = texto.includes('.');
  let normalizado = texto;

  if (temVirgula && temPonto) {
    // O ultimo separador encontrado define a parte decimal.
    if (texto.lastIndexOf(',') > texto.lastIndexOf('.')) {
      normalizado = texto.replace(/\./g, '').replace(',', '.');
    } else {
      normalizado = texto.replace(/,/g, '');
    }
  } else if (temVirgula) {
    const partes = texto.split(',');
    normalizado = partes.length > 2
      ? partes.join('')
      : `${partes[0] || '0'}.${partes[1] || '0'}`;
  } else if (temPonto) {
    const partes = texto.split('.');
    normalizado = partes.length > 2
      ? partes.join('')
      : texto;
  }

  return `${negativo ? '-' : ''}${normalizado}`;
}

function parseNumeroBR(valor) {
  if (valor === null || valor === undefined) return 0;
  if (typeof valor === 'number') return isFinite(valor) ? valor : 0;

  const normalizado = normalizarTextoNumero(valor);
  if (!normalizado) return 0;

  const n = Number(normalizado);
  return isNaN(n) ? 0 : n;
}
