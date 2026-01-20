function gerarId(prefixo) {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();
  const chave = `SEQ_${prefixo}`;

  let seq = Number(props.getProperty(chave)) || 0;
  seq++;

  props.setProperty(chave, seq);

  return `${prefixo}-${String(seq).padStart(4, '0')}`;
}

function agora() {
  return new Date();
}

function parseNumeroBR(valor) {
  if (valor === null || valor === undefined) return 0;
  if (typeof valor === 'number') return valor;

  const s = String(valor).trim();
  if (s === '') return 0;

  const normalizado = s.replace(/\./g, '').replace(',', '.');
  const n = Number(normalizado);
  return isNaN(n) ? 0 : n;
}
