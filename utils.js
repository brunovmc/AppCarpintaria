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
