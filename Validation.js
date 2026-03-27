/**
 * Validação central do payload do formulário.
 */
function validateInspectionPayload_(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Payload inválido.');
  }

  if (!payload.op || !String(payload.op).trim()) {
    throw new Error('OP é obrigatória.');
  }

  var qtd = Number(payload.qtddRevisada);
  if (!qtd || qtd <= 0) {
    throw new Error('Qtdd revisada deve ser maior que zero.');
  }

  if (!payload.origem || !String(payload.origem).trim()) {
    throw new Error('Origem é obrigatória.');
  }

  if (!Array.isArray(payload.operadores) || payload.operadores.length < 1 || payload.operadores.length > 3) {
    throw new Error('Selecione de 1 a 3 colaboradores na equipe.');
  }

  var duplicateCheck = {};
  payload.operadores.forEach(function (op) {
    if (!op || !op.id || !op.name) {
      throw new Error('Colaborador inválido na equipe.');
    }
    if (duplicateCheck[op.id]) {
      throw new Error('Não repita colaboradores na equipe.');
    }
    duplicateCheck[op.id] = true;
  });

  if (!Array.isArray(payload.defeitos) || payload.defeitos.length < 1) {
    throw new Error('Inclua ao menos 1 lançamento de defeito.');
  }

  payload.defeitos.forEach(function (item, idx) {
    if (!item.posicao || !String(item.posicao).trim()) {
      throw new Error('Linha ' + (idx + 1) + ': posição é obrigatória.');
    }
    if (!item.defeito || !String(item.defeito).trim()) {
      throw new Error('Linha ' + (idx + 1) + ': defeito é obrigatório.');
    }

    var quantidade = Number(item.quantidade);
    if (!quantidade || quantidade <= 0) {
      throw new Error('Linha ' + (idx + 1) + ': quantidade deve ser maior que zero.');
    }
  });
}