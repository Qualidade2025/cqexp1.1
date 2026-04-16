/**
 * Validação central do payload do formulário.
 */
function validateInspectionPayload_(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Payload inválido.');
  }

  var op = String(payload.op || '').trim();
  if (op.toLowerCase() === 'sem op') {
    op = '';
  }
  var opLocked = !!payload.opLocked;
  var opIsRequired = isOpRequired_();

  if (opLocked) {
    if (opIsRequired) {
      throw new Error('OP é obrigatória.');
    }
  } else {
    if (!op && opIsRequired) {
      throw new Error('OP é obrigatória.');
    }

    if (op && !/^\d+$/.test(op)) {
      throw new Error('OP deve conter apenas dígitos numéricos.');
    }

    if (op && !getClientByOP(op)) {
      throw new Error('OP ' + op + ' não foi encontrada na aba op_cliente.');
    }
  }

  if (!payload.clienteManual || !String(payload.clienteManual).trim()) {
    throw new Error('Cliente é obrigatório.');
  }

  var qtd = Number(payload.qtddRevisada);
  if (!qtd || qtd <= 0) {
    throw new Error('Qtdd revisada deve ser maior que zero.');
  }

  var originIsRequired = isOriginRequired_();
  if (originIsRequired && (!payload.origem || !String(payload.origem).trim())) {
    throw new Error('Origem é obrigatória.');
  }

  if (!Array.isArray(payload.operadores) || payload.operadores.length < 1) {
    throw new Error('Selecione ao menos 1 operador para a inspeção.');
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

  var defects = normalizePayloadDefects_(payload.defeitos);
  if (defects.hasPartialRows) {
    throw new Error('Preencha posição, defeito e quantidade em cada lançamento iniciado.');
  }

  var defectCatalog = getDefectsByPositionCatalog_();
  var activePairs = defectCatalog.paresAtivos || {};

  defects.items.forEach(function (item, idx) {
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

    if (quantidade > qtd) {
      throw new Error('Linha ' + (idx + 1) + ': quantidade não pode ser maior que Qtdd Revisada (' + qtd + ').');
    }

    var relationKey = getPositionDefectKey_(item.posicao, item.defeito);
    if (!activePairs[relationKey]) {
      throw new Error('Linha ' + (idx + 1) + ': defeito "' + item.defeito + '" não está ativo para posição "' + item.posicao + '".');
    }
  });
}

function normalizePayloadDefects_(rawDefects) {
  if (!Array.isArray(rawDefects) || !rawDefects.length) {
    return {
      items: [],
      hasPartialRows: false
    };
  }

  var hasPartialRows = false;
  var validItems = [];

  rawDefects.forEach(function (item) {
    var posicao = item && item.posicao ? String(item.posicao).trim() : '';
    var defeito = item && item.defeito ? String(item.defeito).trim() : '';
    var quantidadeRaw = item ? item.quantidade : '';
    var quantidade = Number(quantidadeRaw);
    var quantidadeInformada = String(quantidadeRaw === 0 ? '0' : (quantidadeRaw || '')).trim() !== '';

    var hasAnyField = !!(posicao || defeito || quantidadeInformada);
    var isComplete = !!(posicao && defeito && quantidadeInformada);

    if (!hasAnyField) {
      return;
    }

    if (!isComplete) {
      hasPartialRows = true;
      return;
    }

    validItems.push({
      posicao: posicao,
      defeito: defeito,
      quantidade: quantidade
    });
  });

  return {
    items: validItems,
    hasPartialRows: hasPartialRows
  };
}
