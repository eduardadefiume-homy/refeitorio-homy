// ============================================================
// refeitorio-regras.js — Camada central de regras do Refeitório Homy
// v: base-centralizada-v10-17-20260701
//
// Objetivo:
// - Centralizar as regras de produção, ausência, extras, cozinha e cardápio do dia.
// - Não acessa SharePoint diretamente.
// - Não grava dados.
// - Pode ser usado por Admin, Cozinha, Cardápio do Dia e testes.
// ============================================================
(function (global) {
  "use strict";

  const Regras = {};

  // ============================================================
  // NORMALIZAÇÃO / HELPERS
  // ============================================================
  Regras.norm = function norm(value) {
    return String(value ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .trim();
  };

  Regras.pick = function pick(obj, ...keys) {
    for (const key of keys) {
      if (obj && obj[key] !== undefined && obj[key] !== null) return obj[key];
    }
    return "";
  };

  Regras.isTrue = function isTrue(value) {
    if (value === true || value === 1) return true;
    const v = Regras.norm(value);
    return v === "sim" || v === "true" || v === "yes" || v === "1";
  };

  Regras.dataISO = function dataISO(value) {
    if (!value) return "";
    if (value instanceof Date && !isNaN(value)) {
      const y = value.getFullYear();
      const m = String(value.getMonth() + 1).padStart(2, "0");
      const d = String(value.getDate()).padStart(2, "0");
      return `${y}-${m}-${d}`;
    }
    const s = String(value).trim();
    const iso = s.match(/^(\d{4}-\d{2}-\d{2})/);
    if (iso) return iso[1];
    const br = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (br) return `${br[3]}-${br[2]}-${br[1]}`;
    const d = new Date(s);
    if (isNaN(d)) return "";
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const dia = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${dia}`;
  };

  Regras.getWeekDates = function getWeekDates(semanaId) {
    const match = String(semanaId || "").match(/^(\d{4})-W(\d{2})$/);
    if (!match) return [];
    const year = Number(match[1]);
    const week = Number(match[2]);
    const jan4 = new Date(year, 0, 4);
    const start = new Date(jan4);
    start.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1 + (week - 1) * 7);
    return Array.from({ length: 5 }, (_, i) => {
      const d = new Date(start);
      d.setDate(start.getDate() + i);
      return d;
    });
  };

  Regras.dataPorSemanaDia = function dataPorSemanaDia(semanaId, dia, helpers = {}) {
    try {
      if (helpers && typeof helpers.getDataRefBySemanaDia === "function") {
        const data = helpers.getDataRefBySemanaDia(semanaId, dia);
        if (data) return Regras.dataISO(data);
      }
    } catch (_) {}

    const idxPorDia = {
      segunda: 0,
      terca: 1,
      "terça": 1,
      quarta: 2,
      quinta: 3,
      sexta: 4
    };
    const idx = idxPorDia[Regras.norm(dia)];
    const datas = Regras.getWeekDates(semanaId);
    if (idx === undefined || !datas[idx]) return "";
    return Regras.dataISO(datas[idx]);
  };

  // ============================================================
  // CAMPOS OFICIAIS / CHAVES
  // ============================================================
  Regras.getPedidoId = function getPedidoId(p) {
    return String(Regras.pick(p, "id", "ID") || "").trim();
  };

  Regras.getColaboradorId = function getColaboradorId(obj) {
    return String(Regras.pick(
      obj,
      "Colaborador_id", "colaborador_id", "ColaboradorId", "colaboradorId", "Matricula", "Matrícula", "id", "ID"
    ) || "").trim();
  };

  Regras.getNome = function getNome(obj) {
    return Regras.pick(
      obj,
      "Colaborador_nome", "colaborador_nome", "Colaborador", "Nome", "Title", "nome"
    ) || "";
  };

  Regras.getDia = function getDia(obj) {
    return Regras.pick(obj, "Dia", "dia") || "";
  };

  Regras.getOpcao = function getOpcao(obj) {
    return Regras.pick(obj, "Opcao", "opcao", "Opção") || "";
  };

  Regras.getStatus = function getStatus(obj) {
    return Regras.pick(obj, "Status", "status") || "";
  };

  Regras.getOrigem = function getOrigem(obj) {
    return Regras.pick(obj, "Origem", "origem", "tipo", "Tipo") || "";
  };

  Regras.getCentroCusto = function getCentroCusto(obj) {
    return Regras.pick(
      obj,
      "Centro_Custo", "CentroCusto", "Centro Custo", "Centro de Custo", "Departamento", "departamento", "Setor", "setor"
    ) || "";
  };

  Regras.colaboradorKey = function colaboradorKey(obj) {
    const id = Regras.getColaboradorId(obj);
    if (id) return `id:${id}`;
    const nome = Regras.norm(Regras.getNome(obj));
    return nome ? `nome:${nome}` : "";
  };

  // Chaves equivalentes para reconciliar colaboradores entre listas diferentes.
  // Importante: em Pedido, o campo id/ID é o ID do item do SharePoint, não o ID do colaborador.
  // Por isso pedido usa Colaborador_id; colaborador cadastrado usa id/ID e também nome.
  Regras.chavesColaboradorOperacional = function chavesColaboradorOperacional(obj, tipo = "auto") {
    const chaves = [];
    const isPedido = tipo === "pedido" || Regras.pick(obj, "Semana_id", "Dia", "Opcao", "Nome_Prato") !== "";

    const id = isPedido
      ? String(Regras.pick(obj, "Colaborador_id", "colaborador_id", "ColaboradorId", "colaboradorId", "Matricula", "Matrícula") || "").trim()
      : String(Regras.pick(obj, "id", "ID", "Colaborador_id", "colaborador_id", "ColaboradorId", "colaboradorId", "Matricula", "Matrícula") || "").trim();

    const nome = Regras.norm(Regras.getNome(obj));
    if (id) chaves.push(`id:${id}`);
    if (nome) chaves.push(`nome:${nome}`);
    return Array.from(new Set(chaves));
  };

  Regras.mesmoColaboradorOperacional = function mesmoColaboradorOperacional(a, b, tipoA = "auto", tipoB = "auto") {
    const ka = Regras.chavesColaboradorOperacional(a, tipoA);
    const kb = new Set(Regras.chavesColaboradorOperacional(b, tipoB));
    return ka.some(k => kb.has(k));
  };

  // ============================================================
  // STATUS / PRODUÇÃO
  // ============================================================
  Regras.STATUS_PRODUCAO = [
    "confirmado",
    "aprovado",
    "extra"
  ];

  Regras.STATUS_BLOQUEIA_PRODUCAO = [
    "cancelado",
    "bloqueado",
    "nao vai almocar",
    "nao_vai_almocar",
    "não vai almoçar",
    "ausente",
    "ferias",
    "férias",
    "afastado",
    "afastamento",
    "atestado",
    "licenca",
    "licença",
    "banco horas",
    "banco_horas",
    "homy office",
    "homy_office",
    "falta"
  ];

  Regras.STATUS_AUSENCIA = [
    "nao vai almocar",
    "nao_vai_almocar",
    "não vai almoçar",
    "ausente",
    "ferias",
    "férias",
    "afastado",
    "afastamento",
    "atestado",
    "licenca",
    "licença",
    "banco horas",
    "banco_horas",
    "homy office",
    "homy_office",
    "falta"
  ];

  Regras.statusBloqueiaProducao = function statusBloqueiaProducao(status) {
    return Regras.STATUS_BLOQUEIA_PRODUCAO.includes(Regras.norm(status));
  };

  Regras.statusAusencia = function statusAusencia(status) {
    return Regras.STATUS_AUSENCIA.includes(Regras.norm(status));
  };

  Regras.pedidoStatusBloqueado = function pedidoStatusBloqueado(pedido) {
    return Regras.statusBloqueiaProducao(Regras.getStatus(pedido));
  };

  Regras.pedidoAusente = function pedidoAusente(pedido) {
    const status = Regras.getStatus(pedido);
    const origem = Regras.norm(Regras.getOrigem(pedido));
    return Regras.statusAusencia(status) || origem.includes("ausencia") || origem.includes("ausência") || !!pedido?._ausente;
  };


  Regras.isRetornoAutomaticoAusencia = function isRetornoAutomaticoAusencia(pedido) {
    const origem = Regras.norm(Regras.getOrigem(pedido));
    return origem.includes("retorno automatico de ausencia") ||
           origem.includes("retorno automático de ausência") ||
           (origem.includes("retorno automatico") && origem.includes("ausencia")) ||
           (origem.includes("retorno automático") && origem.includes("ausência"));
  };

  Regras.isTravamentoAutomatico = function isTravamentoAutomatico(pedido) {
    const status = Regras.norm(Regras.getStatus(pedido));
    const origem = Regras.norm(Regras.getOrigem(pedido));
    return status === "travado" && origem.includes("travamento");
  };

  Regras.pedidoCanceladoOuBloqueado = function pedidoCanceladoOuBloqueado(pedido) {
    const status = Regras.norm(Regras.getStatus(pedido));
    const origem = Regras.norm(Regras.getOrigem(pedido));
    return ["cancelado", "bloqueado", "duplicado inativado", "duplicidade inativada"].includes(status) ||
           origem.includes("duplicado inativado") || origem.includes("duplicidade inativada");
  };

  Regras.pedidoEscolhaReal = function pedidoEscolhaReal(pedido) {
    if (!pedido) return false;

    // Registro legado gerado pela regra antiga. Ele existe no SharePoint,
    // mas não é escolha do colaborador nem travamento oficial.
    if (Regras.isRetornoAutomaticoAusencia(pedido)) return false;

    // Travamento oficial é a única forma de Principal automático após prazo.
    if (Regras.isTravamentoAutomatico(pedido)) return true;

    const status = Regras.norm(Regras.getStatus(pedido));
    if (status === "travado") return false;
    if (Regras.statusBloqueiaProducao(status)) return false;

    return Regras.STATUS_PRODUCAO.includes(status) || Regras.isTrue(Regras.pick(pedido, "Confirmado", "confirmado"));
  };

  Regras.pedidoPodeAparecerNaOperacao = function pedidoPodeAparecerNaOperacao(pedido) {
    if (!pedido) return false;
    if (Regras.isRetornoAutomaticoAusencia(pedido)) return false;
    if (Regras.pedidoCanceladoOuBloqueado(pedido)) return false;
    return true;
  };

  Regras.criarPendenteAusenciaEncerrada = function criarPendenteAusenciaEncerrada(pedidoBase, overrides = {}) {
    const base = pedidoBase || {};
    return {
      ...base,
      ...overrides,
      _virtualPendente: true,
      _legadoAusenciaEncerrada: true,
      _retornoAutomaticoLegado: Regras.isRetornoAutomaticoAusencia(base),
      Status: "Pendente",
      Confirmado: false,
      Opcao: Regras.getOpcao(base) || "principal",
      Nome_Prato: "Sem marcação",
      Origem: "Sem pedido",
      Observacao: "Ausência encerrada ou retorno automático legado ignorado na Operação viva. Colaborador elegível deve aparecer como pendente normal; Principal só será aplicado no travamento oficial."
    };
  };

  Regras.pedidoConfirmadoProducao = function pedidoConfirmadoProducao(pedido) {
    const status = Regras.norm(Regras.getStatus(pedido));

    // Retorno automático de ausência é dado legado e não conta produção.
    if (Regras.isRetornoAutomaticoAusencia(pedido)) return false;

    // Travado só conta quando é travamento oficial.
    if (Regras.isTravamentoAutomatico(pedido)) return true;
    if (status === "travado") return false;

    // Status de ausência/cancelamento vence Confirmado=true.
    if (Regras.statusBloqueiaProducao(status)) return false;

    if (Regras.STATUS_PRODUCAO.includes(status)) return true;
    return Regras.isTrue(Regras.pick(pedido, "Confirmado", "confirmado"));
  };

  Regras.motivoExclusaoProducao = function motivoExclusaoProducao(pedido) {
    const status = Regras.getStatus(pedido);
    if (Regras.statusBloqueiaProducao(status)) return `Status bloqueia produção: ${status || "sem status"}`;
    if (!Regras.pedidoConfirmadoProducao(pedido)) return "Não confirmado para produção";
    return "";
  };

  // ============================================================
  // TIPOS / EXTRAS / COZINHA / CARDÁPIO DO DIA
  // ============================================================
  Regras.isPedidoAdicionalColaborador = function isPedidoAdicionalColaborador(pedido) {
    const origem = Regras.norm(Regras.getOrigem(pedido));
    const obs = Regras.norm(Regras.pick(pedido, "Observacao", "Observação", "observacao", "Obs"));
    const cid = Regras.norm(Regras.getColaboradorId(pedido));
    return origem.includes("segunda refeicao") || origem.includes("segunda refeição") ||
           origem.includes("transferencia para hoje") || origem.includes("transferência para hoje") ||
           origem.includes("refeicao adicional") || origem.includes("refeição adicional") ||
           obs.includes("adicionalid:") || obs.includes("colaboradorbaseid:") ||
           cid.includes("-adicional-");
  };

  Regras.isGuarda = function isGuarda(obj) {
    const nome = Regras.norm(Regras.getNome(obj));
    const origem = Regras.norm(Regras.getOrigem(obj));
    return nome === "guarda" || /^guarda\s*\d*$/.test(nome) || origem.includes("guarda");
  };

  Regras.isInvestigador = function isInvestigador(obj) {
    const nome = Regras.norm(Regras.getNome(obj));
    const origem = Regras.norm(Regras.getOrigem(obj));
    return /^investigador(?:\s*\d+)?$/.test(nome) || origem.includes("investigador");
  };

  Regras.isRefeicaoExtraAutomatica = function isRefeicaoExtraAutomatica(obj) {
    const nome = Regras.norm(Regras.getNome(obj));
    const origem = Regras.norm(Regras.getOrigem(obj));
    const obs = Regras.norm(Regras.pick(obj, "Observacao", "Observação", "observacao", "Obs"));
    return nome === "refeicao extra" ||
           nome === "refeicao extra automatica" ||
           nome.includes("refeicao extra") ||
           origem.includes("extra automat") ||
           obs.includes("extra automat");
  };

  Regras.isExtra = function isExtra(obj) {
    const origem = Regras.norm(Regras.getOrigem(obj));
    const nome = Regras.norm(Regras.getNome(obj));
    const cid = Regras.norm(Regras.getColaboradorId(obj));
    return origem.includes("extra") || origem.includes("visitante") || origem.includes("prestador") ||
           origem.includes("terceiro") || origem.includes("fornecedor") || origem.includes("representante") ||
           origem.includes("motorista") || Regras.isGuarda(obj) || Regras.isInvestigador(obj) ||
           nome.includes("refeicao extra") || cid.startsWith("extra-");
  };

  Regras.isMarmitaCozinha = function isMarmitaCozinha(obj) {
    return Regras.isGuarda(obj) || Regras.isInvestigador(obj);
  };

  Regras.TIPOS_CARDAPIO_DIA_PERMITIDOS = [
    "colaborador",
    "funcionario",
    "funcionário",
    "prestador",
    "visitante",
    "terceiro"
  ];

  Regras.isTipoPermitidoCardapioDia = function isTipoPermitidoCardapioDia(pedido) {
    const origem = Regras.norm(Regras.getOrigem(pedido));
    const tipo = Regras.norm(Regras.pick(pedido, "tipo", "Tipo"));
    const base = tipo || origem || "colaborador";

    // Regra explícita: guarda/investigador só aparecem na cozinha.
    if (Regras.isGuarda(pedido) || Regras.isInvestigador(pedido)) return false;
    if (Regras.isRefeicaoExtraAutomatica(pedido)) return false;

    // Colaborador normal geralmente vem com Origem = Refeitório.
    if (base === "refeitorio" || base === "refeitório" || base === "automatico" || base === "automático" || base === "travamento") return true;

    return Regras.TIPOS_CARDAPIO_DIA_PERMITIDOS.includes(base);
  };

  Regras.visivelCardapioDia = function visivelCardapioDia(pedido) {
    if (!Regras.pedidoConfirmadoProducao(pedido)) return false;
    return Regras.isTipoPermitidoCardapioDia(pedido);
  };

  Regras.visivelCozinha = function visivelCozinha(pedido) {
    // Cozinha pode ver colaboradores e extras, mas produção exclui cancelado/bloqueado.
    const status = Regras.norm(Regras.getStatus(pedido));
    return !["cancelado", "bloqueado"].includes(status);
  };

  // ============================================================
  // AUSÊNCIAS
  // ============================================================
  Regras.ausenciaAtivaPorCampo = function ausenciaAtivaPorCampo(ausencia) {
    const status = Regras.norm(Regras.pick(ausencia, "Status", "status", "Status_Ausencia", "statusAusencia"));
    const ativo = Regras.pick(ausencia, "Ativo", "ativo");

    if ([
      "inativo",
      "cancelado",
      "cancelada",
      "duplicado inativado",
      "duplicidade inativada",
      "duplicada inativada",
      "excluido",
      "excluído",
      "false",
      "nao",
      "não",
      "0"
    ].includes(status)) return false;

    // "Período encerrado" continua valendo como histórico para datas dentro do intervalo.
    if (status === "periodo encerrado" || status === "período encerrado" || status === "encerrado") return true;

    if (ativo === null || ativo === undefined || String(ativo).trim() === "") return true;
    return Regras.isTrue(ativo);
  };

  Regras.ausenciaInicioISO = function ausenciaInicioISO(ausencia) {
    return Regras.dataISO(Regras.pick(ausencia, "Data_Inicio", "Inicio", "DataInicio", "Data"));
  };

  Regras.ausenciaFimISO = function ausenciaFimISO(ausencia) {
    return Regras.dataISO(Regras.pick(ausencia, "Data_Fim", "Fim", "DataFim", "Data")) || Regras.ausenciaInicioISO(ausencia);
  };

  Regras.ausenciaCobreData = function ausenciaCobreData(ausencia, dataISO) {
    const ini = Regras.ausenciaInicioISO(ausencia);
    const fim = Regras.ausenciaFimISO(ausencia);
    return !!ini && !!fim && ini <= dataISO && fim >= dataISO;
  };

  Regras.ausenciaVigenteParaColaborador = function ausenciaVigenteParaColaborador(ausencias, colabKey, dataISO) {
    const candidatas = (ausencias || []).filter(a => {
      if (!Regras.ausenciaAtivaPorCampo(a)) return false;
      const key = Regras.colaboradorKey({
        Colaborador_id: Regras.pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId"),
        Colaborador_nome: Regras.pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title")
      });
      return key === colabKey && Regras.ausenciaCobreData(a, dataISO);
    });

    if (!candidatas.length) return null;
    candidatas.sort((a, b) => {
      const fimA = Regras.ausenciaFimISO(a);
      const fimB = Regras.ausenciaFimISO(b);
      if (fimA !== fimB) return fimB.localeCompare(fimA);
      return Number(Regras.pick(b, "id", "ID") || 0) - Number(Regras.pick(a, "id", "ID") || 0);
    });
    return candidatas[0];
  };

  Regras.formatarMotivoAusencia = function formatarMotivoAusencia(motivo) {
    const n = Regras.norm(motivo);
    if (n === "nao_vai_almocar" || n === "nao vai almocar" || n === "não vai almoçar" || n === "ausente") return "Não vai almoçar";
    if (n === "ferias" || n === "férias") return "Férias";
    if (n === "afastado" || n === "afastamento") return "Afastado";
    if (n === "atestado") return "Atestado";
    if (n === "licenca" || n === "licença") return "Licença";
    if (n === "banco_horas" || n === "banco horas") return "Banco de horas";
    if (n === "homy_office" || n === "homy office") return "Homy Office";
    if (n === "falta") return "Falta";
    return String(motivo || "Ausente");
  };

  // ============================================================
  // DEDUPLICAÇÃO / ESCOLHA DO PEDIDO VÁLIDO
  // ============================================================
  Regras.timestampPedido = function timestampPedido(pedido) {
    const raw = Regras.pick(pedido, "Modified", "modified", "Data_Hora", "DataHora", "Created", "created", "Data") || "";
    const dt = raw ? new Date(raw) : null;
    if (dt && !isNaN(dt)) return dt.getTime();
    const id = Number(Regras.getPedidoId(pedido) || 0);
    return Number.isFinite(id) ? id : 0;
  };

  Regras.pedidoKey = function pedidoKey(pedido) {
    const dia = Regras.norm(Regras.getDia(pedido));
    const opcao = Regras.norm(Regras.getOpcao(pedido));
    const origem = Regras.norm(Regras.getOrigem(pedido));
    const obs = Regras.norm(Regras.pick(pedido, "Observacao", "Observação", "Obs"));
    const extraId = (obs.match(/extraid:\s*([^|\s]+)/i) || [])[1] || "";

    if (Regras.isPedidoAdicionalColaborador(pedido)) {
      return `adicional|${dia}|${Regras.getPedidoId(pedido) || Regras.colaboradorKey(pedido) || origem + "|" + opcao}`;
    }

    if (Regras.isExtra(pedido)) {
      return `extra|${dia}|${extraId || Regras.getColaboradorId(pedido) || Regras.norm(Regras.getNome(pedido)) + "|" + origem + "|" + opcao}`;
    }

    return `colab|${dia}|${Regras.colaboradorKey(pedido)}`;
  };

  Regras.scorePedido = function scorePedido(pedido) {
    let score = 0;
    const status = Regras.norm(Regras.getStatus(pedido));
    const origem = Regras.norm(Regras.getOrigem(pedido));

    if (Regras.getPedidoId(pedido)) score += 2;
    if (!origem.includes("travamento")) score += 2;

    if (Regras.isRetornoAutomaticoAusencia(pedido)) score -= 80;
    else if (["cancelado", "bloqueado"].includes(status)) score -= 100;
    else if (Regras.pedidoConfirmadoProducao(pedido)) score += 90;
    else if (Regras.pedidoAusente(pedido)) score += 60;
    else if (status === "travado") score += 20;

    return score;
  };

  Regras.compararPedidoPreferido = function compararPedidoPreferido(a, b) {
    const sa = Regras.scorePedido(a);
    const sb = Regras.scorePedido(b);
    if (sa !== sb) return sb - sa;
    return Regras.timestampPedido(b) - Regras.timestampPedido(a);
  };

  Regras.deduplicarPedidos = function deduplicarPedidos(pedidos) {
    const mapa = new Map();
    for (const pedido of pedidos || []) {
      const key = Regras.pedidoKey(pedido);
      const atual = mapa.get(key);
      if (!atual || Regras.compararPedidoPreferido(pedido, atual) < 0) {
        mapa.set(key, pedido);
      }
    }
    return Array.from(mapa.values());
  };

  // ============================================================
  // OPERAÇÃO DO DIA — RECONCILIAÇÃO DE LISTA VIVA
  // ============================================================
  Regras.isItemOperacaoEspecial = function isItemOperacaoEspecial(item) {
    return Regras.isExtra(item) || Regras.isPedidoAdicionalColaborador(item) || !!item?._virtualExtra || !!item?._refeicaoAdicional;
  };

  Regras.scoreItemOperacao = function scoreItemOperacao(item) {
    let score = 0;
    const status = Regras.norm(Regras.getStatus(item));
    const origem = Regras.norm(Regras.getOrigem(item));

    if (Regras.getPedidoId(item)) score += 5;
    if (Regras.isTravamentoAutomatico(item)) score += 900;
    else if (Regras.pedidoEscolhaReal(item)) score += 850;
    else if (Regras.pedidoAusente(item) || item?._virtualAusencia || item?._ausenciaOperacao) score += 700;
    else if (item?._legadoAusenciaEncerrada || item?._retornoAutomaticoLegado || item?._ausenciaEncerradaAguardandoMarcacao || origem.includes("ausencia encerrada")) score += 250;
    else if (origem.includes("sem pedido")) score += 350;
    else if (item?._virtualPendente || status === "pendente") score += 300;
    if (Regras.isRetornoAutomaticoAusencia(item)) score -= 50;
    if (Regras.pedidoCanceladoOuBloqueado(item)) score -= 500;
    score += Math.min(Math.floor(Regras.timestampPedido(item) / 100000000000), 99);
    return score;
  };

  Regras.preferirItemOperacao = function preferirItemOperacao(a, b) {
    if (!a) return b;
    if (!b) return a;
    const sa = Regras.scoreItemOperacao(a);
    const sb = Regras.scoreItemOperacao(b);
    if (sa !== sb) return sa > sb ? a : b;
    return Regras.timestampPedido(a) >= Regras.timestampPedido(b) ? a : b;
  };

  Regras.deduplicarListaOperacao = function deduplicarListaOperacao(lista = []) {
    const especiais = [];
    const grupos = [];

    for (const item of lista || []) {
      if (Regras.isItemOperacaoEspecial(item)) {
        especiais.push(item);
        continue;
      }

      const dia = Regras.norm(Regras.getDia(item));
      const keys = Regras.chavesColaboradorOperacional(item, "pedido");
      if (!keys.length) {
        especiais.push(item);
        continue;
      }

      let grupo = null;
      for (const g of grupos) {
        if (g.dia !== dia) continue;
        if (keys.some(k => g.keys.has(k))) {
          grupo = g;
          break;
        }
      }
      if (!grupo) {
        grupo = { dia, keys: new Set(), item: null };
        grupos.push(grupo);
      }
      for (const k of keys) grupo.keys.add(k);
      grupo.item = Regras.preferirItemOperacao(grupo.item, item);
    }

    return [...especiais, ...grupos.map(g => g.item).filter(Boolean)];
  };

  Regras.existeColaboradorNaListaOperacao = function existeColaboradorNaListaOperacao(lista = [], colaborador) {
    return (lista || []).some(item => {
      if (Regras.isItemOperacaoEspecial(item)) return false;
      return Regras.mesmoColaboradorOperacional(item, colaborador, "pedido", "colaborador");
    });
  };

  // ============================================================
  // RESUMOS
  // ============================================================
  Regras.calcularResumoProducao = function calcularResumoProducao(pedidos, options = {}) {
    const dia = options.dia ? Regras.norm(options.dia) : "";
    const base = (pedidos || []).filter(p => {
      if (dia && Regras.norm(Regras.getDia(p)) !== dia) return false;
      return true;
    });

    const deduplicados = options.deduplicar === false ? base : Regras.deduplicarPedidos(base);
    const incluidos = [];
    const excluidos = [];

    for (const p of deduplicados) {
      if (Regras.pedidoConfirmadoProducao(p)) {
        incluidos.push(p);
      } else {
        excluidos.push({ pedido: p, motivo: Regras.motivoExclusaoProducao(p) });
      }
    }

    const porOpcao = { principal: 0, light: 0, carne: 0, massa: 0, lanche: 0 };
    for (const p of incluidos) {
      const op = Regras.norm(Regras.getOpcao(p) || "principal");
      if (Object.prototype.hasOwnProperty.call(porOpcao, op)) porOpcao[op]++;
    }

    return {
      total: incluidos.length,
      principal: porOpcao.principal,
      light: porOpcao.light,
      carne: porOpcao.carne,
      massa: porOpcao.massa,
      lanche: porOpcao.lanche,
      incluidos,
      excluidos,
      base: deduplicados
    };
  };

  Regras.filtrarFuncionariosCardapioDia = function filtrarFuncionariosCardapioDia(pedidos, options = {}) {
    const dia = options.dia ? Regras.norm(options.dia) : "";
    const deduplicados = options.deduplicar === false ? (pedidos || []) : Regras.deduplicarPedidos(pedidos || []);

    return deduplicados
      .filter(p => !dia || Regras.norm(Regras.getDia(p)) === dia)
      .filter(p => Regras.visivelCardapioDia(p))
      .map(p => ({
        id: Regras.getPedidoId(p),
        nome: Regras.getNome(p) || "Colaborador",
        dept: Regras.getCentroCusto(p) || "",
        opcao: Regras.norm(Regras.getOpcao(p) || "principal"),
        pedido: p
      }))
      .sort((a, b) => String(a.nome).localeCompare(String(b.nome), "pt-BR", { sensitivity: "base", numeric: true }));
  };

  Regras.filtrarCozinha = function filtrarCozinha(pedidos, options = {}) {
    const dia = options.dia ? Regras.norm(options.dia) : "";
    const lista = (pedidos || [])
      .filter(p => !dia || Regras.norm(Regras.getDia(p)) === dia)
      .filter(Regras.visivelCozinha);
    return options.deduplicar === false ? lista : Regras.deduplicarPedidos(lista);
  };

  Regras.diagnosticarRetornosEsperados = function diagnosticarRetornosEsperados({ colaboradores = [], pedidos = [], ausencias = [], semanaId = "", dia = "", helpers = {} } = {}) {
    const dataRef = Regras.dataPorSemanaDia(semanaId, dia, helpers);
    if (!dataRef) return [];

    const pedidosDia = (pedidos || []).filter(p => Regras.norm(Regras.getDia(p)) === Regras.norm(dia));
    const pedidosPorColab = new Map();
    for (const p of pedidosDia) {
      const key = Regras.colaboradorKey(p);
      if (key) pedidosPorColab.set(key, p);
    }

    const alertas = [];
    for (const c of colaboradores || []) {
      const key = Regras.colaboradorKey(c);
      if (!key) continue;
      const ausenciaHoje = Regras.ausenciaVigenteParaColaborador(ausencias, key, dataRef);
      if (ausenciaHoje) continue;

      const ausenciasAnteriores = (ausencias || []).filter(a => {
        const aKey = Regras.colaboradorKey({
          Colaborador_id: Regras.pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId"),
          Colaborador_nome: Regras.pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title")
        });
        return aKey === key && Regras.ausenciaAtivaPorCampo(a) && Regras.ausenciaFimISO(a) && Regras.ausenciaFimISO(a) < dataRef;
      });

      if (!ausenciasAnteriores.length) continue;
      const pedido = pedidosPorColab.get(key);
      if (!pedido) {
        alertas.push({
          tipo: "retorno-sem-pedido",
          colaborador: c,
          dataRef,
          dia,
          mensagem: `${Regras.getNome(c) || key} teve ausência encerrada antes de ${dataRef} e não possui pedido para o dia.`
        });
        continue;
      }

      if (!Regras.pedidoConfirmadoProducao(pedido) && Regras.pedidoAusente(pedido)) {
        alertas.push({
          tipo: "retorno-bloqueado-por-status-antigo",
          colaborador: c,
          pedido,
          dataRef,
          dia,
          mensagem: `${Regras.getNome(c) || key} tem pedido no dia, mas continua com status de ausência vencida.`
        });
      }
    }

    return alertas;
  };



  // ============================================================
  // FECHAMENTO OFICIAL / TRAVAMENTO / RELATÓRIOS
  // ============================================================
  Regras.DIAS_OPERACIONAIS = ["segunda", "terca", "quarta", "quinta", "sexta"];

  Regras.diasOperacionais = function diasOperacionais() {
    return [...Regras.DIAS_OPERACIONAIS];
  };

  Regras.statusFechamentoAtivo = function statusFechamentoAtivo(item) {
    const status = Regras.norm(Regras.pick(item, "Status_Fechamento", "Status") || "");
    return status === "fechado" || status === "recalculado";
  };

  Regras.dataFechamentoOrdem = function dataFechamentoOrdem(item) {
    const raw = Regras.pick(item, "Fechado_Em", "Modified", "Gerado_Em", "Created") || "";
    const t = raw ? new Date(raw).getTime() : 0;
    return Number.isFinite(t) ? t : Number(Regras.pick(item, "id", "ID") || 0) || 0;
  };

  Regras.resumoFechamentoParaTela = function resumoFechamentoParaTela(fechamento) {
    if (!fechamento) return null;
    const n = (...keys) => Number(Regras.pick(fechamento, ...keys) || 0) || 0;
    return {
      id: String(Regras.pick(fechamento, "id", "ID") || ""),
      title: Regras.pick(fechamento, "Title") || "",
      semanaId: Regras.pick(fechamento, "Semana_id") || "",
      dia: Regras.norm(Regras.pick(fechamento, "Dia") || ""),
      data: Regras.dataISO(Regras.pick(fechamento, "Data_Operacao", "Data", "Created")),
      status: Regras.pick(fechamento, "Status_Fechamento", "Status") || "",
      total: n("Total"),
      principal: n("Principal"),
      light: n("Light"),
      carne: n("Carne"),
      massa: n("Massa"),
      lanche: n("Lanche"),
      extras: n("Extras"),
      cancelados: n("Cancelados"),
      ausentes: n("Ausentes"),
      checkins: n("Checkins"),
      fechadoPor: Regras.pick(fechamento, "Fechado_Por") || "",
      fechadoEm: Regras.pick(fechamento, "Fechado_Em", "Modified", "Created") || "",
      observacao: Regras.pick(fechamento, "Observacao") || ""
    };
  };

  Regras.analisarFechamentoSemana = function analisarFechamentoSemana(fechamentos = [], options = {}) {
    const semanaId = options.semanaId || Regras.pick((fechamentos || [])[0] || {}, "Semana_id") || "";
    const dias = options.diasOperacionais || Regras.diasOperacionais();
    const ativos = (fechamentos || [])
      .filter(Regras.statusFechamentoAtivo)
      .sort((a, b) => Regras.dataFechamentoOrdem(b) - Regras.dataFechamentoOrdem(a));
    const porDia = {};
    for (const item of ativos) {
      const dia = Regras.norm(Regras.pick(item, "Dia") || "");
      if (!dia || porDia[dia]) continue;
      porDia[dia] = Regras.resumoFechamentoParaTela(item);
    }
    const diasFechados = dias.filter(d => !!porDia[d]);
    const temFechamento = diasFechados.length > 0;
    const semanaFechada = dias.length > 0 && dias.every(d => !!porDia[d]);
    return {
      semanaId,
      temFechamento,
      semanaFechada,
      diasOperacionais: dias,
      diasFechados,
      diasAbertos: dias.filter(d => !porDia[d]),
      totalDiasFechados: diasFechados.length,
      porDia,
      mensagem: semanaFechada
        ? "Semana fechada: relatório usa Fechamento Oficial."
        : (temFechamento ? "Semana parcialmente fechada: travamento será aplicado somente nos dias ainda abertos." : "Semana sem fechamento oficial.")
    };
  };

  Regras.diaFechadoNoTravamento = function diaFechadoNoTravamento(fechamentoSemana, dia) {
    const d = Regras.norm(dia || "");
    return !!(d && fechamentoSemana?.porDia && fechamentoSemana.porDia[d]);
  };

  Regras.motivoBloqueioTravamentoSemana = function motivoBloqueioTravamentoSemana(fechamentoSemana) {
    if (!fechamentoSemana?.temFechamento) return "";
    if (fechamentoSemana.semanaFechada) {
      return "Travamento bloqueado: todos os dias da semana já possuem Fechamento Oficial. Use reabertura/auditoria, não travamento.";
    }
    return "A semana possui dia(s) fechado(s). O travamento será aplicado somente nos dias ainda abertos.";
  };

  Regras.podeTravarSemana = function podeTravarSemana({ fechamentoSemana = null, prazoInfo = null, permitirAntesDoPrazo = false, ignorarFechamento = false } = {}) {
    if (!ignorarFechamento && fechamentoSemana?.semanaFechada) {
      return { podeTravar: false, bloqueado: true, bloqueadoPorFechamento: true, bloqueadoPorPrazo: false, motivoBloqueio: Regras.motivoBloqueioTravamentoSemana(fechamentoSemana) };
    }
    const prazoConfigurado = !!prazoInfo?.prazoConfigurado;
    const prazoVencido = !!prazoInfo?.prazoVencido;
    const podePorPrazo = permitirAntesDoPrazo || (prazoConfigurado && prazoVencido);
    if (!podePorPrazo) {
      return {
        podeTravar: false,
        bloqueado: true,
        bloqueadoPorFechamento: false,
        bloqueadoPorPrazo: true,
        motivoBloqueio: !prazoConfigurado ? "Travamento bloqueado: prazo da semana não está configurado." : "Travamento bloqueado: prazo da semana ainda não venceu."
      };
    }
    return { podeTravar: true, bloqueado: false, bloqueadoPorFechamento: false, bloqueadoPorPrazo: false, motivoBloqueio: "" };
  };

  Regras.motivoTotalZeroValido = function motivoTotalZeroValido(motivo) {
    const texto = String(motivo || "").trim();
    const n = Regras.norm(texto);
    if (texto.length < 5) return false;
    return !["fechamento conferido pela operacao do dia", "fechamento oficial gerado pela operacao do dia", "fechamento oficial confirmado", "ok", "sim"].includes(n);
  };

  Regras.definirFonteRelatorioDia = function definirFonteRelatorioDia({ fechamento = null } = {}) {
    return fechamento && Regras.statusFechamentoAtivo(fechamento)
      ? { fonte: "Fechamento Oficial", usaFechamento: true }
      : { fonte: "Pedidos calculados", usaFechamento: false };
  };

  Regras.opcaoResumoFechamento = function opcaoResumoFechamento(opcao) {
    const op = Regras.norm(opcao || "");
    if (["principal", "light", "carne", "massa", "lanche"].includes(op)) return op;
    if (!op) return "semOpcao";
    return "outros";
  };

  Regras.categoriaPedidoFechamento = function categoriaPedidoFechamento(p) {
    const origem = Regras.norm(Regras.getOrigem(p));
    const nome = Regras.norm(Regras.getNome(p));
    const id = Regras.norm(Regras.getColaboradorId(p));
    if (origem.includes("investigador") || nome.includes("investigador")) return "investigador";
    if (origem.includes("guarda") || nome.includes("guarda")) return "guarda";
    if (origem.includes("prestador")) return "prestador";
    if (origem.includes("visitante")) return "visitante";
    if (origem.includes("terceiro")) return "terceiro";
    if (origem.includes("extra") || nome.includes("refeicao extra") || id.startsWith("extra-")) return "extra";
    if (origem.includes("ausencia") || origem.includes("ausência")) return "ausencia";
    return "colaborador";
  };

  Regras.statusBloqueiaFechamento = function statusBloqueiaFechamento(p) {
    const status = Regras.norm(Regras.getStatus(p));
    return Regras.statusBloqueiaProducao(status) || status === "duplicado inativado" || status === "duplicidade inativada";
  };

  Regras.pedidoProdutivoFechamento = function pedidoProdutivoFechamento(p) {
    return Regras.pedidoConfirmadoProducao(p);
  };

  Regras.pedidoKeyFechamento = function pedidoKeyFechamento(p) {
    const categoria = Regras.categoriaPedidoFechamento(p);
    const dia = Regras.norm(Regras.getDia(p));
    const opcao = Regras.norm(Regras.getOpcao(p) || "principal");
    const id = Regras.getColaboradorId(p);
    const nome = Regras.norm(Regras.getNome(p));
    if (["extra", "guarda", "investigador", "prestador", "visitante", "terceiro"].includes(categoria)) {
      return `especial|${dia}|${categoria}|${opcao}|${id || nome}`;
    }
    if (id) return `colaborador|${dia}|id:${id}`;
    return `colaborador|${dia}|nome:${nome}`;
  };

  Regras.compararPreferenciaFechamento = function compararPreferenciaFechamento(a, b) {
    const ap = Regras.pedidoProdutivoFechamento(a) ? 1 : 0;
    const bp = Regras.pedidoProdutivoFechamento(b) ? 1 : 0;
    if (ap !== bp) return bp - ap;
    const ab = Regras.statusBloqueiaFechamento(a) ? 1 : 0;
    const bb = Regras.statusBloqueiaFechamento(b) ? 1 : 0;
    if (ab !== bb) return ab - bb;
    return Regras.timestampPedido(b) - Regras.timestampPedido(a);
  };



  // ============================================================
  // OPERAÇÃO DO DIA CONGELADA POR FECHAMENTO OFICIAL
  // ============================================================
  Regras.operacaoDiaUsaFechamento = function operacaoDiaUsaFechamento(fechamento) {
    return !!fechamento && Regras.statusFechamentoAtivo(fechamento);
  };

  Regras.parseSnapshotFechamento = function parseSnapshotFechamento(fechamento) {
    const raw = Regras.pick(fechamento, "Snapshot_JSON", "Snapshot", "Resumo_JSON") || "";
    if (!raw) return null;
    if (typeof raw === "object") return raw;
    try {
      return JSON.parse(String(raw));
    } catch (_) {
      return null;
    }
  };

  Regras.itemSnapshotParaPedidoOperacao = function itemSnapshotParaPedidoOperacao(item, tipo, fechamento) {
    const raw = item?.raw || item?.pedido || item || {};
    const statusInformado = Regras.pick(item, "status", "Status") || Regras.pick(raw, "Status", "status");
    const origemInformada = Regras.pick(item, "origem", "Origem") || Regras.pick(raw, "Origem", "origem") || (tipo === "incluido" ? "Fechamento Oficial" : "Fechamento Oficial - não produção");
    const opcaoInformada = Regras.pick(item, "opcao", "Opcao") || Regras.pick(raw, "Opcao", "opcao") || "principal";

    let status = statusInformado || (tipo === "incluido" ? "Confirmado" : "Não conta produção");
    let nomePrato = Regras.pick(raw, "Nome_Prato", "NomePrato") || Regras.pick(item, "prato", "nomePrato") || "";

    // Se o item estava fora da produção e veio de retorno automático legado,
    // ele não deve aparecer como pendente em dia fechado. É histórico excluído.
    if (tipo === "excluido" && Regras.isRetornoAutomaticoAusencia(raw)) {
      status = "Retorno legado ignorado";
      nomePrato = nomePrato || "Sem produção";
    }

    if (tipo === "excluido" && !nomePrato) nomePrato = status || "Não conta produção";
    if (tipo === "incluido" && !nomePrato) nomePrato = "Prato Principal";

    return {
      ...raw,
      id: String(Regras.pick(raw, "id", "ID") || Regras.pick(item, "pedidoId", "id", "ID") || ""),
      Title: Regras.pick(raw, "Title") || Regras.pick(item, "title") || `${Regras.pick(fechamento, "Semana_id") || ""}-${Regras.pick(fechamento, "Dia") || ""}`,
      Semana_id: Regras.pick(raw, "Semana_id") || Regras.pick(item, "semanaId") || Regras.pick(fechamento, "Semana_id"),
      Dia: Regras.pick(raw, "Dia") || Regras.pick(item, "dia") || Regras.pick(fechamento, "Dia"),
      Colaborador_id: Regras.pick(raw, "Colaborador_id", "ColaboradorId", "colaboradorId") || Regras.pick(item, "colaboradorId"),
      Colaborador_nome: Regras.pick(raw, "Colaborador_nome", "Colaborador", "Nome", "Title") || Regras.pick(item, "colaboradorNome", "nome") || "Colaborador",
      Centro_Custo: Regras.pick(raw, "Centro_Custo", "CentroCusto", "Setor", "Departamento") || Regras.pick(item, "centroCusto"),
      Opcao: opcaoInformada || "principal",
      Nome_Prato: nomePrato,
      Status: status,
      Confirmado: tipo === "incluido" ? true : false,
      Origem: origemInformada,
      Modified: Regras.pick(raw, "Modified", "modified") || Regras.pick(item, "modified") || Regras.pick(fechamento, "Fechado_Em", "Modified", "Created"),
      _snapshotFechamento: true,
      _snapshotTipo: tipo,
      _fechamentoId: String(Regras.pick(fechamento, "id", "ID") || ""),
      _contaProducaoSnapshot: tipo === "incluido"
    };
  };

  Regras.listaOperacaoPorFechamento = function listaOperacaoPorFechamento(fechamento) {
    if (!Regras.operacaoDiaUsaFechamento(fechamento)) return [];
    const snapshot = Regras.parseSnapshotFechamento(fechamento);
    if (!snapshot) return [];

    const incluidos = Array.isArray(snapshot.incluidos) ? snapshot.incluidos : [];
    const excluidos = Array.isArray(snapshot.excluidos) ? snapshot.excluidos : [];

    return [
      ...incluidos.map(item => Regras.itemSnapshotParaPedidoOperacao(item, "incluido", fechamento)),
      ...excluidos.map(item => Regras.itemSnapshotParaPedidoOperacao(item, "excluido", fechamento))
    ];
  };


  // ============================================================
  // INTEGRIDADE / LIMPEZA ASSISTIDA — REGRAS CENTRAIS
  // ============================================================
  // Estas funções NÃO acessam SharePoint e NÃO gravam dados.
  // Elas só decidem: o que é sujeira, o que pode ser aplicado com segurança,
  // qual registro manter, qual cancelar e o que deve ficar para revisão.

  Regras.resumoCorrecaoVazio = function resumoCorrecaoVazio() {
    return { total: 0, principal: 0, light: 0, carne: 0, massa: 0, lanche: 0, outros: 0, semOpcao: 0 };
  };

  Regras.normalizarResumoCorrecao = function normalizarResumoCorrecao(resumo = {}) {
    const out = Regras.resumoCorrecaoVazio();
    for (const k of Object.keys(out)) out[k] = Number(resumo?.[k] || 0);
    return out;
  };

  Regras.deltaResumoCorrecao = function deltaResumoCorrecao(atual = {}, alvo = {}) {
    const a = Regras.normalizarResumoCorrecao(atual);
    const b = Regras.normalizarResumoCorrecao(alvo);
    return {
      total: a.total - b.total,
      principal: a.principal - b.principal,
      light: a.light - b.light,
      carne: a.carne - b.carne,
      massa: a.massa - b.massa,
      lanche: a.lanche - b.lanche
    };
  };

  Regras.resumoBateCorrecao = function resumoBateCorrecao(atual = {}, alvo = {}) {
    const d = Regras.deltaResumoCorrecao(atual, alvo);
    return ["total", "principal", "light", "carne", "massa", "lanche"].every(k => Number(d[k] || 0) === 0);
  };

  Regras.aplicarDeltaResumoCorrecao = function aplicarDeltaResumoCorrecao(resumo = {}, delta = {}) {
    const r = Regras.normalizarResumoCorrecao(resumo);
    for (const k of ["total", "principal", "light", "carne", "massa", "lanche", "outros", "semOpcao"]) {
      if (delta[k]) r[k] = Number(r[k] || 0) + Number(delta[k] || 0);
    }
    return r;
  };

  Regras.opcaoCorrecao = function opcaoCorrecao(opcao) {
    const op = Regras.norm(opcao || "principal");
    if (op.includes("carne")) return "carne";
    if (op.includes("light") || op.includes("salada")) return "light";
    if (op.includes("massa")) return "massa";
    if (op.includes("lanche")) return "lanche";
    if (op.includes("principal")) return "principal";
    return ["principal", "light", "carne", "massa", "lanche"].includes(op) ? op : "outros";
  };

  Regras._observacaoCorrecao = function _observacaoCorrecao(item, texto) {
    const raw = item?.raw || item?.pedido || {};
    const anterior = Regras.pick(raw, "Observacao", "Observação", "observacao", "Obs") || item?.observacao || "";
    const complemento = `Correção assistida: ${texto}`;
    if (!anterior) return complemento;
    if (Regras.norm(anterior).includes(Regras.norm(complemento))) return anterior;
    return `${anterior} | ${complemento}`;
  };

  Regras.criarAcaoCancelarPedidoCorrecao = function criarAcaoCancelarPedidoCorrecao(item, motivo, justificativa, manterId = "") {
    const opcao = Regras.opcaoCorrecao(item?.opcao || Regras.getOpcao(item) || "principal");
    const delta = { total: -1 };
    delta[opcao] = -1;
    const textoObs = manterId
      ? `${justificativa} Mantido o registro ${manterId}.`
      : justificativa;
    return {
      acao: "cancelar",
      autoAplicavel: true,
      motivo,
      pedidoId: String(item?.pedidoId || Regras.getPedidoId(item?.raw || item?.pedido || item) || ""),
      nome: item?.colaboradorNome || item?.nome || Regras.getNome(item?.raw || item?.pedido || item) || "",
      dia: item?.dia || Regras.getDia(item?.raw || item?.pedido || item) || "",
      opcao,
      statusAtual: item?.status || Regras.getStatus(item?.raw || item?.pedido || item) || "",
      confirmadoAtual: !!item?.confirmado,
      origemAtual: item?.origem || Regras.getOrigem(item?.raw || item?.pedido || item) || "",
      categoria: item?.categoria || "outro",
      manterId: manterId || "",
      delta,
      justificativa,
      camposSugeridos: {
        Status: "Cancelado",
        Confirmado: false,
        Origem: motivo === "retorno-automatico-legado-regularizado" ? "Retorno automático legado inativado" : "Duplicidade inativada",
        Observacao: Regras._observacaoCorrecao(item, textoObs)
      },
      pedido: item
    };
  };

  Regras.criarAcaoReativarPedidoCorrecao = function criarAcaoReativarPedidoCorrecao(item, motivo, justificativa) {
    const opcao = Regras.opcaoCorrecao(item?.opcao || Regras.getOpcao(item) || "principal");
    const delta = { total: 1 };
    delta[opcao] = 1;
    return {
      acao: "reativar",
      autoAplicavel: false,
      motivo,
      pedidoId: String(item?.pedidoId || Regras.getPedidoId(item?.raw || item?.pedido || item) || ""),
      nome: item?.colaboradorNome || item?.nome || Regras.getNome(item?.raw || item?.pedido || item) || "",
      dia: item?.dia || Regras.getDia(item?.raw || item?.pedido || item) || "",
      opcao,
      statusAtual: item?.status || Regras.getStatus(item?.raw || item?.pedido || item) || "",
      confirmadoAtual: !!item?.confirmado,
      origemAtual: item?.origem || Regras.getOrigem(item?.raw || item?.pedido || item) || "",
      categoria: item?.categoria || "outro",
      delta,
      justificativa,
      camposSugeridos: {
        Status: "Confirmado",
        Confirmado: true,
        Opcao: opcao === "outros" ? (item?.opcao || "principal") : opcao,
        Origem: "Correção de integridade",
        Observacao: Regras._observacaoCorrecao(item, justificativa)
      },
      pedido: item
    };
  };

  Regras.criarAcaoRevisarCorrecao = function criarAcaoRevisarCorrecao(dia, opcao, quantidade, justificativa, candidatos = []) {
    return {
      acao: "revisar",
      autoAplicavel: false,
      motivo: "revisao-manual",
      dia,
      opcao,
      quantidade: Number(quantidade || 0),
      delta: {},
      justificativa,
      candidatos: (candidatos || []).slice(0, 20)
    };
  };

  Regras.chaveDuplicidadeEspecialCorrecao = function chaveDuplicidadeEspecialCorrecao(item, dia) {
    const cat = Regras.norm(item?.categoria || "");
    const nome = Regras.norm(item?.colaboradorNome || item?.nome || Regras.getNome(item?.raw || item?.pedido || item) || "");
    const colabId = Regras.norm(item?.colaboradorId || Regras.getColaboradorId(item?.raw || item?.pedido || item) || "");
    const opcao = Regras.opcaoCorrecao(item?.opcao || Regras.getOpcao(item?.raw || item?.pedido || item) || "principal");
    // Para Guarda/Investigador, o nome sequencial é essencial: Investigador 1 não é Investigador 2.
    return [Regras.norm(dia), cat, colabId, nome, opcao].join("|");
  };

  Regras.pedidoEspecialLimpezaCorrecao = function pedidoEspecialLimpezaCorrecao(item) {
    const cat = Regras.norm(item?.categoria || "");
    return ["guarda", "investigador", "extra", "prestador", "visitante", "terceiro"].includes(cat);
  };

  Regras.compararPedidoManterCorrecao = function compararPedidoManterCorrecao(a, b) {
    // Mantém o registro mais antigo/menor ID para preservar o vínculo original.
    const ia = Number(a?.pedidoId || Regras.getPedidoId(a?.raw || a?.pedido || a) || 0);
    const ib = Number(b?.pedidoId || Regras.getPedidoId(b?.raw || b?.pedido || b) || 0);
    if (Number.isFinite(ia) && Number.isFinite(ib) && ia !== ib) return ia - ib;
    return Regras.timestampPedido(a?.raw || a?.pedido || a) - Regras.timestampPedido(b?.raw || b?.pedido || b);
  };

  Regras.gerarCandidatasCorrecaoDia = function gerarCandidatasCorrecaoDia(calc = {}, alvo = {}, dia = "") {
    const incluidos = calc?.incluidos || [];
    const excluidos = calc?.excluidos || [];
    const candidatas = [];

    // 1) Extras/pedidos especiais duplicados: cancelar todos menos o preferido.
    const gruposExtras = new Map();
    for (const item of incluidos) {
      if (!Regras.pedidoEspecialLimpezaCorrecao(item)) continue;
      const key = Regras.chaveDuplicidadeEspecialCorrecao(item, dia);
      if (!gruposExtras.has(key)) gruposExtras.set(key, []);
      gruposExtras.get(key).push(item);
    }

    for (const grupo of gruposExtras.values()) {
      if (grupo.length <= 1) continue;
      const ordenado = [...grupo].sort(Regras.compararPedidoManterCorrecao);
      const manter = ordenado[0];
      for (const dup of ordenado.slice(1)) {
        candidatas.push(Regras.criarAcaoCancelarPedidoCorrecao(
          dup,
          "extra-duplicado",
          `Duplicidade de ${dup.categoria || "pedido especial"}: manter ID ${manter.pedidoId} e cancelar este registro.`,
          manter.pedidoId
        ));
      }
    }

    // 2) Retorno automático legado ainda contado em produção: cancelar.
    for (const item of incluidos) {
      const origem = Regras.norm(item?.origem || Regras.getOrigem(item?.raw || item?.pedido || item) || "");
      if (!origem.includes("retorno automatico") && !origem.includes("retorno automático")) continue;
      candidatas.push(Regras.criarAcaoCancelarPedidoCorrecao(
        item,
        "retorno-automatico-legado-regularizado",
        "Retorno automático legado não é escolha real nem travamento oficial; inativar para não contaminar relatórios."
      ));
    }

    // 3) Reativação é sempre revisão manual por segurança.
    const deltaInicial = Regras.deltaResumoCorrecao(calc?.resumo || {}, alvo || {});
    for (const item of excluidos) {
      const status = Regras.norm(item?.status || Regras.getStatus(item?.raw || item?.pedido || item) || "");
      const origem = Regras.norm(item?.origem || Regras.getOrigem(item?.raw || item?.pedido || item) || "");
      const op = Regras.opcaoCorrecao(item?.opcao || Regras.getOpcao(item?.raw || item?.pedido || item) || "principal");
      if (Number(deltaInicial[op] || 0) >= 0) continue;
      if (status !== "cancelado" && !origem.includes("duplicado inativado") && !origem.includes("duplicidade inativada")) continue;
      candidatas.push(Regras.criarAcaoReativarPedidoCorrecao(
        item,
        "pedido-correto-cancelado-revisar",
        `Possível pedido ${op} cancelado/inativado necessário para fechar a referência. Reativação exige revisão manual.`
      ));
    }

    return candidatas;
  };

  Regras.selecionarAcoesQueAproximamCorrecao = function selecionarAcoesQueAproximamCorrecao(atual, alvo, candidatas = []) {
    let simulado = Regras.normalizarResumoCorrecao(atual);
    const selecionadas = [];
    const revisoes = [];

    const aindaSobra = (opcao) => {
      const d = Regras.deltaResumoCorrecao(simulado, alvo);
      return Number(d.total || 0) > 0 && Number(d[opcao] || 0) > 0;
    };

    // Só aplica automaticamente cancelamentos seguros.
    const cancelamentos = (candidatas || [])
      .filter(a => a.acao === "cancelar" && a.autoAplicavel)
      .sort((a, b) => {
        const ordem = { "extra-duplicado": 1, "retorno-automatico-legado-regularizado": 2, "retorno-automatico-retroativo": 2 };
        return (ordem[a.motivo] || 9) - (ordem[b.motivo] || 9);
      });

    for (const acao of cancelamentos) {
      const op = Regras.opcaoCorrecao(acao.opcao || "principal");
      if (!aindaSobra(op)) continue;
      selecionadas.push(acao);
      simulado = Regras.aplicarDeltaResumoCorrecao(simulado, acao.delta);
    }

    for (const acao of (candidatas || []).filter(a => a.acao !== "cancelar" || !a.autoAplicavel)) {
      revisoes.push(acao);
    }

    return { selecionadas, simulado, revisoes };
  };

  Regras.gerarPlanoCorrecaoDia = function gerarPlanoCorrecaoDia({ semanaId, dia, dataOperacao, calc, referencia }) {
    const atual = Regras.normalizarResumoCorrecao(calc?.resumo || {});
    const alvo = Regras.normalizarResumoCorrecao(referencia?.valores || referencia || {});
    const deltaInicial = Regras.deltaResumoCorrecao(atual, alvo);
    const candidatas = Regras.gerarCandidatasCorrecaoDia(calc, alvo, dia);
    const { selecionadas, simulado, revisoes: revisoesCandidatas } = Regras.selecionarAcoesQueAproximamCorrecao(atual, alvo, candidatas);
    const deltaFinal = Regras.deltaResumoCorrecao(simulado, alvo);
    const fechaExato = Regras.resumoBateCorrecao(simulado, alvo);
    const jaBate = Regras.resumoBateCorrecao(atual, alvo);

    const revisoes = [...(revisoesCandidatas || [])];
    for (const op of ["principal", "light", "carne", "massa", "lanche"]) {
      const sobra = Number(deltaFinal[op] || 0);
      if (sobra > 0) {
        const candidatos = (calc?.incluidos || []).filter(i => Regras.opcaoCorrecao(i.opcao || "principal") === op);
        revisoes.push(Regras.criarAcaoRevisarCorrecao(dia, op, sobra, `Ainda sobram ${sobra} ${op} após aplicar as ações seguras.`, candidatos));
      }
      if (sobra < 0) {
        revisoes.push(Regras.criarAcaoRevisarCorrecao(dia, op, Math.abs(sobra), `Ainda faltam ${Math.abs(sobra)} ${op} após aplicar as ações seguras.`, []));
      }
    }

    return {
      semanaId,
      dia,
      dataOperacao: dataOperacao || "",
      referencia,
      atual,
      alvo,
      deltaInicial,
      candidatas,
      acoesSeguras: selecionadas,
      revisoes,
      simulado,
      deltaFinal,
      jaBate,
      fechaExato,
      status: jaBate ? "ok" : (fechaExato ? "corrigivel" : (selecionadas.length ? "parcial" : "revisao")),
      mensagem: jaBate
        ? "Base atual já bate com a referência."
        : (fechaExato ? "Ações seguras fecham exatamente com a referência." : "Há ações seguras, mas ainda fica pendência para revisão.")
    };
  };

  Regras.podeAplicarAcaoCorrecao = function podeAplicarAcaoCorrecao(acao) {
    if (!acao || !acao.autoAplicavel || !acao.pedidoId) return false;
    if (acao.acao !== "cancelar") return false;
    return ["extra-duplicado", "retorno-automatico-legado-regularizado", "retorno-automatico-retroativo"].includes(acao.motivo);
  };


  // Exporta para browser e para testes em Node, sem dependências.
  global.HomyRefeitorioRegras = Regras;
  if (typeof module !== "undefined" && module.exports) module.exports = Regras;
})(typeof window !== "undefined" ? window : globalThis);
