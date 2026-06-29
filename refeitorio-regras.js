// ============================================================
// refeitorio-regras.js — Camada central de regras do Refeitório Homy
// v: base-centralizada-v10-7-20260629
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
      _ausenciaEncerradaAguardandoMarcacao: true,
      Status: "Pendente",
      Confirmado: false,
      Opcao: Regras.getOpcao(base) || "principal",
      Nome_Prato: "Sem marcação",
      Origem: "Ausência encerrada - aguardando marcação",
      Observacao: "Ausência encerrada ou retorno automático legado. Colaborador liberado para escolher; Principal só será aplicado no travamento oficial se ninguém marcar."
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

  // Exporta para browser e para testes em Node, sem dependências.
  global.HomyRefeitorioRegras = Regras;
  if (typeof module !== "undefined" && module.exports) module.exports = Regras;
})(typeof window !== "undefined" ? window : globalThis);
