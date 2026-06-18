// admin-operacao-dia.js — Operação do Dia do Admin Homy
// Correção: operação do dia deduplica pedidos, mostra ausências sem pedido espelho e sincroniza todos os extras/investigadores.

const AdminOperacao = window.AdminOperacao = {
  _lista: [],

  PORTARIA_CC: "120602 - PORTARIA",

  _CC_MAPA: {
    "110101":"DIRETORIA PRESIDENCIAL","110201":"DIRETORIA ADMINISTRATIVA",
    "110202":"DIRETORIA DE PRODUTOS","120101":"ADM GERAL","120102":"CUSTOS",
    "120103":"LEGALIZAÇÃO","120201":"CONTABILIDADE","120202":"FISCAL",
    "120301":"FINANCEIRO","120401":"RECURSOS HUMANOS","120402":"DEPARTAMENTO PESSOAL",
    "120501":"TI","120601":"RECEPÇÃO","120602":"PORTARIA",
    "120603":"ASSEIO E CONSERVAÇÃO","120604":"JARDINAGEM","150101":"SUPRIMENTOS",
    "160101":"CONTROLADORIA E COMPLIANCE","160102":"ADM CONTRATOS","170101":"SGI",
    "180101":"P&D","190101":"PATIO EXTERNO","220101":"ADM VENDAS",
    "220201":"COML INTERNO - SUPORTE","220202":"COML INTERNO - ATIVO",
    "220301":"COML EXTERNO - CLT","220302":"COML EXTERNO - REPRESENTANTE",
    "230101":"SUPORTE TECNICO INDUSTRIAL","230102":"SUPORTE TECNICO OBRAS/INFRA",
    "240101":"MARKETING","250101":"FATURAMENTO","250102":"LOGISTICA",
    "250103":"EXPEDIÇÃO","320101":"PRODUÇÃO","320201":"ENVASE MANUAL",
    "320202":"ENVASE AUTOMATICO","320301":"LABORATORIO E CONTROLE QUALIDADE",
    "360101":"APOIO A PRODUÇÃO","360102":"PCP","360201":"MANUTENÇÃO",
    "360301":"ALMOXARIFADO DE INSUMOS"
  },

  async load(semanaId) {
    this._bindControles(semanaId);
    await this._carregar(semanaId);
  },

  _norm(v) {
    return String(v || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
  },

  _pick(obj, ...keys) {
    for (const k of keys) {
      const v = SP.pick ? SP.pick(obj, k) : obj?.[k];
      if (v !== undefined && v !== null && String(v).trim() !== "") return v;
    }
    return "";
  },

  _esc(v) {
    return AdminUtils.esc ? AdminUtils.esc(v) : String(v ?? "");
  },

  _formatarCC(valor) {
    if (!valor) return "Sem setor";
    const v = String(valor).replace(/[–—]/g, " - ").trim();
    if (!v || v === "—") return "Sem setor";
    const match = v.match(/(\d{6})/);
    const codigo = match ? match[1] : "";
    if (codigo) return `${codigo} - ${this._CC_MAPA[codigo] || this._nomeNoValor(v) || "SETOR"}`;
    return v;
  },

  _nomeNoValor(valor) {
    const partes = String(valor || "").split(" - ").map(x => x.trim()).filter(Boolean);
    return partes.find(x => !/^\d{6}$/.test(x)) || "";
  },

  _isExtraPedido(p) {
    const origem = this._norm(this._pick(p, "Origem", "tipo", "Tipo"));
    const nome = this._norm(this._pick(p, "Colaborador_nome", "Nome", "Title"));
    return origem.includes("extra") || origem.includes("guarda") || origem.includes("investigador") ||
           origem.includes("visitante") || origem.includes("motorista") || origem.includes("prestador") ||
           origem.includes("marmita") || nome.includes("refeicao extra") || nome.includes("guarda") ||
           nome.includes("investigador");
  },

  _isExtraAutomatico(p) {
    const origem = this._norm(this._pick(p, "Origem", "tipo", "Tipo"));
    const nome = this._norm(this._pick(p, "Colaborador_nome", "Nome", "Title"));
    const obs = this._norm(this._pick(p, "Observacao", "Observação"));
    return nome.includes("refeicao extra") || origem.includes("extra automatica") || obs.includes("extra automatica");
  },

  _centroCustoPedido(p) {
    const raw = this._pick(p, "Centro_Custo", "CentroCusto", "Setor", "Departamento");
    const origem = this._norm(this._pick(p, "Origem", "tipo", "Tipo"));
    const nome = this._norm(this._pick(p, "Colaborador_nome", "Nome", "Title"));
    const obs = this._pick(p, "Observacao", "Observação");

    // Extras sem CC, guarda e refeições extras entram na Portaria.
    const extraPortaria = origem.includes("guarda") || origem.includes("extra") || nome.includes("guarda") || nome.includes("refeicao extra");
    if (extraPortaria && (!raw || this._norm(raw).includes("sem setor") || raw === "—")) return this.PORTARIA_CC;

    // Se a observação trouxe um CC explícito, usa ele.
    const ccObs = String(obs || "").match(/(\d{6})\s*-?\s*([A-Za-zÀ-ÿ\s/.-]+)?/);
    if ((!raw || raw === "—") && ccObs) return this._formatarCC(ccObs[0]);

    return this._formatarCC(raw || (extraPortaria ? this.PORTARIA_CC : "Sem setor"));
  },

  _pedidoTemConteudo(p) {
    const nome = this._pick(p, "Colaborador_nome", "Nome", "Title");
    const dia = this._pick(p, "Dia");
    const opcao = this._pick(p, "Opcao");
    const prato = this._pick(p, "Nome_Prato");
    return !!(String(nome || "").trim() || String(dia || "").trim() || String(opcao || "").trim() || String(prato || "").trim());
  },

  _dataHoraDoDia(semanaId, dia) {
    try {
      if (typeof SP.getDataRefBySemanaDia === "function") return `${SP.getDataRefBySemanaDia(semanaId, dia)}T12:00:00`;
    } catch (_) {}
    return new Date().toISOString();
  },

  _dateISO(v) {
    if (!v) return "";
    const s = String(v);
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    const d = new Date(s);
    return isNaN(d) ? "" : d.toISOString().slice(0, 10);
  },

  _dataPorDia(semanaId, dia) {
    try {
      if (typeof SP.getDataRefBySemanaDia === "function") return SP.getDataRefBySemanaDia(semanaId, dia);
    } catch (_) {}
    try {
      const datas = typeof SP.getWeekDates === "function" ? SP.getWeekDates(semanaId) : [];
      const idx = { segunda:0, terca:1, terça:1, quarta:2, quinta:3, sexta:4 }[this._norm(dia)];
      if (idx !== undefined && datas[idx]) return this._dateISO(datas[idx]);
    } catch (_) {}
    return "";
  },

  _motivoAusencia(a) {
    return this._pick(a, "Motivo", "motivo", "Status", "status") || "Ausente";
  },

  _formatarMotivoAusencia(motivo) {
    const n = this._norm(motivo);
    if (n === "nao_vai_almocar" || n === "nao vai almocar" || n === "não vai almoçar") return "Não vai almoçar";
    if (n === "ferias" || n === "férias") return "Férias";
    if (n === "afastado") return "Afastado";
    if (n === "atestado") return "Atestado";
    if (n === "licenca" || n === "licença") return "Licença";
    return String(motivo || "Ausente");
  },

  _ausenciaAtiva(a) {
    const ativo = this._pick(a, "Ativo", "ativo");
    const status = this._norm(this._pick(a, "Status", "status"));
    if (["inativo", "cancelado", "false", "nao", "não", "0"].includes(status)) return false;
    if (ativo === "") return true;
    return SP.isTrue ? SP.isTrue(ativo) : !!ativo;
  },

  _ausenciaInicio(a) {
    return this._dateISO(this._pick(a, "Data_Inicio", "Inicio", "DataInicio", "Data"));
  },

  _ausenciaFim(a) {
    return this._dateISO(this._pick(a, "Data_Fim", "Fim", "DataFim", "Data")) || this._ausenciaInicio(a);
  },

  _mesmoColaboradorPedidoAusencia(p, a) {
    const pId = String(this._pick(p, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
    const aId = String(this._pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
    if (pId && aId && pId === aId) return true;

    const pNome = this._norm(this._pick(p, "Colaborador_nome", "Colaborador", "Nome", "Title"));
    const aNome = this._norm(this._pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title"));
    return !!pNome && !!aNome && pNome === aNome;
  },

  _ausenciaDoPedidoNoDia(p, ausencias, semanaId, dia) {
    const dataRef = this._dataPorDia(semanaId, dia);
    if (!dataRef) return null;
    return (ausencias || []).find(a => {
      if (!this._ausenciaAtiva(a)) return false;
      if (!this._mesmoColaboradorPedidoAusencia(p, a)) return false;
      const ini = this._ausenciaInicio(a);
      const fim = this._ausenciaFim(a);
      return !!ini && !!fim && ini <= dataRef && fim >= dataRef;
    }) || null;
  },

  async _buscarAusencias() {
    const tentativas = [
      () => SP.getAusencias?.(true),
      () => SP.getAusenciasRefeitorio?.(),
      () => SP.getItems?.("Ausencias do Refeitorio"),
      () => SP.getItems?.("Ausências do Refeitório"),
      () => SP.getItems?.("Ausencias_Refeitorio"),
      () => SP.getItems?.("Ausências")
    ];
    for (const fn of tentativas) {
      try {
        const r = await fn();
        if (Array.isArray(r)) return r;
      } catch (_) {}
    }
    return [];
  },

  _isRefeicaoAdicionalColaborador(p) {
    const origem = this._norm(this._pick(p, "Origem", "origem", "tipo", "Tipo"));
    const obs = this._norm(this._pick(p, "Observacao", "Observação", "observacao"));
    const cid = this._norm(this._pick(p, "Colaborador_id", "ColaboradorId", "colaboradorId"));
    return origem.includes("segunda refeicao") || origem.includes("segunda refeição") ||
           origem.includes("transferencia para hoje") || origem.includes("transferência para hoje") ||
           origem.includes("refeicao adicional") || origem.includes("refeição adicional") ||
           obs.includes("adicionalid:") || obs.includes("colaboradorbaseid:") ||
           cid.includes("-adicional-");
  },

  _pedidoKeyOperacao(p) {
    const dia = this._norm(this._pick(p, "Dia", "dia"));
    const isEx = this._isExtraPedido(p);
    const obs = this._norm(this._pick(p, "Observacao", "Observação", "Obs"));
    const extraId = (obs.match(/extraid:\s*([^|\s]+)/i) || [])[1] || "";
    const idColab = String(this._pick(p, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
    const nome = this._norm(this._pick(p, "Colaborador_nome", "Nome", "Title"));
    const origem = this._norm(this._pick(p, "Origem", "tipo", "Tipo"));
    const opcao = this._norm(this._pick(p, "Opcao", "opcao"));

    if (this._isRefeicaoAdicionalColaborador(p)) return `adicional|${dia}|${this._pick(p, "id", "ID") || idColab || nome + "|" + origem + "|" + opcao}`;
    if (isEx) return `extra|${dia}|${extraId || idColab || nome + "|" + origem + "|" + opcao}`;
    return `colab|${dia}|${idColab || nome}`;
  },

  _timestampPedidoOperacao(p) {
    const raw = this._pick(p, "Modified", "modified", "Data_Hora", "DataHora", "Created", "created", "Data") || "";
    const dt = raw ? new Date(raw) : null;
    if (dt && !isNaN(dt)) return dt.getTime();
    const id = Number(this._pick(p, "id", "ID") || 0);
    return Number.isFinite(id) ? id : 0;
  },

  _prioridadePedidoOperacao(p) {
    const origem = this._norm(this._pick(p, "Origem", "tipo", "Tipo"));
    const status = this._norm(this._statusDisplay(p) || this._pick(p, "Status", "status"));
    let score = 0;
    if (this._pick(p, "id", "ID")) score += 2;
    if (!origem.includes("travamento")) score += 2;
    if (this._isAusenteOperacao(p) || ["nao vai almocar", "não vai almoçar", "ausente", "ferias", "férias", "afastado", "atestado"].includes(status)) score += 90;
    else if (["confirmado", "aprovado", "extra"].includes(status) || (SP.isTrue && SP.isTrue(this._pick(p, "Confirmado")))) score += 80;
    else if (status === "travado") score += 70;
    else if (["cancelado", "bloqueado"].includes(status)) score += 50;
    return score;
  },

  _compararPedidoOperacao(a, b) {
    // Quando existe mais de um Pedido para o mesmo colaborador/dia, a Operação
    // precisa exibir a versão mais recente. Esse era o motivo de alguns colaboradores
    // "sumirem" ou aparecerem com status antigo após novas marcações no Refeitório.
    const ta = this._timestampPedidoOperacao(a);
    const tb = this._timestampPedidoOperacao(b);
    if (ta !== tb) return ta - tb;
    return this._prioridadePedidoOperacao(a) - this._prioridadePedidoOperacao(b);
  },

  _deduplicarPedidosOperacao(lista) {
    const mapa = new Map();
    for (const p of lista || []) {
      const key = this._pedidoKeyOperacao(p);
      const atual = mapa.get(key);
      if (!atual || this._compararPedidoOperacao(p, atual) >= 0) mapa.set(key, p);
    }
    return Array.from(mapa.values());
  },

  _statusDisplay(p) {
    if (p?._ausenciaOperacao) return this._formatarMotivoAusencia(this._motivoAusencia(p._ausenciaOperacao));
    const status = this._pick(p, "Status") || "Pendente";
    const origem = this._norm(this._pick(p, "Origem", "tipo", "Tipo"));
    // Compatibilidade com pedidos antigos criados como Status Travado: para a cozinha/operação isso é pedido principal confirmado.
    if (this._norm(status) === "travado" && origem.includes("travamento")) return "Confirmado";
    return status;
  },

  _isAusenteOperacao(p) {
    const s = this._norm(this._statusDisplay(p));
    return ["ausente", "nao vai almocar", "não vai almoçar", "ferias", "férias", "afastado", "atestado", "licenca", "licença"].includes(s);
  },

  async _pratoPorOpcao(semanaId, dia, opcao) {
    try {
      if (typeof SP.getCardapio !== "function") return "Cardápio do Dia";
      const cardapio = await SP.getCardapio(semanaId);
      const item = (cardapio || []).find(c =>
        this._norm(this._pick(c, "Dia")) === this._norm(dia) &&
        this._norm(this._pick(c, "Opcao", "Opção")) === this._norm(opcao)
      );
      return this._pick(item, "Nome_Prato", "Descricao", "Descrição", "Title") || "Cardápio do Dia";
    } catch (_) {
      return "Cardápio do Dia";
    }
  },

  _extrairCentroExtra(extra) {
    const raw = this._pick(extra, "Centro_Custo", "CentroCusto", "Setor", "Departamento");
    if (raw) return this._formatarCC(raw);
    const obs = this._pick(extra, "Observacao", "Observação");
    const m = String(obs || "").match(/(\d{6})(?:\s*-\s*([A-Za-zÀ-ÿ\s/.-]+))?/);
    if (m) return this._formatarCC(m[0]);
    return this.PORTARIA_CC;
  },

  _extraDia(extra) {
    return this._pick(extra, "Dia") || "";
  },

  _extraNome(extra) {
    return this._pick(extra, "Nome", "Title") || "Refeição Extra";
  },

  _extraTipo(extra) {
    return this._pick(extra, "tipo", "Tipo", "Origem") || "Extra";
  },

  _extraOpcao(extra) {
    return this._pick(extra, "Opcao", "Opção") || "principal";
  },

  _pedidoCorrespondeExtra(p, extra, dia) {
    if (this._norm(this._pick(p, "Dia")) !== this._norm(dia)) return false;
    const nome = this._norm(this._pick(p, "Colaborador_nome", "Nome", "Title"));
    if (!nome) return false;

    const extraId = String(extra?.id || extra?.ID || "").trim();
    const colabId = String(this._pick(p, "Colaborador_id", "colaboradorId", "ColaboradorId") || "").trim();
    const extraNome = this._norm(this._extraNome(extra));
    const obs = this._norm(this._pick(p, "Observacao", "Observação", "Obs"));

    // vínculo seguro: ExtraID na observação ou colaborador_id extra-<id>.
    if (extraId && (obs.includes(`extraid:${this._norm(extraId)}`) || colabId === `extra-${extraId}`)) return true;

    // Sem id, só considera o mesmo extra por nome EXATO.
    // Não usar apenas tipo "investigador", pois isso fazia Investigador 2/3 sobrescrever o 1.
    if (extraNome && nome === extraNome) return true;
    return false;
  },

  async _criarOuAtualizarPedidoDoExtra(semanaId, dia, extra, pedidos) {
    const nome = this._extraNome(extra);
    const tipo = this._extraTipo(extra);
    const opcao = this._extraOpcao(extra);
    const centroCusto = this._extrairCentroExtra(extra);
    const nomePrato = await this._pratoPorOpcao(semanaId, dia, opcao);
    const obsBase = this._pick(extra, "Observacao", "Observação") || "Criado a partir do módulo Extras";
    const observacao = [obsBase, extra.id ? `ExtraID:${extra.id}` : ""].filter(Boolean).join(" | ");
    const dataHora = this._dataHoraDoDia(semanaId, dia);
    const colabId = extra.id ? `extra-${extra.id}` : `extra-${this._norm(tipo)}-${this._norm(nome)}-${this._norm(dia)}`;

    const existente = (pedidos || []).find(p => this._pedidoCorrespondeExtra(p, extra, dia));

    const campos = {
      Colaborador_id: colabId,
      colaboradorId: colabId,
      Colaborador_nome: nome,
      colaboradorNome: nome,
      Dia: dia,
      dia,
      Opcao: opcao,
      opcao,
      Nome_Prato: nomePrato,
      nomePrato,
      Confirmado: true,
      confirmado: true,
      Data_Hora: dataHora,
      dataHora,
      Centro_Custo: centroCusto,
      centroCusto,
      Status: "Confirmado",
      status: "Confirmado",
      Observacao: observacao,
      observacao,
      Origem: tipo,
      origem: tipo,
      Alterado_Por: SP.getUserName ? SP.getUserName() : "Admin",
      alteradoPor: SP.getUserName ? SP.getUserName() : "Admin"
    };

    if (existente?.id && typeof SP.updatePedido === "function") {
      await SP.updatePedido(existente.id, campos).catch(e => console.warn("[Operação] Falha ao atualizar pedido espelho", e));
      return { ...existente, ...campos, id: existente.id };
    }

    try {
      if (typeof SP.savePedido === "function") {
        const r = await SP.savePedido(semanaId, colabId, nome, dia, opcao, nomePrato, {
          confirmado: true,
          status: "Confirmado",
          dataHora,
          centroCusto,
          origem: tipo,
          observacao,
          alteradoPor: SP.getUserName ? SP.getUserName() : "Admin"
        });
        return { id: r?.id || "", Semana_id: semanaId, Colaborador_id: colabId, Colaborador_nome: nome, Dia: dia, Opcao: opcao, Nome_Prato: nomePrato, Confirmado: true, Data_Hora: dataHora, Centro_Custo: centroCusto, Status: "Confirmado", Observacao: observacao, Origem: tipo };
      }

      const r = await SP.createItem("Pedidos", {
        Title: `${semanaId}-${colabId}-${dia}`,
        Semana_id: semanaId,
        Colaborador_id: colabId,
        Colaborador_nome: nome,
        Dia: dia,
        Opcao: opcao,
        Nome_Prato: nomePrato,
        Confirmado: true,
        Data_Hora: dataHora,
        Centro_Custo: centroCusto,
        Status: "Confirmado",
        Observacao: observacao,
        Origem: tipo,
        Alterado_Por: SP.getUserName ? SP.getUserName() : "Admin"
      });
      return { id: r?.id || "", Semana_id: semanaId, Colaborador_id: colabId, Colaborador_nome: nome, Dia: dia, Opcao: opcao, Nome_Prato: nomePrato, Confirmado: true, Data_Hora: dataHora, Centro_Custo: centroCusto, Status: "Confirmado", Observacao: observacao, Origem: tipo };
    } catch (e) {
      console.warn("[Operação] Não foi possível criar pedido espelho; usando linha virtual.", e);
      return { _virtualExtra: true, Semana_id: semanaId, Colaborador_id: colabId, Colaborador_nome: nome, Dia: dia, Opcao: opcao, Nome_Prato: nomePrato, Confirmado: true, Data_Hora: dataHora, Centro_Custo: centroCusto, Status: "Confirmado", Observacao: observacao, Origem: tipo };
    }
  },

  async _sincronizarExtrasParaPedidos(semanaId, dia, pedidos) {
    if (typeof SP.getExtras !== "function") return pedidos;

    const todosExtras = await SP.getExtras(semanaId).catch(() => []);
    const extrasDoDia = (todosExtras || []).filter(e => {
      const status = this._norm(this._pick(e, "Status", "status"));
      const ativo = this._pick(e, "Ativo", "ativo");
      const inativo = ["cancelado", "bloqueado", "excluido", "excluído", "inativo"].includes(status) || (ativo !== "" && SP.isTrue && !SP.isTrue(ativo));
      return this._norm(this._extraDia(e)) === this._norm(dia) && !inativo;
    });
    if (!extrasDoDia.length) return pedidos;

    const resultado = [...(pedidos || [])];

    for (const extra of extrasDoDia) {
      const espelho = await this._criarOuAtualizarPedidoDoExtra(semanaId, dia, extra, resultado);
      const idx = resultado.findIndex(p => this._pedidoCorrespondeExtra(p, extra, dia));
      if (idx >= 0) resultado[idx] = { ...resultado[idx], ...espelho };
      else resultado.push(espelho);
    }

    return resultado;
  },


  _incluirAusenciasSemPedido(lista, ausencias, semanaId, dia) {
    const dataRef = this._dataPorDia(semanaId, dia);
    if (!dataRef) return lista || [];

    const resultado = [...(lista || [])];

    for (const a of (ausencias || [])) {
      if (!this._ausenciaAtiva(a)) continue;

      const ini = this._ausenciaInicio(a);
      const fim = this._ausenciaFim(a);
      if (!ini || !fim || ini > dataRef || fim < dataRef) continue;

      const jaExiste = resultado.some(p => this._mesmoColaboradorPedidoAusencia(p, a));
      if (jaExiste) continue;

      const colabId = String(this._pick(a, "Colaborador_id", "ColaboradorId", "colaboradorId") || "").trim();
      const nome = this._pick(a, "Colaborador_nome", "Colaborador", "Nome", "Title") || "Colaborador ausente";
      const motivo = this._formatarMotivoAusencia(this._motivoAusencia(a));
      const centroCusto = this._formatarCC(this._pick(a, "Centro_Custo", "CentroCusto", "Setor", "Departamento") || "Sem setor");
      const idAus = String(this._pick(a, "id", "ID") || `${this._norm(nome)}-${dataRef}`);

      resultado.push({
        id: `ausencia-${idAus}-${this._norm(dia)}`,
        _virtualAusencia: true,
        _ausenciaOperacao: a,
        Semana_id: semanaId,
        Colaborador_id: colabId,
        Colaborador_nome: nome,
        Dia: dia,
        Opcao: "principal",
        Nome_Prato: motivo,
        Confirmado: false,
        Data_Hora: `${dataRef}T12:00:00`,
        Centro_Custo: centroCusto,
        Status: motivo,
        Observacao: this._pick(a, "Observacao", "Observação", "Obs") || "Ausência cadastrada sem pedido espelho.",
        Origem: "Ausência"
      });
    }

    return resultado;
  },

  async _carregar(semanaId) {
    const tbody = document.getElementById("operacaoTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      const dia = AdminUtils.getVal("operacaoDia") || AdminUtils.DIA_HOJE();
      let [pedidos, ausencias] = await Promise.all([
        SP.getPedidos(semanaId),
        this._buscarAusencias()
      ]);
      pedidos = await this._sincronizarExtrasParaPedidos(semanaId, dia, pedidos);

      const seenAuto = new Set();
      let lista = (pedidos || []).filter(p => {
        if (!this._pedidoTemConteudo(p)) return false;
        if (this._norm(this._pick(p, "Dia")) !== this._norm(dia)) return false;

        // Deduplica só a refeição extra automática. Guarda/prestador/investigador continuam visíveis.
        if (this._isExtraAutomatico(p)) {
          const k = `auto-${this._norm(dia)}`;
          if (seenAuto.has(k)) return false;
          seenAuto.add(k);
        }
        return true;
      });

      lista = this._deduplicarPedidosOperacao(lista).map(p => {
        const aus = this._ausenciaDoPedidoNoDia(p, ausencias, semanaId, dia);
        return aus ? { ...p, _ausenciaOperacao: aus } : p;
      });

      // Se a ausência foi cadastrada no módulo Ausências ou no Marcar Refeição
      // mas ainda não existe um pedido espelho, a Operação do Dia também precisa mostrar.
      lista = this._incluirAusenciasSemPedido(lista, ausencias, semanaId, dia);

      this._lista = lista;

      this._renderTotais(dia);
      this._renderTabela();
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell" style="color:#ff8080">Erro: ${this._esc(e.message || e)}</td></tr>`;
    }
  },

  _renderTotais(dia) {
    const STATUS_PROD = ["confirmado", "extra", "aprovado"];
    const STATUS_CANC = ["cancelado", "afastado", "ferias", "férias", "nao vai almocar", "não vai almoçar", "bloqueado", "atestado", "ausente", "licenca", "licença"];

    const isCanc = p => STATUS_CANC.includes(this._norm(this._statusDisplay(p) || ""));
    const isConf = p => !isCanc(p) && (STATUS_PROD.includes(this._norm(this._statusDisplay(p) || "")) || (SP.isTrue && SP.isTrue(this._pick(p, "Confirmado"))));
    const conf = this._lista.filter(isConf);

    const setCard = (id, val) => AdminUtils.setTxt(id, val);
    setCard("opTotalConfirmado", conf.length);
    setCard("opTotalPrincipal", conf.filter(p => this._norm(this._pick(p, "Opcao")) === "principal").length);
    setCard("opTotalLight", conf.filter(p => this._norm(this._pick(p, "Opcao")) === "light").length);
    setCard("opTotalCarne", conf.filter(p => this._norm(this._pick(p, "Opcao")) === "carne").length);
    setCard("opTotalMassa", conf.filter(p => this._norm(this._pick(p, "Opcao")) === "massa").length);
    setCard("opTotalCancelado", this._lista.filter(isCanc).length);

    const cardLanche = document.getElementById("cardOpLanche");
    if (cardLanche) {
      cardLanche.style.display = this._norm(dia) === "sexta" ? "" : "none";
      setCard("opTotalLanche", conf.filter(p => this._norm(this._pick(p, "Opcao")) === "lanche").length);
    }
  },

  _renderTabela() {
    const tbody = document.getElementById("operacaoTable");
    if (!tbody) return;

    const statusFiltro = this._norm(AdminUtils.getVal("operacaoFiltroStatus"));
    const busca = this._norm(AdminUtils.getVal("operacaoBusca"));

    let lista = this._lista;
    if (statusFiltro) lista = lista.filter(p => this._norm(this._statusDisplay(p) || this._pick(p, "Status") || "") === statusFiltro);
    if (busca) lista = lista.filter(p => this._norm([
      this._pick(p, "Colaborador_nome", "Title", "Nome"),
      this._centroCustoPedido(p),
      this._pick(p, "Origem", "tipo", "Tipo")
    ].join(" ")).includes(busca));

    lista = lista.slice().sort((a, b) => {
      const na = String(this._pick(a, "Colaborador_nome", "Title", "Nome") || "");
      const nb = String(this._pick(b, "Colaborador_nome", "Title", "Nome") || "");
      return na.localeCompare(nb, "pt-BR", { sensitivity:"base", numeric:true });
    });

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Nenhum pedido encontrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.map(p => {
      const id = this._esc(p.id || "");
      const nome = this._esc(this._pick(p, "Colaborador_nome", "Title", "Nome") || "—");
      const cc = this._esc(this._centroCustoPedido(p));
      const opcao = this._esc(this._pick(p, "Opcao") || "—");
      const prato = this._esc(this._pick(p, "Nome_Prato") || "—");
      const status = this._statusDisplay(p);
      const origem = this._esc(this._pick(p, "Origem", "tipo", "Tipo") || "Refeitório");
      const isEx = this._isExtraPedido(p);
      const isAus = this._isAusenteOperacao(p);
      const disabled = p._virtualExtra ? "disabled title='Extra sem pedido espelho no SharePoint'" : "";
      const nomeStyle = isAus ? ' style="color:#ff9a90;font-weight:700"' : (isEx ? ' style="color:#ffd36d;font-weight:700"' : "");

      return `<tr${isAus ? ' style="background:rgba(192,40,28,.055)"' : ""}>
        <td${nomeStyle}>${nome}</td>
        <td>${cc}</td>
        <td><span class="badge badge-blue">${opcao}</span></td>
        <td>${prato}</td>
        <td>${AdminUtils.badge(status)}</td>
        <td>${origem}</td>
        <td><div class="table-actions">
          <button class="btn-icon" ${disabled} title="Confirmar" onclick="AdminOperacao.alterarStatus('${id}','Confirmado')">✅</button>
          <button class="btn-icon danger" ${disabled} title="Cancelar" onclick="AdminOperacao.alterarStatus('${id}','Cancelado')">❌</button>
          <button class="btn-icon" ${disabled} title="Não vai almoçar" onclick="AdminOperacao.alterarStatus('${id}','Não vai almoçar')">🚫</button>
          ${(!isEx && !p._virtualAusencia && !p._virtualExtra) ? `<button class="btn-icon" ${disabled} title="Refeição adicional" onclick="AdminOperacao.abrirAdicional('${id}')">➕</button>` : ""}
          <button class="btn-icon danger" ${disabled} title="Excluir" onclick="AdminOperacao.excluir('${id}')">🗑️</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  _diaSeguinteOperacao(dia) {
    const dias = ["segunda", "terca", "quarta", "quinta", "sexta"];
    const idx = dias.indexOf(this._norm(dia));
    return dias[Math.min((idx < 0 ? 0 : idx + 1), dias.length - 1)] || "sexta";
  },

  _labelDiaOperacao(dia) {
    const labels = { segunda:"Segunda", terca:"Terça", terça:"Terça", quarta:"Quarta", quinta:"Quinta", sexta:"Sexta" };
    return labels[this._norm(dia)] || dia || "—";
  },

  _opcoesRefeicaoHtml(dia) {
    const ops = [
      ["principal", "Principal"],
      ["light", "Light"],
      ["carne", "Carne"],
      ["massa", "Massa"]
    ];
    if (this._norm(dia) === "sexta") ops.push(["lanche", "Lanche"]);
    return ops.map(([v, l]) => `<option value="${v}">${l}</option>`).join("");
  },

  _ensureModalAdicional() {
    let modal = document.getElementById("modalOperacaoAdicional");
    if (modal) return modal;

    modal = document.createElement("div");
    modal.id = "modalOperacaoAdicional";
    modal.style.cssText = "display:none;position:fixed;inset:0;z-index:9999;background:rgba(3,8,20,.78);backdrop-filter:blur(8px);align-items:center;justify-content:center;padding:1rem";
    modal.innerHTML = `
      <div style="width:min(620px,96vw);background:#0b1a35;border:1px solid rgba(255,255,255,.14);border-radius:18px;box-shadow:0 24px 80px rgba(0,0,0,.65);overflow:hidden">
        <div style="display:flex;align-items:center;justify-content:space-between;padding:1rem 1.2rem;border-bottom:1px solid rgba(255,255,255,.08)">
          <div>
            <div style="font-family:'Barlow Condensed',sans-serif;font-size:1.35rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#fff">➕ Refeição adicional</div>
            <div id="opAdicionalSub" style="font-size:.78rem;color:rgba(143,170,210,.65);margin-top:2px">Ajuste vinculado ao colaborador</div>
          </div>
          <button type="button" id="opAdicionalFechar" class="btn-secondary" style="padding:.45rem .7rem">✕</button>
        </div>
        <div style="padding:1rem 1.2rem;display:grid;grid-template-columns:1fr 1fr;gap:.9rem">
          <input type="hidden" id="opAdicionalPedidoId">
          <div class="form-group" style="grid-column:1/-1">
            <label class="form-label">Colaborador</label>
            <input class="form-input" id="opAdicionalColaborador" disabled>
          </div>
          <div class="form-group" style="grid-column:1/-1">
            <label class="form-label">Tipo de ajuste</label>
            <select class="form-select" id="opAdicionalTipo">
              <option value="segunda">Segunda refeição para o colaborador</option>
              <option value="transferir">Transferir refeição para hoje</option>
            </select>
          </div>
          <div class="form-group">
            <label class="form-label">Dia que receberá a refeição</label>
            <select class="form-select" id="opAdicionalDiaDestino">
              <option value="segunda">Segunda</option>
              <option value="terca">Terça</option>
              <option value="quarta">Quarta</option>
              <option value="quinta">Quinta</option>
              <option value="sexta">Sexta</option>
            </select>
          </div>
          <div class="form-group">
            <label class="form-label">Opção da refeição adicional</label>
            <select class="form-select" id="opAdicionalOpcao"></select>
          </div>
          <div class="form-group" id="opAdicionalDiaOrigemBox" style="grid-column:1/-1;display:none">
            <label class="form-label">Dia que será removido / marcado como não vai almoçar</label>
            <select class="form-select" id="opAdicionalDiaOrigem">
              <option value="segunda">Segunda</option>
              <option value="terca">Terça</option>
              <option value="quarta">Quarta</option>
              <option value="quinta">Quinta</option>
              <option value="sexta">Sexta</option>
            </select>
          </div>
          <div class="form-group" style="grid-column:1/-1">
            <label class="form-label">Observação</label>
            <textarea class="form-textarea" id="opAdicionalObs" placeholder="Ex.: colaboradora pediu para trazer a refeição de amanhã para hoje"></textarea>
          </div>
          <div id="opAdicionalAviso" style="grid-column:1/-1;font-size:.8rem;color:#ffcf8a;background:rgba(255,180,0,.07);border:1px solid rgba(255,180,0,.18);border-radius:10px;padding:.7rem .8rem;line-height:1.45"></div>
        </div>
        <div style="display:flex;justify-content:flex-end;gap:.6rem;padding:1rem 1.2rem;border-top:1px solid rgba(255,255,255,.08)">
          <button type="button" id="opAdicionalCancelar" class="btn-secondary">Cancelar</button>
          <button type="button" id="opAdicionalSalvar" class="btn-success">💾 Salvar ajuste</button>
        </div>
      </div>`;
    document.body.appendChild(modal);

    const fechar = () => { modal.style.display = "none"; };
    document.getElementById("opAdicionalFechar")?.addEventListener("click", fechar);
    document.getElementById("opAdicionalCancelar")?.addEventListener("click", fechar);
    document.getElementById("opAdicionalTipo")?.addEventListener("change", () => this._atualizarModalAdicional());
    document.getElementById("opAdicionalDiaDestino")?.addEventListener("change", () => this._atualizarModalAdicional());
    document.getElementById("opAdicionalSalvar")?.addEventListener("click", () => this.salvarAdicional());
    return modal;
  },

  _atualizarModalAdicional() {
    const tipo = AdminUtils.getVal("opAdicionalTipo") || "segunda";
    const diaDestino = AdminUtils.getVal("opAdicionalDiaDestino") || AdminUtils.getVal("operacaoDia") || AdminUtils.DIA_HOJE();
    const boxOrigem = document.getElementById("opAdicionalDiaOrigemBox");
    const opcao = document.getElementById("opAdicionalOpcao");
    const aviso = document.getElementById("opAdicionalAviso");

    if (boxOrigem) boxOrigem.style.display = tipo === "transferir" ? "" : "none";
    if (opcao) opcao.innerHTML = this._opcoesRefeicaoHtml(diaDestino);

    if (aviso) {
      aviso.innerHTML = tipo === "transferir"
        ? `Será criada uma refeição adicional para o colaborador em <b>${this._labelDiaOperacao(diaDestino)}</b> e o dia removido será marcado como <b>Não vai almoçar</b>. O relatório continuará no centro de custo do colaborador.`
        : `Será criada uma segunda refeição para o colaborador em <b>${this._labelDiaOperacao(diaDestino)}</b>. O relatório contará 2 refeições no nome e centro de custo dele.`;
    }
  },

  abrirAdicional(id) {
    const pedido = this._lista.find(p => String(p.id) === String(id));
    if (!pedido) { AdminUtils.toast("Pedido não encontrado para ajuste adicional.", "error"); return; }
    if (this._isExtraPedido(pedido) || pedido._virtualAusencia || pedido._virtualExtra) {
      AdminUtils.toast("Use essa função apenas para colaboradores do quadro.", "error");
      return;
    }

    const modal = this._ensureModalAdicional();
    const diaAtual = AdminUtils.getVal("operacaoDia") || this._pick(pedido, "Dia") || AdminUtils.DIA_HOJE();
    AdminUtils.setVal("opAdicionalPedidoId", id);
    AdminUtils.setVal("opAdicionalColaborador", this._pick(pedido, "Colaborador_nome", "Nome", "Title") || "—");
    AdminUtils.setVal("opAdicionalTipo", "segunda");
    AdminUtils.setVal("opAdicionalDiaDestino", this._norm(diaAtual));
    AdminUtils.setVal("opAdicionalDiaOrigem", this._diaSeguinteOperacao(diaAtual));
    AdminUtils.setVal("opAdicionalObs", "");
    this._atualizarModalAdicional();
    modal.style.display = "flex";
  },

  async salvarAdicional() {
    const id = AdminUtils.getVal("opAdicionalPedidoId");
    const pedido = this._lista.find(p => String(p.id) === String(id));
    if (!pedido) { AdminUtils.toast("Pedido não encontrado.", "error"); return; }

    const tipoAjuste = AdminUtils.getVal("opAdicionalTipo") || "segunda";
    const diaDestino = AdminUtils.getVal("opAdicionalDiaDestino") || AdminUtils.getVal("operacaoDia") || this._pick(pedido, "Dia");
    const opcaoDestino = AdminUtils.getVal("opAdicionalOpcao") || "principal";
    const diaOrigem = AdminUtils.getVal("opAdicionalDiaOrigem") || "";
    const observacao = AdminUtils.getVal("opAdicionalObs") || "";

    if (tipoAjuste === "transferir" && this._norm(diaOrigem) === this._norm(diaDestino)) {
      AdminUtils.toast("Na transferência, o dia removido precisa ser diferente do dia que recebe a refeição.", "error");
      return;
    }

    const btn = document.getElementById("opAdicionalSalvar");
    const old = btn?.textContent;
    if (btn) { btn.disabled = true; btn.textContent = "⏳ Salvando..."; }

    try {
      await SP.init();
      if (typeof SP.criarRefeicaoAdicionalColaborador !== "function") {
        throw new Error("sharepoint.js não possui criarRefeicaoAdicionalColaborador.");
      }
      await SP.criarRefeicaoAdicionalColaborador({
        semanaId: AdminState.getSemanaId(),
        pedidoBase: pedido,
        pedidoId: id,
        tipoAjuste,
        diaDestino,
        opcaoDestino,
        diaOrigem,
        observacao
      });
      document.getElementById("modalOperacaoAdicional").style.display = "none";
      await this._carregar(AdminState.getSemanaId());
      AdminUtils.toast(tipoAjuste === "transferir" ? "Transferência criada." : "Segunda refeição criada.", "success");
    } catch (e) {
      AdminUtils.toast("Erro ao salvar adicional: " + (e.message || e), "error");
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = old || "💾 Salvar ajuste"; }
    }
  },

  async alterarStatus(id, status) {
    if (!id) return;
    const pedidoAtual = this._lista.find(x => String(x.id) === String(id)) || {};
    try {
      await SP.init();

      // Operação do Dia é a tela de ajuste final:
      // altera o Pedido base e sincroniza a Ausência do mesmo colaborador/dia.
      if (typeof SP.alterarPedidoOperacao === "function") {
        await SP.alterarPedidoOperacao(id, status, pedidoAtual);
      } else {
        await SP.updatePedido(id, {
          Status: status,
          Confirmado: ["Confirmado", "Extra"].includes(status),
          Alterado_Por: SP.getUserName()
        });
      }

      await this._carregar(AdminState.getSemanaId());
      AdminUtils.toast(`Status: ${status}`, "success");
    } catch (e) {
      AdminUtils.toast("Erro: " + (e.message || e), "error");
    }
  },

  async excluir(id) {
    if (!id) return;
    if (!confirm("Excluir este pedido?")) return;
    try {
      await SP.init();
      const pedidoAtual = this._lista.find(p => String(p.id) === String(id));
      if (pedidoAtual?._virtualAusencia) {
        const ausId = this._pick(pedidoAtual._ausenciaOperacao, "id", "ID");
        if (ausId && typeof SP.updateAusencia === "function") await SP.updateAusencia(ausId, { Ativo: false });
      } else {
        await SP.deletePedido(id);
      }
      await this._carregar(AdminState.getSemanaId());
      AdminUtils.toast("Pedido excluído/inativado.", "success");
    } catch (e) {
      AdminUtils.toast("Erro: " + (e.message || e), "error");
    }
  },

  _bindControles(semanaId) {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundOp) { el.dataset.boundOp = "1"; el.addEventListener(ev, fn); }
    };

    const diaEl = document.getElementById("operacaoDia");
    if (diaEl && !diaEl.dataset.boundOp && !diaEl.value) diaEl.value = AdminUtils.DIA_HOJE();

    bind("operacaoDia", "change", () => this._carregar(AdminState.getSemanaId()));
    bind("operacaoFiltroStatus", "change", () => this._renderTabela());
    bind("operacaoBusca", "input", () => this._renderTabela());
    bind("btnRecarregarOperacao", "click", () => this._carregar(AdminState.getSemanaId()));
  }
};
