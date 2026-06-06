// admin-operacao-dia.js — Operação do Dia do Admin Homy
// Cards mostram todas as opções da semana (Principal, Light, Carne, Massa)
// + Lanche SOMENTE na sexta-feira

const AdminOperacao = window.AdminOperacao = {

  _lista: [],

  async load(semanaId) {
    this._bindControles(semanaId);
    await this._carregar(semanaId);
  },

  async _carregar(semanaId) {
    const tbody = document.getElementById("operacaoTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      const dia     = AdminUtils.getVal("operacaoDia") || AdminUtils.DIA_HOJE();
      const pedidos = await SP.getPedidos(semanaId);
      const norm    = v => AdminUtils.norm(v);

      // Remove duplicatas de extra automático por dia
      const seenAuto = new Set();
      this._lista = pedidos.filter(p => {
        if (norm(SP.pick(p, "Dia")) !== norm(dia)) return false;
        if (SP.isExtraPedido(p)) {
          const k = `auto-${norm(dia)}`;
          if (seenAuto.has(k)) return false;
          seenAuto.add(k);
        }
        return true;
      });

      this._renderTotais(dia);
      this._renderTabela();
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</td></tr>`;
    }
  },

  _renderTotais(dia) {
    const norm = v => AdminUtils.norm(v);
    const STATUS_PROD = ["confirmado", "extra", "aprovado"];
    const STATUS_CANC = ["cancelado", "afastado", "ferias", "nao vai almocar", "bloqueado", "travado"];

    const isConf = p => STATUS_PROD.includes(norm(SP.pick(p, "Status") || "")) || SP.isTrue(SP.pick(p, "Confirmado"));
    const isCanc = p => STATUS_CANC.includes(norm(SP.pick(p, "Status") || ""));
    const conf   = this._lista.filter(isConf);

    // Cards fixos (sempre aparecem)
    const setCard = (id, val) => AdminUtils.setTxt(id, val);
    setCard("opTotalConfirmado", conf.length);
    setCard("opTotalPrincipal",  conf.filter(p => norm(SP.pick(p, "Opcao")) === "principal").length);
    setCard("opTotalLight",      conf.filter(p => norm(SP.pick(p, "Opcao")) === "light").length);
    setCard("opTotalCarne",      conf.filter(p => norm(SP.pick(p, "Opcao")) === "carne").length);
    setCard("opTotalMassa",      conf.filter(p => norm(SP.pick(p, "Opcao")) === "massa").length);
    setCard("opTotalCancelado",  this._lista.filter(isCanc).length);

    // Card de Lanche: só aparece na sexta-feira
    const cardLanche = document.getElementById("cardOpLanche");
    if (cardLanche) {
      cardLanche.style.display = norm(dia) === "sexta" ? "" : "none";
      setCard("opTotalLanche", conf.filter(p => norm(SP.pick(p, "Opcao")) === "lanche").length);
    }
  },

  _renderTabela() {
    const tbody = document.getElementById("operacaoTable");
    if (!tbody) return;

    const statusFiltro = AdminUtils.norm(AdminUtils.getVal("operacaoFiltroStatus"));
    const busca        = AdminUtils.norm(AdminUtils.getVal("operacaoBusca"));
    const norm         = v => AdminUtils.norm(v);

    let lista = this._lista;
    if (statusFiltro) lista = lista.filter(p => norm(SP.pick(p, "Status") || "") === statusFiltro);
    if (busca) lista = lista.filter(p =>
      norm([SP.pick(p, "Colaborador_nome", "Title"), SP.pick(p, "Centro_Custo")].join(" ")).includes(busca)
    );

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Nenhum pedido encontrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.map(p => {
      const id     = AdminUtils.esc(p.id || "");
      const nome   = AdminUtils.esc(SP.pick(p, "Colaborador_nome", "Title") || "—");
      const cc     = AdminUtils.esc(SP.pick(p, "Centro_Custo")              || "—");
      const opcao  = AdminUtils.esc(SP.pick(p, "Opcao")                     || "—");
      const prato  = AdminUtils.esc(SP.pick(p, "Nome_Prato")                || "—");
      const status = SP.pick(p, "Status") || "Pendente";
      const origem = AdminUtils.esc(SP.pick(p, "Origem", "tipo")            || "Refeitório");
      const isEx   = SP.isExtraPedido(p);
      return `<tr>
        <td${isEx ? ' style="color:#ffd36d;font-weight:700"' : ""}>${nome}</td>
        <td>${cc}</td>
        <td><span class="badge badge-blue">${opcao}</span></td>
        <td>${prato}</td>
        <td>${AdminUtils.badge(status)}</td>
        <td>${origem}</td>
        <td><div class="table-actions">
          <button class="btn-icon" title="Confirmar"      onclick="AdminOperacao.alterarStatus('${id}','Confirmado')">✅</button>
          <button class="btn-icon danger" title="Cancelar" onclick="AdminOperacao.alterarStatus('${id}','Cancelado')">❌</button>
          <button class="btn-icon" title="Não vai almoçar" onclick="AdminOperacao.alterarStatus('${id}','Não vai almoçar')">🚫</button>
          <button class="btn-icon danger" title="Excluir"  onclick="AdminOperacao.excluir('${id}')">🗑️</button>
        </div></td>
      </tr>`;
    }).join("");
  },

  async alterarStatus(id, status) {
    if (!id) return;
    try {
      await SP.init();
      await SP.updatePedido(id, {
        Status:       status,
        Confirmado:   ["Confirmado", "Extra"].includes(status),
        Alterado_Por: SP.getUserName(),
        Origem:       "Admin"
      });
      const p = this._lista.find(x => String(x.id) === String(id));
      if (p) { p.Status = status; p.Confirmado = ["Confirmado", "Extra"].includes(status); }
      const dia = AdminUtils.getVal("operacaoDia") || AdminUtils.DIA_HOJE();
      this._renderTotais(dia);
      this._renderTabela();
      AdminUtils.toast(`Status: ${status}`, "success");
    } catch (e) {
      AdminUtils.toast("Erro: " + e.message, "error");
    }
  },

  async excluir(id) {
    if (!confirm("Excluir este pedido?")) return;
    try {
      await SP.init();
      await SP.deletePedido(id);
      this._lista = this._lista.filter(p => String(p.id) !== String(id));
      const dia = AdminUtils.getVal("operacaoDia") || AdminUtils.DIA_HOJE();
      this._renderTotais(dia);
      this._renderTabela();
      AdminUtils.toast("Pedido excluído.", "success");
    } catch (e) {
      AdminUtils.toast("Erro: " + e.message, "error");
    }
  },

  _bindControles(semanaId) {
    const bind = (id, ev, fn) => {
      const el = document.getElementById(id);
      if (el && !el.dataset.boundOp) { el.dataset.boundOp = "1"; el.addEventListener(ev, fn); }
    };

    const diaEl = document.getElementById("operacaoDia");
    if (diaEl && !diaEl.dataset.boundOp) diaEl.value = AdminUtils.DIA_HOJE();

    bind("operacaoDia",           "change", () => this._carregar(semanaId));
    bind("operacaoFiltroStatus",  "change", () => this._renderTabela());
    bind("operacaoBusca",         "input",  () => this._renderTabela());
    bind("btnRecarregarOperacao", "click",  () => this._carregar(semanaId));
  }
};
