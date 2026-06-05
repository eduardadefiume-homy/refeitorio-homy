// admin-operacao-dia.js — Operação do Dia do Admin Homy

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
      const dia        = AdminUtils.getVal("operacaoDia") || AdminUtils.DIA_HOJE();
      const pedidos    = await SP.getPedidos(semanaId);
      const norm       = v => AdminUtils.norm(v);

      // Remove duplicatas de extra automático (mantém só 1 por dia)
      const seenAuto   = new Set();
      this._lista = pedidos.filter(p => {
        if (norm(SP.pick(p, "Dia")) !== norm(dia)) return false;
        if (SP.isExtraPedido(p)) {
          const k = `auto-${norm(SP.pick(p, "Dia"))}`;
          if (seenAuto.has(k)) return false;
          seenAuto.add(k);
        }
        return true;
      });

      this._renderTotais();
      this._renderTabela();
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</td></tr>`;
    }
  },

  _renderTotais() {
    const norm = v => AdminUtils.norm(v);
    const STATUS_PROD = ["confirmado", "extra", "aprovado"];
    const STATUS_CANC = ["cancelado", "afastado", "ferias", "nao vai almocar", "bloqueado", "travado"];
    const isConf = p => STATUS_PROD.includes(norm(SP.pick(p, "Status") || "")) || SP.isTrue(SP.pick(p, "Confirmado"));
    const isCanc = p => STATUS_CANC.includes(norm(SP.pick(p, "Status") || ""));

    AdminUtils.setTxt("opTotalConfirmado", this._lista.filter(isConf).length);
    AdminUtils.setTxt("opTotalPrincipal",  this._lista.filter(p => isConf(p) && norm(SP.pick(p, "Opcao")) === "principal").length);
    AdminUtils.setTxt("opTotalLight",      this._lista.filter(p => isConf(p) && norm(SP.pick(p, "Opcao")) === "light").length);
    AdminUtils.setTxt("opTotalCancelado",  this._lista.filter(isCanc).length);
  },

  _renderTabela() {
    const tbody      = document.getElementById("operacaoTable");
    if (!tbody) return;

    const statusFiltro = AdminUtils.norm(AdminUtils.getVal("operacaoFiltroStatus"));
    const busca        = AdminUtils.norm(AdminUtils.getVal("operacaoBusca"));
    const norm         = v => AdminUtils.norm(v);

    let lista = this._lista;
    if (statusFiltro) lista = lista.filter(p => norm(SP.pick(p, "Status") || "") === statusFiltro);
    if (busca) lista = lista.filter(p =>
      norm([SP.pick(p, "Colaborador_nome"), SP.pick(p, "Centro_Custo")].join(" ")).includes(busca)
    );

    if (!lista.length) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Nenhum pedido encontrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = lista.map(p => {
      const id     = AdminUtils.esc(p.id || "");
      const nome   = AdminUtils.esc(SP.pick(p, "Colaborador_nome") || "—");
      const cc     = AdminUtils.esc(SP.pick(p, "Centro_Custo")     || "—");
      const opcao  = AdminUtils.esc(SP.pick(p, "Opcao")            || "—");
      const prato  = AdminUtils.esc(SP.pick(p, "Nome_Prato")       || "—");
      const status = SP.pick(p, "Status") || "Pendente";
      const origem = AdminUtils.esc(SP.pick(p, "Origem", "tipo")   || "Refeitório");
      const isEx   = SP.isExtraPedido(p);
      return `<tr>
        <td${isEx ? ' style="color:#ffd36d;font-weight:700"' : ""}>${nome}</td>
        <td>${cc}</td>
        <td><span class="badge badge-blue">${opcao}</span></td>
        <td>${prato}</td>
        <td>${AdminUtils.badge(status)}</td>
        <td>${origem}</td>
        <td><div class="table-actions">
          <button class="btn-icon" title="Confirmar"     onclick="AdminOperacao.alterarStatus('${id}','Confirmado')">✅</button>
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
      AdminUtils.toast(`Status: ${status}`, "success");
      // Atualiza localmente sem recarregar tudo
      const p = this._lista.find(x => String(x.id) === String(id));
      if (p) { p.Status = status; p.Confirmado = ["Confirmado", "Extra"].includes(status); }
      this._renderTotais();
      this._renderTabela();
    } catch (e) {
      AdminUtils.toast("Erro: " + e.message, "error");
    }
  },

  async excluir(id) {
    if (!confirm("Excluir este pedido da operação?")) return;
    try {
      await SP.init();
      await SP.deletePedido(id);
      this._lista = this._lista.filter(p => String(p.id) !== String(id));
      this._renderTotais();
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

    // Inicializa dia com o dia de hoje
    const diaEl = document.getElementById("operacaoDia");
    if (diaEl && !diaEl.dataset.boundOp) diaEl.value = AdminUtils.DIA_HOJE();

    bind("operacaoDia",           "change", () => this._carregar(semanaId));
    bind("operacaoFiltroStatus",  "change", () => this._renderTabela());
    bind("operacaoBusca",         "input",  () => this._renderTabela());
    bind("btnRecarregarOperacao", "click",  () => this._carregar(semanaId));
  }
};
