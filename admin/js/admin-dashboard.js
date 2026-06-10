// admin-dashboard.js — Dashboard do Admin Homy

const AdminDashboard = window.AdminDashboard = {

  async load(semanaId) {
    try {
      await SP.init();

      const r = await SP.getDashboardResumo(semanaId);

      this._renderCards(r);
      await this._renderAusencias(r, semanaId);
      await this._renderOperacaoDia(r, semanaId);
      this._renderGerencial(r);
      this._renderSetores(r);

    } catch (e) {
      console.error("[Dashboard]", e);
      AdminUtils.toast("Erro ao carregar dashboard: " + (e.message || e), "error");
    }
  },

  _renderCards(r) {
    AdminUtils.setTxt("dashColaboradores", r.colaboradoresAtivos ?? 0);
    AdminUtils.setTxt("dashConfirmados", r.pedidosConfirmadosColaboradores ?? 0);
    AdminUtils.setTxt("dashPendentes", r.pendentesColaboradores ?? 0);
    AdminUtils.setTxt("dashCheckins", r.checkinsHoje ?? 0);
    AdminUtils.setTxt("dashExtras", r.extrasAtivos ?? 0);
    AdminUtils.setTxt("dashTotalHoje", r.totalPedidosHoje ?? 0);
    AdminUtils.setTxt("dashAusencias", r.ausenciasHoje ?? 0);
    AdminUtils.setTxt("dashTotalSemana", r.totalPedidosSemana ?? 0);
  },

  async _renderAusencias(r, semanaId) {
    AdminUtils.setTxt("dashNaoMarcaram", r.pendentesColaboradores ?? 0);
    AdminUtils.setTxt("dashCancelados", r.ausenciasHoje ?? 0);
    AdminUtils.setTxt("dashTravados", "—");

    const tbody = document.getElementById("dashAusenciasTable");
    if (!tbody) return;

    tbody.innerHTML = `<tr><td colspan="3" class="empty-cell">Carregando...</td></tr>`;

    try {
      const norm = v => AdminUtils.norm(v);
      const hojeISO = new Date().toISOString().slice(0, 10);

      const [pedidos, ausencias] = await Promise.all([
        SP.getPedidos(semanaId).catch(() => []),
        SP.getAusenciasRefeitorio().catch(() => [])
      ]);

      const ausentesPedidos = pedidos.filter(p =>
        norm(SP.pick(p, "Dia")) === norm(r.diaHoje) &&
        (
          SP.isAusenciaPedido(p) ||
          ["cancelado", "afastado", "ferias", "férias", "nao vai almocar", "não vai almoçar", "bloqueado", "travado", "ausente"]
            .includes(norm(SP.pick(p, "Status") || ""))
        )
      ).map(p => ({
        nome: SP.pick(p, "Colaborador_nome", "Title") || "—",
        centroCusto: SP.pick(p, "Centro_Custo") || "—",
        status: SP.pick(p, "Status") || "Ausente"
      }));

      const ausentesLista = ausencias.filter(a => {
        const ativo = SP.isTrue(SP.pick(a, "Ativo") ?? true);
        const ini = String(SP.pick(a, "Data_Inicio") || "").slice(0, 10);
        const fim = String(SP.pick(a, "Data_Fim") || "").slice(0, 10);

        if (!ativo || !ini || !fim) return false;

        return hojeISO >= ini && hojeISO <= fim;
      }).map(a => ({
        nome: SP.pick(a, "Colaborador_nome", "Title") || "—",
        centroCusto: SP.pick(a, "Centro_Custo") || "—",
        status: SP.pick(a, "Motivo") || "Ausência"
      }));

      const mapa = new Map();

      [...ausentesPedidos, ...ausentesLista].forEach(a => {
        const key = AdminUtils.norm(a.nome + "|" + a.status);
        if (!mapa.has(key)) mapa.set(key, a);
      });

      const ausentes = Array.from(mapa.values());

      AdminUtils.setTxt("dashCancelados", ausentes.length);

      if (!ausentes.length) {
        tbody.innerHTML = `<tr><td colspan="3" class="empty-cell">Nenhuma ausência hoje.</td></tr>`;
        return;
      }

      tbody.innerHTML = ausentes.slice(0, 14).map(p => `
        <tr>
          <td>${AdminUtils.esc(p.nome || "—")}</td>
          <td>${AdminUtils.esc(p.centroCusto || "—")}</td>
          <td>${AdminUtils.badge(p.status || "Ausente")}</td>
        </tr>
      `).join("");

    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="3" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message || e)}</td></tr>`;
    }
  },

  async _renderOperacaoDia(r, semanaId) {
    const sel = document.getElementById("dashOperacaoDia");

    if (sel && !sel.dataset.bound) {
      sel.value = r.diaHoje;
      sel.dataset.bound = "1";
      sel.addEventListener("change", () => this._carregarOperacaoTabela(semanaId, sel.value));
    }

    const btn = document.getElementById("btnAtualizarDashboard");
    if (btn && !btn.dataset.bound) {
      btn.dataset.bound = "1";
      btn.addEventListener("click", () => AdminDashboard.load(AdminState.getSemanaId()));
    }

    await this._carregarOperacaoTabela(semanaId, sel?.value || r.diaHoje);
  },

  async _carregarOperacaoTabela(semanaId, dia) {
    const tbody = document.getElementById("dashOperacaoTable");
    if (!tbody) return;

    tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Carregando...</td></tr>`;

    try {
      const pedidos = await SP.getPedidos(semanaId).catch(() => []);
      const norm = v => AdminUtils.norm(v);
      const lista = pedidos.filter(p => norm(SP.pick(p, "Dia")) === norm(dia));

      if (!lista.length) {
        tbody.innerHTML = `<tr><td colspan="6" class="empty-cell">Nenhum pedido para este dia.</td></tr>`;
        return;
      }

      tbody.innerHTML = lista.slice(0, 60).map(p => {
        const id = AdminUtils.esc(p.id || "");
        const nome = AdminUtils.esc(SP.pick(p, "Colaborador_nome", "Title") || "—");
        const opcao = AdminUtils.esc(SP.pick(p, "Opcao") || "—");
        const prato = AdminUtils.esc(SP.pick(p, "Nome_Prato") || "—");
        const status = SP.pick(p, "Status") || "Pendente";

        return `<tr>
          <td>${nome}</td>
          <td>${AdminUtils.esc(SP.pick(p, "Centro_Custo") || "—")}</td>
          <td><span class="badge badge-blue">${opcao}</span></td>
          <td>${prato}</td>
          <td>${AdminUtils.badge(status)}</td>
          <td><div class="table-actions">
            <button class="btn-icon" onclick="AdminOperacao.alterarStatus('${id}','Confirmado')">✅</button>
            <button class="btn-icon danger" onclick="AdminOperacao.alterarStatus('${id}','Cancelado')">❌</button>
            <button class="btn-icon" onclick="AdminOperacao.alterarStatus('${id}','Não vai almoçar')">🚫</button>
          </div></td>
        </tr>`;
      }).join("");

    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="6" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message || e)}</td></tr>`;
    }
  },

  _renderGerencial(r) {
    const el = document.getElementById("dashGerencialList");
    if (!el) return;

    const itens = [
      {
        titulo: "Consumo da semana",
        valor: `${r.totalPedidosSemana || 0} refeições`
      },
      {
        titulo: "Pendências da semana",
        valor: `${r.pendentesColaboradores || 0} registros`
      },
      {
        titulo: "Extras confirmados",
        valor: `${r.extrasConfirmados || 0} extras`
      },
      {
        titulo: "Ausências hoje",
        valor: `${r.ausenciasHoje || 0} registros`
      }
    ];

    el.innerHTML = itens.map(item => `
      <div class="dashboard-list-item">
        <div>
          <div class="dashboard-list-main">${AdminUtils.esc(item.titulo)}</div>
          <div class="dashboard-list-sub">${AdminUtils.esc(item.valor)}</div>
        </div>
      </div>
    `).join("");
  },

  _renderSetores(r) {
    const el = document.getElementById("dashSetoresList");
    if (!el) return;

    const setores = Array.isArray(r.setoresHoje) ? r.setoresHoje : [];

    if (!setores.length) {
      el.innerHTML = `
        <div class="dashboard-list-item">
          <div>
            <div class="dashboard-list-main">Nenhum setor com pedido hoje</div>
            <div class="dashboard-list-sub">0 refeições</div>
          </div>
        </div>
      `;
      return;
    }

    el.innerHTML = setores.map(item => {
      let nome = "Sem CC";
      let total = 0;

      if (Array.isArray(item)) {
        nome = item[0] || "Sem CC";
        total = item[1] || 0;
      } else if (item && typeof item === "object") {
        nome = item.nome || item.setor || item.centroCusto || "Sem CC";
        total = item.total || item.quantidade || 0;
      }

      return `
        <div class="dashboard-list-item">
          <div>
            <div class="dashboard-list-main">${AdminUtils.esc(nome)}</div>
            <div class="dashboard-list-sub">${AdminUtils.esc(total)} refeições</div>
          </div>
        </div>
      `;
    }).join("");
  }
};
