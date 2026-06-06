// admin-relatorios.js — Relatórios do Admin Homy
// Relatórios por dia, por centro de custo e por colaborador (total mensal)

const AdminRelatorios = window.AdminRelatorios = {

  async load(semanaId) {
    try {
      await SP.init();
      const pedidos = await SP.getPedidos(semanaId);
      this._renderCards(pedidos);
      this._renderPorDia(pedidos);
      this._renderPorCentroCusto(pedidos);
      this._renderPorColaborador(pedidos);
      this._bindExportar(pedidos, semanaId);
    } catch (e) {
      ["relTableDia", "relTableCC", "relTableColab"].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.innerHTML = `<tr><td colspan="7" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</td></tr>`;
      });
    }
  },

  _isConf(p) {
    const s = AdminUtils.norm(SP.pick(p, "Status") || "");
    return s === "confirmado" || s === "extra" || SP.isTrue(SP.pick(p, "Confirmado"));
  },

  _countOp(lista, op) {
    return lista.filter(p => AdminUtils.norm(SP.pick(p, "Opcao")) === op).length;
  },

  _renderCards(pedidos) {
    const conf = pedidos.filter(p => this._isConf(p));
    AdminUtils.setTxt("rel-principal", this._countOp(conf, "principal"));
    AdminUtils.setTxt("rel-light",     this._countOp(conf, "light"));
    AdminUtils.setTxt("rel-carne",     this._countOp(conf, "carne"));
    AdminUtils.setTxt("rel-massa",     this._countOp(conf, "massa"));
    AdminUtils.setTxt("rel-lanche",    this._countOp(conf, "lanche"));
  },

  _renderPorDia(pedidos) {
    const tbody = document.getElementById("relTableDia");
    if (!tbody) return;
    const conf = pedidos.filter(p => this._isConf(p));

    const linhas = AdminUtils.DIAS.map(dia => {
      const lista = conf.filter(p => AdminUtils.norm(SP.pick(p, "Dia")) === dia);
      return `<tr>
        <td>${AdminUtils.DIA_LABEL[dia]}</td>
        <td>${this._countOp(lista, "principal")}</td>
        <td>${this._countOp(lista, "light")}</td>
        <td>${this._countOp(lista, "carne")}</td>
        <td>${this._countOp(lista, "massa")}</td>
        <td>${this._countOp(lista, "lanche")}</td>
        <td><strong>${lista.length}</strong></td>
      </tr>`;
    }).join("");

    const total = conf.length;
    tbody.innerHTML = linhas + `<tr style="border-top:2px solid rgba(255,255,255,.15)">
      <td><strong>Total</strong></td>
      <td>${this._countOp(conf, "principal")}</td>
      <td>${this._countOp(conf, "light")}</td>
      <td>${this._countOp(conf, "carne")}</td>
      <td>${this._countOp(conf, "massa")}</td>
      <td>${this._countOp(conf, "lanche")}</td>
      <td><strong>${total}</strong></td>
    </tr>`;
  },

  _renderPorCentroCusto(pedidos) {
    const tbody = document.getElementById("relTableCC");
    if (!tbody) return;
    const conf = pedidos.filter(p => this._isConf(p));

    // Agrupa por Centro_Custo
    const mapa = {};
    conf.forEach(p => {
      const cc = SP.pick(p, "Centro_Custo") || "Sem CC";
      if (!mapa[cc]) mapa[cc] = [];
      mapa[cc].push(p);
    });

    const sorted = Object.entries(mapa).sort((a, b) => b[1].length - a[1].length);

    if (!sorted.length) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Nenhum pedido com centro de custo registrado.</td></tr>`;
      return;
    }

    tbody.innerHTML = sorted.map(([cc, lista]) => `<tr>
      <td>${AdminUtils.esc(cc)}</td>
      <td>${this._countOp(lista, "principal")}</td>
      <td>${this._countOp(lista, "light")}</td>
      <td>${this._countOp(lista, "carne")}</td>
      <td>${this._countOp(lista, "massa")}</td>
      <td>${this._countOp(lista, "lanche")}</td>
      <td><strong>${lista.length}</strong></td>
    </tr>`).join("");
  },

  _renderPorColaborador(pedidos) {
    const tbody = document.getElementById("relTableColab");
    if (!tbody) return;
    const conf = pedidos.filter(p => this._isConf(p));

    // Agrupa por colaborador
    const mapa = {};
    conf.forEach(p => {
      const nome = SP.pick(p, "Colaborador_nome", "Title") || "Desconhecido";
      const cc   = SP.pick(p, "Centro_Custo") || "—";
      const key  = nome;
      if (!mapa[key]) mapa[key] = { nome, cc, lista: [] };
      mapa[key].lista.push(p);
    });

    const sorted = Object.values(mapa).sort((a, b) => b.lista.length - a.lista.length);

    if (!sorted.length) {
      tbody.innerHTML = `<tr><td colspan="4" class="empty-cell">Nenhum pedido confirmado nesta semana.</td></tr>`;
      return;
    }

    tbody.innerHTML = sorted.map(({ nome, cc, lista }) => `<tr>
      <td>${AdminUtils.esc(nome)}</td>
      <td>${AdminUtils.esc(cc)}</td>
      <td>${lista.length}</td>
      <td style="font-size:.78rem;color:rgba(143,170,210,.6)">
        ${AdminUtils.DIAS.filter(d => lista.some(p => AdminUtils.norm(SP.pick(p, "Dia")) === d))
          .map(d => AdminUtils.DIA_LABEL[d]).join(", ")}
      </td>
    </tr>`).join("");
  },

  _bindExportar(pedidos, semanaId) {
    const btn = document.getElementById("btnExportarRelatorio");
    if (btn && !btn.dataset.bound) {
      btn.dataset.bound = "1";
      btn.addEventListener("click", () => {
        if (typeof XLSX === "undefined") { AdminUtils.toast("Biblioteca XLSX não carregou.", "error"); return; }
        const conf = pedidos.filter(p => this._isConf(p));
        const linhas = conf.map(p => ({
          Colaborador:   SP.pick(p, "Colaborador_nome") || "",
          Centro_Custo:  SP.pick(p, "Centro_Custo")     || "",
          Dia:           SP.pick(p, "Dia")               || "",
          Opcao:         SP.pick(p, "Opcao")             || "",
          Prato:         SP.pick(p, "Nome_Prato")        || "",
          Status:        SP.pick(p, "Status")            || ""
        }));
        const ws = XLSX.utils.json_to_sheet(linhas);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Relatorio");
        XLSX.writeFile(wb, `relatorio-${semanaId}.xlsx`);
        AdminUtils.toast("Relatório exportado.", "success");
      });
    }
  }
};
