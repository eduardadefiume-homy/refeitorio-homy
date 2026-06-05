// admin-relatorios.js — Relatórios do Admin Homy

const AdminRelatorios = window.AdminRelatorios = {

  async load(semanaId) {
    const tbody = document.getElementById("relTable");
    if (!tbody) return;
    tbody.innerHTML = `<tr><td colspan="7" class="empty-cell">Carregando...</td></tr>`;

    try {
      await SP.init();
      const pedidos = await SP.getPedidos(semanaId);
      const norm    = v => AdminUtils.norm(v);

      const isConf = p => {
        const s = norm(SP.pick(p, "Status") || "");
        return s === "confirmado" || s === "extra" || SP.isTrue(SP.pick(p, "Confirmado"));
      };

      const countOp = (lista, op) => lista.filter(p => norm(SP.pick(p, "Opcao")) === op).length;
      const conf    = pedidos.filter(isConf);

      // Cards de totais
      AdminUtils.setTxt("rel-principal", countOp(conf, "principal"));
      AdminUtils.setTxt("rel-light",     countOp(conf, "light"));
      AdminUtils.setTxt("rel-carne",     countOp(conf, "carne"));
      AdminUtils.setTxt("rel-massa",     countOp(conf, "massa"));
      AdminUtils.setTxt("rel-lanche",    countOp(conf, "lanche"));

      // Tabela por dia
      tbody.innerHTML = AdminUtils.DIAS.map(dia => {
        const lista = conf.filter(p => norm(SP.pick(p, "Dia")) === dia);
        return `<tr>
          <td>${AdminUtils.DIA_LABEL[dia]}</td>
          <td>${countOp(lista, "principal")}</td>
          <td>${countOp(lista, "light")}</td>
          <td>${countOp(lista, "carne")}</td>
          <td>${countOp(lista, "massa")}</td>
          <td>${countOp(lista, "lanche")}</td>
          <td><strong>${lista.length}</strong></td>
        </tr>`;
      }).join("") + `<tr style="border-top:2px solid rgba(255,255,255,.15)">
        <td><strong>Total</strong></td>
        <td>${countOp(conf, "principal")}</td>
        <td>${countOp(conf, "light")}</td>
        <td>${countOp(conf, "carne")}</td>
        <td>${countOp(conf, "massa")}</td>
        <td>${countOp(conf, "lanche")}</td>
        <td><strong>${conf.length}</strong></td>
      </tr>`;

    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty-cell" style="color:#ff8080">Erro: ${AdminUtils.esc(e.message)}</td></tr>`;
    }
  }
};
