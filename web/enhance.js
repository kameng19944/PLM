/* ============================================================
   NutraPLM v3 增强层
   - 登录 / 会话 / 角色权限（与服务端一致，服务端为准）
   - 价格管理：价格历史曲线、调价申请与多级审批
   - 配方 BOM ↔ 产品 BOM 打通、BOM 版本存档与差异对比
   - 营养成分 NR% 计算
   - Excel(.xlsx) 直接导入
   ============================================================ */
(function () {
  'use strict';

  /* ------------------------------------------------ 1. API */
  const API = {
    token: localStorage.getItem('plm_token') || '',
    user: null, perms: [], users: [],

    async req(method, url, body, isForm) {
      const opt = { method, headers: {} };
      if (this.token) opt.headers['Authorization'] = 'Bearer ' + this.token;
      if (body !== undefined) {
        if (isForm) opt.body = body;
        else { opt.headers['Content-Type'] = 'application/json'; opt.body = JSON.stringify(body); }
      }
      const r = await fetch(url, opt);
      const ct = r.headers.get('content-type') || '';
      const data = ct.indexOf('json') >= 0 ? await r.json() : {};
      if (r.status === 401) {
        localStorage.removeItem('plm_token'); API.token = '';
        showLogin('登录已过期，请重新登录');
        throw { status: 401, message: '登录已过期' };
      }
      if (!r.ok) throw { status: r.status, message: data.detail || data.message || '请求失败' };
      return data;
    },
    get(u) { return this.req('GET', u); },
    post(u, b) { return this.req('POST', u, b); },
    upload(u, form) { return this.req('POST', u, form, true); },
    commit(db) { return this.post('/api/commit', { version: DB_VER, db: db }); },
    async load() { applyState(await this.get('/api/snapshot')); },
    async reload() {
      const s = await this.get('/api/snapshot');
      const keepFormula = ui.formulaId, keepEdit = ui.editing ? ui.editing.id : null;
      applyState(s);
      if (cur === 'formulas' && keepFormula) { ui.formulaId = keepFormula; ui.editing = null; render(); }
    }
  };
  window.PLM_API = API;

  const ROLE_LABEL = { admin: '管理员', rd: '研发', qa: '质量', ra: '法规', mfg: '生产', viewer: '访客' };
  const can = p => API.perms.indexOf('*') >= 0 || API.perms.indexOf(p) >= 0;

  /* ------------------------------------------------ 2. 登录界面 */
  function showLogin(msg) {
    closeModal();
    let host = document.getElementById('loginRoot');
    if (host) host.remove();
    host = document.createElement('div');
    host.id = 'loginRoot';
    host.innerHTML =
      '<div class="login-wrap">'
      + '<div class="login-side"><div class="lg-logo">N</div>'
      +   '<h1>NutraPLM</h1>'
      +   '<p>功能性食品原料研发与产品全生命周期管理系统<br>物料 · 配方 · BOM · 变更 · 价格 · 合规，一处协同。</p>'
      +   '<ul><li>服务端数据库，多人实时协作</li><li>角色权限与操作审计</li><li>调价审批与价格历史追溯</li><li>BOM 版本对比与成本联动</li></ul>'
      +   '<div class="ver">v3.0 · FastAPI + SQLite</div></div>'
      + '<div class="login-main"><div class="login-card">'
      +   '<h2>登录工作台</h2><div class="sub">请使用企业账号登录，不同角色可见的功能略有差异</div>'
      +   '<div class="lg-err" id="lgErr"' + (msg ? ' style="display:block"' : '') + '>' + esc(msg || '') + '</div>'
      +   '<div class="lg-fld"><label>账号</label><input id="lgUser" autocomplete="username" placeholder="如 rd01"></div>'
      +   '<div class="lg-fld"><label>密码</label><input id="lgPwd" type="password" autocomplete="current-password" placeholder="请输入密码"></div>'
      +   '<button class="lg-btn" id="lgBtn">登 录</button>'
      +   '<div class="lg-users"><div class="t">演示账号（点击自动填入）</div><div class="chips">'
      +     [['admin', 'admin123', '管理员'], ['rd01', 'rd123', '研发'], ['qa01', 'qa123', '质量'],
               ['ra01', 'ra123', '法规'], ['mfg01', 'mfg123', '生产'], ['viewer', 'view123', '访客（只读）']]
             .map(a => '<span class="chip" data-u="' + a[0] + '" data-p="' + a[1] + '"><b>' + a[2] + '</b> ' + a[0] + '</span>').join('')
      +   '</div></div>'
      + '</div></div></div>';
    document.body.appendChild(host);

    const submit = async () => {
      const btn = document.getElementById('lgBtn');
      const err = document.getElementById('lgErr');
      const username = document.getElementById('lgUser').value.trim();
      const password = document.getElementById('lgPwd').value;
      if (!username || !password) { err.style.display = 'block'; err.textContent = '请输入账号与密码'; return; }
      btn.disabled = true; btn.textContent = '登录中…';
      try {
        const r = await API.post('/api/login', { username, password });
        API.token = r.token; API.user = r.user; API.perms = r.user.perms || [];
        localStorage.setItem('plm_token', r.token);
        document.getElementById('loginRoot').remove();
        await API.load();
        toast('欢迎回来，' + r.user.name);
      } catch (e) {
        err.style.display = 'block'; err.textContent = e.message || '登录失败';
      } finally {
        btn.disabled = false; btn.textContent = '登 录';
      }
    };
    document.getElementById('lgBtn').onclick = submit;
    host.querySelectorAll('.chip').forEach(c => {
      c.onclick = () => {
        document.getElementById('lgUser').value = c.dataset.u;
        document.getElementById('lgPwd').value = c.dataset.p;
      };
    });
    host.addEventListener('keydown', e => { if (e.key === 'Enter') submit(); });
    setTimeout(() => { const u = document.getElementById('lgUser'); if (u) u.focus(); }, 60);
  }

  /* ------------------------------------------------ 3. 状态装载 */
  function applyState(s) {
    DB = s.db;
    DB.priceReqs = DB.priceReqs || [];
    DB.priceHist = DB.priceHist || {};
    DB.bomVersions = DB.bomVersions || [];
    DB_VER = s.version;
    API.user = s.user; API.perms = s.perms || []; API.users = s.users || [];
    CURRENT_USER = s.user;
    ui.editing = null;
    render();
    updateMe();
  }

  function updateMe() {
    const nm = document.getElementById('meName'), rl = document.getElementById('meRole');
    const av = document.querySelector('.me .av');
    if (nm && API.user) {
      nm.firstChild ? (nm.childNodes[0].nodeValue = API.user.name) : (nm.textContent = API.user.name);
      if (rl) rl.textContent = (ROLE_LABEL[API.user.role] || API.user.role) + ' · ' + API.user.username;
    }
    if (av && API.user) av.textContent = (ROLE_LABEL[API.user.role] || '用').slice(0, 1);
    const me = document.querySelector('.me');
    if (me && !me.querySelector('[data-act="logout"]')) {
      const b = document.createElement('button');
      b.className = 'btn sm ghost'; b.dataset.act = 'logout'; b.textContent = '退出';
      b.style.marginLeft = '8px';
      me.appendChild(b);
    }
  }

  /* ------------------------------------------------ 4. 权限显隐 */
  const ACT_PERM = {
    'new-mat': 'materials.write', 'edit-mat': 'materials.write', 'del-mat': 'materials.write',
    'imp-mat': 'materials.write',
    'new-formula': 'formulas.write', 'fx-save': 'formulas.write', 'fx-archive': 'formulas.write',
    'fx-add': 'formulas.write', 'fx-del': 'formulas.write', 'fx-edit-info': 'formulas.write',
    'new-prod': 'products.write', 'bom-add': 'products.write', 'bom-del': 'products.write',
    'bom-sync': 'products.write', 'bom-snap': 'products.write',
    'new-change': 'changes.write', 'chg-approve': 'changes.write',
    'new-prj': 'projects.write', 'prj-next': 'projects.write',
    'new-doc': 'docs.write', 'new-sup': 'materials.write', 'del-sup': 'materials.write',
    'reset': 'admin', 'imp-json': 'admin',
    'pr-new': 'price.request', 'pr-ok': 'price.approve', 'pr-no': 'price.approve',
    'user-new': 'admin'
  };
  function applyPerms() {
    document.querySelectorAll('[data-act]').forEach(el => {
      const p = ACT_PERM[el.dataset.act];
      if (p && !can(p)) el.remove();
    });
  }

  /* ------------------------------------------------ 5. 价格模块 */
  function matName(id) { const m = mat(id); return m ? m.name : id; }
  function matCode(id) { const m = mat(id); return m ? m.code : ''; }
  function nextStep(r) { return (r.approvals || []).find(a => a.state === '待审批'); }
  function myStep(r) {
    const lab = { qa: '质量', ra: '法规', admin: '管理员' }[API.user.role];
    return (r.approvals || []).find(a => a.state === '待审批' && (a.role === lab || API.user.role === 'admin'));
  }
  function pctOf(r) { return r.oldPrice ? (r.newPrice - r.oldPrice) / r.oldPrice * 100 : 0; }

  VIEWS.price = function () {
    const all = (DB.priceReqs || []).slice().sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
    const tab = ui.prTab || 'all';
    const list = all.filter(r => tab === 'all' ? true : tab === 'pending' ? r.status === '审批中' : r.status !== '审批中');
    const pending = all.filter(r => r.status === '审批中');
    const approved = all.filter(r => r.status === '已批准');
    const avg = approved.length ? approved.reduce((s, r) => s + pctOf(r), 0) / approved.length : 0;
    const mine = pending.filter(r => myStep(r)).length;

    const rows = list.map(r => {
      const pct = pctOf(r);
      const steps = (r.approvals || []).map(a => {
        const cls = a.state === '已通过' ? 'ok' : a.state === '已驳回' ? 'no' : 'cur';
        return '<span class="step ' + cls + '">' + esc(a.role) + (a.state === '已通过' ? ' ✓' : a.state === '已驳回' ? ' ✕' : ' ·') + '</span>';
      }).join('');
      const ops = (r.status === '审批中' && myStep(r))
        ? '<button class="btn sm primary" data-act="pr-ok" data-id="' + r.id + '">通过</button>'
          + '<button class="btn sm danger" data-act="pr-no" data-id="' + r.id + '">驳回</button>'
        : '<button class="btn sm" data-act="pr-view" data-id="' + r.id + '">详情</button>';
      return '<tr><td class="code">' + esc(r.code) + '</td>'
        + '<td><div class="nm">' + esc(matName(r.mid)) + '</div><div class="sub2">' + esc(matCode(r.mid)) + '</div></td>'
        + '<td class="pr-arrow">¥' + num(r.oldPrice, 2) + ' → <b>¥' + num(r.newPrice, 2) + '</b></td>'
        + '<td class="' + (r.newPrice >= r.oldPrice ? 'pr-up' : 'pr-down') + '">' + (pct >= 0 ? '+' : '') + pct.toFixed(1) + '%</td>'
        + '<td class="sub2" style="max-width:250px">' + esc(String(r.reason).slice(0, 46)) + (String(r.reason).length > 46 ? '…' : '') + '</td>'
        + '<td><span class="badge ' + (r.status === '已批准' ? 't-green' : r.status === '已驳回' ? 't-red' : 't-amber') + '">' + esc(r.status) + '</span></td>'
        + '<td class="sub2">' + esc(r.createdBy) + '<br>' + esc(String(r.createdAt).slice(0, 10)) + '</td>'
        + '<td><div class="steps">' + steps + '</div></td>'
        + '<td><div class="acts">' + ops + '</div></td></tr>';
    }).join('');

    return '<div class="page"><div><h1>价格管理</h1><p>原辅料价格历史追溯 · 调价申请与多级审批（质量 → 法规），批准后单价自动生效</p></div>'
      + '<div class="row"><button class="btn primary" data-act="pr-new">+ 发起调价</button></div></div>'
      + '<div class="kpis" style="grid-template-columns:repeat(4,1fr)">'
      +   '<div class="kpi"><div class="lb">待审批调价单</div><div class="vl">' + pending.length + '<small>单</small></div><div class="fo">其中待我审批 <em>' + mine + '</em></div></div>'
      +   '<div class="kpi b"><div class="lb">已批准</div><div class="vl">' + approved.length + '<small>单</small></div><div class="fo">累计调价 ' + all.length + ' 单</div></div>'
      +   '<div class="kpi a"><div class="lb">平均调整幅度</div><div class="vl" style="color:' + (avg >= 0 ? 'var(--danger)' : 'var(--brand)') + '">' + (avg >= 0 ? '+' : '') + avg.toFixed(1) + '<small>%</small></div><div class="fo">按已批准单计算</div></div>'
      +   '<div class="kpi p"><div class="lb">价格记录</div><div class="vl">' + Object.values(DB.priceHist || {}).reduce((s, a) => s + a.length, 0) + '<small>条</small></div><div class="fo">覆盖 ' + Object.keys(DB.priceHist || {}).length + ' 项原辅料</div></div>'
      + '</div>'
      + '<div class="card"><div class="pr-tabs">'
      +   [['all', '全部', all.length], ['pending', '待审批', pending.length], ['done', '已完成', all.length - pending.length]]
            .map(t => '<div class="pr-tab ' + (tab === t[0] ? 'on' : '') + '" data-act="pr-tab" data-id="' + t[0] + '">' + t[1] + ' ' + t[2] + '</div>').join('')
      + '</div><div class="tb-wrap"><table><thead><tr><th>单号</th><th>原辅料</th><th>单价变化</th><th>幅度</th><th>调价原因</th>'
      + '<th>状态</th><th>发起</th><th>审批进度</th><th style="text-align:right">操作</th></tr></thead><tbody>' + rows + '</tbody></table></div></div>';
  };

  function openPriceNew(preMid) {
    const mats = DB.materials.filter(m => (m.price || 0) > 0);
    openModal({
      title: '发起调价申请', okText: '提交审批',
      body: '<div class="fgrid">'
        + fld('原辅料', 'mid', preMid || (mats[0] && mats[0].id), { req: true, type: 'select', full: true, options: mats.map(m => ({ v: m.id, t: m.code + ' · ' + m.name + '（现价 ¥' + m.price + '/' + m.unit + '）' })) })
        + fld('新单价（元）', 'newPrice', '', { req: true, type: 'number', step: '0.01', min: 0 })
        + fld('生效说明', 'reason', '', { req: true, type: 'textarea', full: true, ph: '如：供应商报价上调 6%，需评估对配方成本影响' })
        + '</div><div id="prHint" style="font-size:12px;color:var(--muted);line-height:1.8;margin-top:4px"></div>',
      onOk: async d => {
        const v = parseFloat(d.newPrice) || 0;
        if (v <= 0) { toast('请输入有效单价'); return false; }
        try {
          const r = await API.post('/api/price-requests', { mid: d.mid, newPrice: v, reason: d.reason });
          toast('已提交 ' + r.code + '，进入审批');
          await API.reload(); go('price');
        } catch (e) { toast(e.message || '提交失败'); return false; }
      }
    });
    const upd = () => {
      const m = mat(document.querySelector('[name=mid]').value);
      const v = parseFloat(document.querySelector('[name=newPrice]').value);
      const host = document.getElementById('prHint'); if (!host || !m) return;
      if (!v) { host.innerHTML = '当前单价 <b>¥' + m.price + '</b>/' + esc(m.unit) + '；提交后由「质量 → 法规」两级审批，全部通过后自动生效。'; return; }
      const pct = m.price ? (v - m.price) / m.price * 100 : 0;
      const affected = DB.formulas.filter(f => (f.items || []).some(i => i.ing === m.id));
      const delta = affected.reduce((s, f) => s + (f.items || []).filter(i => i.ing === m.id)
        .reduce((a, i) => a + (Number(i.qty) || 0) * (1 + (Number(i.loss) || 0) / 100) * (v - m.price), 0), 0);
      host.innerHTML = '当前单价 <b>¥' + m.price + '</b> → 新单价 <b>¥' + v + '</b>，'
        + '<span style="color:' + (pct >= 0 ? 'var(--danger)' : 'var(--brand)') + ';font-weight:650">' + (pct >= 0 ? '+' : '') + pct.toFixed(1) + '%</span>'
        + '；影响 <b>' + affected.length + '</b> 个配方，批次原料成本合计变化 <b>' + (delta >= 0 ? '+' : '') + '¥' + delta.toFixed(2) + '</b>。';
    };
    document.querySelector('[name=mid]').addEventListener('change', upd);
    document.querySelector('[name=newPrice]').addEventListener('input', upd);
    upd();
  }

  function openPriceDetail(id) {
    const r = (DB.priceReqs || []).find(x => x.id === id); if (!r) return;
    const steps = (r.approvals || []).map(a =>
      '<div class="verline"><span class="v">' + esc(a.role) + '</span>'
      + '<span class="d">' + (a.state === '待审批' ? '待审批' : esc(a.user) + ' · ' + esc(String(a.date).slice(0, 16))) + (a.cmt ? ' · ' + esc(a.cmt) : '') + '</span>'
      + '<span class="pill ' + (a.state === '已通过' ? 'on' : a.state === '已驳回' ? 'off' : '') + '">' + esc(a.state) + '</span></div>').join('');
    const his = (r.history || []).map(h => '<div class="hist"><b>' + esc(String(h.t).slice(0, 16)) + '</b> ' + esc(h.u) + ' ' + esc(h.a) + '</div>').join('');
    viewOnlyModal('调价单 ' + r.code,
      '<div class="dl"><div class="d"><span>原辅料</span><b>' + esc(matName(r.mid)) + '</b></div>'
      + '<div class="d"><span>单价</span><b>¥' + num(r.oldPrice, 2) + ' → ¥' + num(r.newPrice, 2) + '（' + (pctOf(r) >= 0 ? '+' : '') + pctOf(r).toFixed(1) + '%）</b></div>'
      + '<div class="d"><span>状态</span><b>' + esc(r.status) + '</b></div>'
      + '<div class="d"><span>发起</span><b>' + esc(r.createdBy) + ' · ' + esc(String(r.createdAt).slice(0, 16)) + '</b></div></div>'
      + '<div class="sec-t">调价说明</div><div style="font-size:13px;color:var(--ink2);line-height:1.85">' + esc(r.reason) + '</div>'
      + '<div class="sec-t">审批流</div>' + steps
      + '<div class="sec-t">操作记录</div>' + his, 'wide');
  }

  function decide(id, action) {
    const r = (DB.priceReqs || []).find(x => x.id === id); if (!r) return;
    openModal({
      title: (action === 'approve' ? '审批通过' : '驳回') + ' · ' + r.code, okText: action === 'approve' ? '确认通过' : '确认驳回',
      body: '<div class="imp-tip" style="margin-bottom:12px">原辅料 <b>' + esc(matName(r.mid)) + '</b>：¥' + num(r.oldPrice, 2) + ' → ¥' + num(r.newPrice, 2)
        + '（' + (pctOf(r) >= 0 ? '+' : '') + pctOf(r).toFixed(1) + '%）<br>' + (action === 'approve'
          ? '通过后若所有节点均已通过，单价将<b>立即生效</b>并写入价格历史。'
          : '驳回后原单价保持不变，发起人可修改后重新提交。') + '</div>'
        + fld('审批意见', 'comment', '', { type: 'textarea', full: true, ph: '选填，将记录到审批流与审计日志' }),
      onOk: async d => {
        try {
          const res = await API.post('/api/price-requests/' + id + '/decide', { action: action === 'approve' ? 'approve' : 'reject', comment: d.comment || '' });
          toast(action === 'approve' ? ('已通过，当前状态：' + res.status) : '已驳回');
          await API.reload(); go('price');
        } catch (e) { toast(e.message || '操作失败'); return false; }
      }
    });
  }

  /* 价格曲线（档案弹窗内追加） */
  function priceBlock(m) {
    const hist = (DB.priceHist || {})[m.id] || [];
    const reqs = (DB.priceReqs || []).filter(r => r.mid === m.id).slice(0, 3);
    let chart = '<div class="pxrow"><span>暂无价格记录</span></div>';
    if (hist.length >= 2) {
      const vals = hist.map(h => h.price);
      const min = Math.min.apply(null, vals), max = Math.max.apply(null, vals);
      const W = 620, H = 88, pad = 6;
      const span = (max - min) || 1;
      const pts = hist.map((h, i) => {
        const x = pad + i * (W - pad * 2) / (hist.length - 1);
        const y = H - pad - (h.price - min) / span * (H - pad * 2);
        return [x, y];
      });
      const line = pts.map(p => p[0].toFixed(1) + ',' + p[1].toFixed(1)).join(' ');
      const area = line + ' ' + (W - pad) + ',' + H + ' ' + pad + ',' + H;
      chart = '<svg class="pxchart" viewBox="0 0 ' + W + ' ' + H + '" preserveAspectRatio="none">'
        + '<polygon points="' + area + '" fill="rgba(22,184,119,.12)"/>'
        + '<polyline points="' + line + '" fill="none" stroke="#12a56d" stroke-width="2"/>'
        + pts.map(p => '<circle cx="' + p[0].toFixed(1) + '" cy="' + p[1].toFixed(1) + '" r="2.6" fill="#0b8557"/>').join('')
        + '</svg>'
        + '<div class="pxrow"><span>' + esc(hist[0].date) + '</span><span>' + esc(hist[hist.length - 1].date) + '</span></div>';
    }
    const rows = hist.slice(-3).reverse().map((h, i, arr) => {
      const prev = arr[i + 1];
      const d = prev ? (h.price - prev.price) : 0;
      return '<div class="pxrow"><span>' + esc(h.date) + ' · ' + esc(h.src || '') + '</span>'
        + '<b>¥' + num(h.price, 2) + (d ? ' <span class="' + (d > 0 ? 'pr-up' : 'pr-down') + '">' + (d > 0 ? '+' : '') + num(d, 2) + '</span>' : '') + '</b></div>';
    }).join('');
    const reqHtml = reqs.map(r => '<div class="pxrow"><span>' + esc(r.code) + ' · ' + esc(r.status) + '</span><b>¥' + num(r.newPrice, 2) + '</b></div>').join('');
    return '<div class="pxcard pfx"><div class="h"><b>价格历史</b><span class="v">当前 ¥' + num(m.price, 2) + ' / ' + esc(m.unit) + '</span></div>'
      + chart + rows
      + (reqHtml ? '<div class="sec-t" style="margin-top:10px">调价记录</div>' + reqHtml : '')
      + '<div class="row" style="margin-top:10px"><button class="btn sm primary" data-act="pr-new" data-id="' + m.id + '">发起调价</button>'
      + '<button class="btn sm" data-act="go-price">价格管理</button></div></div>';
  }

  /* ------------------------------------------------ 6. BOM 打通与版本对比 */
  function bomExtra(p) {
    const fs = DB.formulas || [];
    const linked = p.formulaId ? fs.find(f => f.id === p.formulaId) : null;
    const vers = (DB.bomVersions || []).filter(v => v.productId === p.id);
    const sel = '<select class="sel" id="bomFormula" style="max-width:320px"><option value="">— 未关联配方 —</option>'
      + fs.map(f => '<option value="' + f.id + '" ' + (p.formulaId === f.id ? 'selected' : '') + '>' + esc(f.code + ' · ' + f.name) + '</option>').join('')
      + '</select>';
    return '<div class="bomx">'
      + '<div class="hd">配方 ↔ 产品 BOM 打通</div>'
      + '<div class="row" style="gap:8px;align-items:center;flex-wrap:wrap">' + sel
      + '<button class="btn sm primary" data-act="bom-sync" data-id="' + p.id + '">从配方同步 BOM</button>'
      + '<span style="font-size:11.5px;color:var(--muted)">按配方每批产量折算到「' + esc(p.base || '每千单位') + '」</span></div>'
      + (linked ? '<div style="font-size:11.5px;color:var(--muted);margin-top:6px">已关联：' + esc(linked.code + ' ' + linked.name)
        + ' · 批量 ' + linked.batch + 'kg · 产出 ' + linked.servings + ' ' + esc(linked.servingUnit || '份') + '</div>' : '')
      + '<div class="hd" style="margin-top:14px">BOM 版本</div>'
      + '<div class="row" style="gap:8px;margin-bottom:8px"><button class="btn sm" data-act="bom-snap" data-id="' + p.id + '">存档当前 BOM</button>'
      + '<span style="font-size:11.5px;color:var(--muted)">已存档 ' + vers.length + ' 个版本，可与当前 BOM 逐项对比</span></div>'
      + (vers.length ? vers.map(v => '<div class="verline"><span class="v">' + esc(v.ver) + '</span>'
        + '<span class="d">' + esc(v.createdBy) + ' · ' + esc(String(v.createdAt).slice(0, 16)) + ' · ' + esc(v.note || '') + '</span>'
        + '<button class="btn sm" data-act="bom-diff" data-id="' + p.id + '" data-v="' + v.id + '">与当前对比</button></div>').join('')
        : '<div style="font-size:12px;color:var(--muted)">暂无版本存档</div>')
      + '</div>';
  }

  function flattenBom(bom) {
    const out = {};
    (function walk(arr, depth) {
      (arr || []).forEach(n => {
        const key = (n.mid || '') + '|' + depth;
        if (!out[key]) out[key] = { mid: n.mid, depth, qty: 0, loss: 0 };
        out[key].qty += Number(n.qty) || 0;
        out[key].loss = Math.max(out[key].loss, Number(n.loss) || 0);
        walk(n.children, depth + 1);
      });
    })(bom, 0);
    return out;
  }

  function openBomDiff(pid, vid) {
    const p = DB.products.find(x => x.id === pid); if (!p) return;
    API.get('/api/bom/' + pid + '/versions').then(res => {
      const v = (res.items || []).find(x => x.id === vid); if (!v) { toast('版本不存在'); return; }
      const a = flattenBom(v.bom), b = flattenBom(p.bom || []);
      const keys = Array.from(new Set(Object.keys(a).concat(Object.keys(b))));
      let add = 0, del = 0, chg = 0, costA = 0, costB = 0;
      const rows = keys.map(k => {
        const o = a[k], n = b[k];
        const mid = (o && o.mid) || (n && n.mid);
        const m = mat(mid) || {};
        const price = Number(m.price) || 0;
        const qa = o ? o.qty : 0, qb = n ? n.qty : 0;
        const ca = qa * price * (1 + (o ? o.loss : 0) / 100);
        const cb = qb * price * (1 + (n ? n.loss : 0) / 100);
        costA += ca; costB += cb;
        let st = '', cls = '';
        if (!o && n) { st = '新增'; cls = 'add'; add++; }
        else if (o && !n) { st = '删除'; cls = 'del'; del++; }
        else if (Math.abs(qa - qb) > 1e-6) { st = '用量调整'; cls = 'chg'; chg++; }
        else return '';
        return '<tr><td>' + esc(m.name || mid) + '<div class="sub2">' + esc(m.code || '') + ' · 层级 ' + ((n || o).depth + 1) + '</div></td>'
          + '<td class="num">' + (o ? num(qa, 4) : '—') + '</td>'
          + '<td class="num">' + (n ? num(qb, 4) : '—') + '</td>'
          + '<td class="num">' + (o && n ? (qb - qa > 0 ? '+' : '') + num(qb - qa, 4) : '—') + '</td>'
          + '<td class="num">' + num(ca, 2) + '</td><td class="num">' + num(cb, 2) + '</td>'
          + '<td class="num"><b>' + (cb - ca >= 0 ? '+' : '') + num(cb - ca, 2) + '</b></td>'
          + '<td><span class="' + cls + '">' + st + '</span></td></tr>';
      }).join('');
      viewOnlyModal('BOM 版本对比 · ' + p.code,
        '<div class="imp-tip">对比基准：<b>' + esc(v.ver) + '</b>（' + esc(String(v.createdAt).slice(0, 16)) + ' · ' + esc(v.createdBy) + '） ↔ <b>当前 BOM</b>'
        + '<br>新增 <b>' + add + '</b> 项 · 删除 <b>' + del + '</b> 项 · 用量调整 <b>' + chg + '</b> 项 · 成本变化 <b>'
        + (costB - costA >= 0 ? '+' : '') + '¥' + num(costB - costA, 2) + '</b></div>'
        + '<div class="imp-scroll" style="max-height:340px"><table class="diff-tb"><thead><tr><th>物料</th><th>旧用量</th><th>新用量</th><th>变化</th>'
        + '<th>旧小计</th><th>新小计</th><th>差额</th><th>类型</th></tr></thead><tbody>'
        + (rows || '<tr><td colspan="8" style="text-align:center;color:#9aa9a2;padding:18px">两个版本 BOM 完全一致</td></tr>')
        + '</tbody></table></div>', 'wide');
    }).catch(e => toast(e.message || '读取版本失败'));
  }

  /* ------------------------------------------------ 7. NR% */
  const NRV = { e: [8400, 'kJ', '能量'], p: [60, 'g', '蛋白质'], f: [60, 'g', '脂肪'], c: [300, 'g', '碳水化合物'], fib: [25, 'g', '膳食纤维'], na: [2000, 'mg', '钠'] };
  function injectNR() {
    if (cur !== 'formulas' || !ui.formulaId) return;
    const f = DB.formulas.find(x => x.id === ui.formulaId); if (!f) return;
    const card = Array.from(document.querySelectorAll('.card')).find(c => c.textContent.indexOf('营养折算（每 100g）') >= 0);
    if (!card || card.querySelector('.nrt')) return;
    const nut = card.querySelector('.nutri'); if (!nut) return;
    const servingG = f.servings ? (Number(f.batch) || 0) * 1000 / Number(f.servings) : 0;
    const keys = ['e', 'p', 'f', 'c', 'fib', 'na'];
    let any = false;
    const rows = keys.map(k => {
      const v = formulaNutri(f, k);
      const cfg = NRV[k];
      if (v == null) return '';
      any = true;
      const nr = cfg ? (v / cfg[0] * 100) : null;
      const perV = servingG ? v * servingG / 100 : null;
      const perNR = (perV != null && cfg) ? perV / cfg[0] * 100 : null;
      return '<tr><td>' + cfg[2] + '（' + cfg[1] + '）</td><td>' + num(v, 1) + '</td>'
        + '<td class="nr">' + (nr == null ? '—' : nr.toFixed(0) + '%') + '</td>'
        + '<td>' + (perV == null ? '—' : num(perV, 2)) + '</td>'
        + '<td class="nr">' + (perNR == null ? '—' : perNR.toFixed(0) + '%') + '</td></tr>';
    }).join('');
    if (!any) return;
    const html = '<table class="nrt"><thead><tr><th>项目</th><th>每 100g</th><th>NRV%</th>'
      + '<th>每' + esc(f.servingUnit || '份') + (servingG ? '（' + num(servingG, 2) + 'g）' : '') + '</th><th>NRV%</th></tr></thead>'
      + '<tbody>' + rows + '</tbody></table>'
      + '<div style="font-size:11px;color:var(--muted);margin-top:6px">NRV% = 含量 ÷ 营养素参考值 × 100%（参照 GB 28050：能量 8400kJ、蛋白质 60g、脂肪 ≤60g、碳水 300g、膳食纤维 25g、钠 2000mg）</div>';
    nut.insertAdjacentHTML('afterend', html);
  }

  /* ------------------------------------------------ 8. 团队卡片 */
  function injectTeam() {
    if (cur !== 'data') return;
    const host = document.getElementById('view'); if (!host || host.querySelector('.tmcard')) return;
    const users = API.users || [];
    const html = '<div class="card tmcard" style="margin-top:14px"><div class="card-h"><h3>团队成员与角色</h3>'
      + '<span class="sub">' + users.length + ' 个账号 · 权限由服务端强制校验</span></div><div class="card-b">'
      + users.map(u => '<div class="tm"><div class="av">' + esc((ROLE_LABEL[u.role] || '用').slice(0, 1)) + '</div>'
        + '<div class="n">' + esc(u.name) + '<i>' + esc(u.username) + ' · ' + esc(u.roleName || ROLE_LABEL[u.role]) + '</i></div>'
        + '<span class="pill ' + (u.active ? 'on' : 'off') + '">' + (u.active ? '在职' : '停用') + '</span></div>').join('')
      + '<div class="row" style="margin-top:8px"><button class="btn sm" data-act="user-new">+ 新建账号</button>'
      + '<span style="font-size:11.5px;color:var(--muted)">仅管理员可新建账号与调整角色</span></div>'
      + '</div></div>';
    host.insertAdjacentHTML('beforeend', html);
  }

  function openUserNew() {
    openModal({
      title: '新建账号', okText: '创建',
      body: '<div class="fgrid">'
        + fld('登录账号', 'username', '', { req: true, ph: '如 rd02' })
        + fld('姓名', 'name', '', { req: true, ph: '如 李工' })
        + fld('角色', 'role', 'rd', { req: true, type: 'select', options: Object.keys(ROLE_LABEL).map(k => ({ v: k, t: ROLE_LABEL[k] })) })
        + fld('初始密码', 'password', '', { req: true, type: 'password', ph: '至少 6 位' })
        + '</div>',
      onOk: async d => {
        try {
          await API.post('/api/users', d);
          toast('账号已创建');
          const s = await API.get('/api/snapshot'); API.users = s.users;
        } catch (e) { toast(e.message || '创建失败'); return false; }
        render();
      }
    });
  }

  /* ------------------------------------------------ 9. Excel 导入 */
  function excelImport() {
    const inp = document.createElement('input');
    inp.type = 'file'; inp.accept = '.xlsx';
    inp.onchange = async () => {
      const f = inp.files[0]; if (!f) return;
      const fd = new FormData(); fd.append('file', f);
      toast('正在解析 Excel…');
      try {
        const r = await API.upload('/api/import/xlsx', fd);
        const rows = r.rows || [];
        if (!rows.length) { toast('表格为空'); return; }
        const csv = rows.map(row => row.map(c => {
          const s = String(c == null ? '' : c);
          return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
        }).join(',')).join('\n');
        const ta = document.getElementById('impText');
        if (!ta) { toast('请先打开批量导入窗口'); return; }
        ta.value = csv;
        impPreview();
        toast('已读取「' + r.filename + '」' + (rows.length - 1) + ' 行，请核对预览');
      } catch (e) { toast(e.message || 'Excel 解析失败'); }
    };
    inp.click();
  }

  /* ------------------------------------------------ 10. 交互分发 */
  const HANDLERS = {
    'logout': async () => {
      try { await API.post('/api/logout'); } catch (e) { }
      localStorage.removeItem('plm_token'); API.token = ''; API.user = null;
      showLogin();
    },
    'pr-new': t => openPriceNew(t.dataset.id || ''),
    'pr-view': t => openPriceDetail(t.dataset.id),
    'pr-ok': t => decide(t.dataset.id, 'approve'),
    'pr-no': t => decide(t.dataset.id, 'reject'),
    'pr-tab': t => { ui.prTab = t.dataset.id; render(); },
    'go-price': () => go('price'),
    'bom-sync': t => {
      const p = DB.products.find(x => x.id === t.dataset.id); if (!p) return;
      const fid = document.getElementById('bomFormula').value;
      if (!fid) { toast('请先选择要关联的配方'); return; }
      const f = DB.formulas.find(x => x.id === fid); if (!f) return;
      if (!f.servings) { toast('该配方未设置产出份数，无法折算'); return; }
      const factor = 1000 / Number(f.servings);
      const oldChildren = {};
      (p.bom || []).forEach(n => { oldChildren[n.mid] = n.children || []; });
      p.bom = (f.items || []).map((it, i) => {
        const m = mat(it.ing) || {};
        return { id: 'n' + (i + 1) + Date.now().toString(36).slice(-3), mid: it.ing, qty: +(Number(it.qty) * factor).toFixed(4),
          unit: m.unit || 'kg', loss: Number(it.loss) || 0, children: oldChildren[it.ing] || [] };
      });
      p.formulaId = f.id;
      p.date = todayStr();
      save();
      toast('已按配方同步 ' + p.bom.length + ' 行 BOM（折算系数 ' + factor.toFixed(4) + '）');
      openBom(p);
    },
    'bom-snap': async t => {
      const note = prompt('本次 BOM 版本说明（可留空）：', '') ;
      if (note === null) return;
      try {
        const r = await API.post('/api/bom/' + t.dataset.id + '/snapshot', { note });
        toast('已存档 ' + r.ver);
        await API.reload();
        const p = DB.products.find(x => x.id === t.dataset.id);
        if (p) openBom(p);
      } catch (e) { toast(e.message || '存档失败'); }
    },
    'bom-diff': t => openBomDiff(t.dataset.id, t.dataset.v),
    'imp-xlsx': () => excelImport(),
    'user-new': () => openUserNew()
  };

  let lastT = 0, lastEl = null;
  function dispatchEnh(e) {
    const t = e.target.closest && e.target.closest('[data-act]');
    if (!t) return;
    const h = HANDLERS[t.dataset.act];
    if (!h) return;
    e.preventDefault();
    h(t);
  }
  document.addEventListener('pointerdown', e => {
    const t = e.target.closest && e.target.closest('[data-act]');
    if (!t || !HANDLERS[t.dataset.act]) return;
    lastEl = t; lastT = Date.now();
    dispatchEnh(e);
  }, true);
  document.addEventListener('click', e => {
    const t = e.target.closest && e.target.closest('[data-act]');
    if (!t || !HANDLERS[t.dataset.act]) return;
    if (t === lastEl && Date.now() - lastT < 900) return;
    dispatchEnh(e);
  }, true);

  /* ------------------------------------------------ 11. 挂载与联动 */
  const _render = render;
  render = function () {
    _render();
    try { applyPerms(); decorate(); } catch (err) { console.warn('[decorate]', err); }
  };

  function decorate() {
    injectNR();
    injectTeam();
  }

  const _openProfile = openProfile;
  openProfile = function (m) {
    _openProfile(m);
    const host = document.querySelector('#mform');
    if (host && !host.querySelector('.pfx')) host.insertAdjacentHTML('beforeend', priceBlock(m));
  };

  const _openBom = openBom;
  openBom = function (p) {
    _openBom(p);
    const host = document.querySelector('#mform');
    if (host && !host.querySelector('.bomx')) host.insertAdjacentHTML('beforeend', bomExtra(p));
  };

  const _openImportMat = openImportMat;
  openImportMat = function () {
    _openImportMat();
    const bar = document.querySelector('.imp-bar');
    if (bar && !bar.querySelector('[data-act="imp-xlsx"]')) {
      bar.insertAdjacentHTML('afterbegin', '<button type="button" class="btn sm" data-act="imp-xlsx">选择 Excel(.xlsx)</button>');
    }
  };

  /* 侧边栏：价格管理 */
  NAV.splice(4, 0, {
    k: 'price', t: '价格管理',
    ic: '<path d="M12 3v18"/><path d="M17 6.5c0-1.7-2.2-2.8-5-2.8S7 4.8 7 6.5s2.2 2.4 5 3 5 1.3 5 3-2.2 2.8-5 2.8-5-1.1-5-2.8"/>'
  });

  /* 多用户同步：版本轮询 */
  setInterval(async () => {
    if (!API.token || document.querySelector('.mask')) return;
    try {
      const r = await API.get('/api/version');
      if (r.version !== DB_VER) { await API.reload(); toast('已同步其他成员的更新'); }
    } catch (e) { }
  }, 15000);

  /* 启动 */
  (async function boot() {
    /* 支持从其他系统单点跳转：/?token=xxx（取出后立即从地址栏清除） */
    const qt = new URLSearchParams(location.search).get('token');
    if (qt) {
      API.token = qt;
      localStorage.setItem('plm_token', qt);
      history.replaceState({}, '', location.pathname + (location.hash || ''));
    }
    if (API.token) {
      try { await API.load(); return; } catch (e) { if (e.status !== 401) console.warn(e); }
    }
    showLogin();
  })();
})();
