/*
 * Sagamore plan logic -- JavaScript port of src/Sagamore.Common.psm1 (Get-SagamoreFolderPlan,
 * Get-SagamoreChannelPlan, Test-SagamoreConfig, ConvertTo-SagamoreFolderName). tests/Plan.Parity.Tests.ps1 runs both
 * implementations against the same configs and fails if they ever disagree.
 */
(function (root) {
  'use strict';

  const VALID_ROLES = ['Full Control', 'Edit', 'Contribute', 'Add Only', 'Read'];
  const ROLE_RANK = { 'Full Control': 5, 'Edit': 4, 'Contribute': 3, 'Add Only': 2, 'Read': 1 };
  const GROUP_KEYS = ['Owner', 'Member', 'Support'];
  const SHARING = ['Disabled', 'ExistingExternalUserSharingOnly', 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing'];
  const INVALID_PATH = /["*:<>?\\|#%]/;
  const INVALID_PATH_G = /["*:<>?\\|#%]/g;
  const UPN = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
  const INVALID_CHANNEL_G = /[~#%&*{}+/\\:<>?|'"]/g;

  const isObj = (v) => v !== null && typeof v === 'object' && !Array.isArray(v);
  const get = (o, k, d) => (isObj(o) && o[k] !== undefined && o[k] !== null ? o[k] : d);
  const arr = (v) => (Array.isArray(v) ? v : v === undefined || v === null ? [] : [v]);

  function normPath(p) {
    return String(p || '').replace(/\\/g, '/').split('/').map((s) => s.trim()).filter((s) => s !== '').join('/');
  }

  function folderName(name) {
    let n = String(name).replace(INVALID_PATH_G, '').replace(/\s+/g, '-').replace(/-{2,}/g, '-');
    n = n.replace(/^[-. ]+|[-. ]+$/g, '');
    if (!n) throw new Error(`Name '${name}' has no usable characters for a SharePoint folder.`);
    if (n.length > 100) n = n.substring(0, 100).replace(/[-.]+$/, '');
    return n;
  }

  function channelName(name) {
    let n = String(name).replace(INVALID_CHANNEL_G, '').trim().replace(/\.+$/, '');
    if (n.length > 50) n = n.substring(0, 50).trim();
    if (!n) throw new Error(`Name '${name}' has no usable characters for a Teams channel.`);
    return n;
  }

  /* Client channels are always private (client silo); members = people invited by name on the client folder. */
  function channelPlan(cfg, extraClients) {
    extraClients = extraClients || [];
    const teams = get(cfg, 'teams', {});
    if (get(teams, 'enabled', true) === false) return [];
    const owners = arr(cfg.people).filter((p) => p.role === 'Owner').map((p) => p.upn.toLowerCase());
    const out = [];
    arr(get(teams, 'extraChannels', [])).forEach((x) => out.push({ Channel: channelName(x), Folder: null, Type: 'Standard', Owners: [], Members: [] }));
    if (get(teams, 'channelPerClient', true) !== false) {
      const byName = {};
      arr(cfg.clients).forEach((c) => { if (c.name) byName[c.name.toLowerCase()] = c; });
      const all = [];
      arr(cfg.clients).map((c) => c.name).concat(extraClients).forEach((n) => { if (n && !all.includes(n)) all.push(n); });
      const tplAccess = get(cfg.clientTemplate, 'access', null);
      all.forEach((name) => {
        const client = byName[name.toLowerCase()];
        const access = client ? get(client, 'access', tplAccess) : tplAccess;
        const invited = isObj(access)
          ? [...new Set(Object.keys(access).filter((k) => k.includes('@')).map((k) => k.toLowerCase()).filter((k) => !owners.includes(k)))].sort()
          : [];
        out.push({ Channel: channelName(name), Folder: `Clients/${name}`, Type: 'Private', Owners: owners, Members: invited });
      });
    }
    const seen = new Set();
    return out.filter((c) => { const k = c.Channel.toLowerCase(); if (seen.has(k)) return false; seen.add(k); return true; });
  }

  function resolveAccess(access) {
    if (access === null || access === undefined) return null;
    const out = { Owner: 'Full Control' };
    Object.keys(access).forEach((k) => { if (k !== 'Owner') out[k] = access[k]; });
    return out;
  }

  function folderPlan(cfg, extraClients) {
    extraClients = extraClients || [];
    const plan = new Map(); // lower path -> {Path, Access, Source}
    const add = (path, access, source) => {
      path = normPath(path);
      const parts = path.split('/');
      for (let i = 1; i < parts.length; i++) {
        const anc = parts.slice(0, i).join('/');
        if (!plan.has(anc.toLowerCase())) plan.set(anc.toLowerCase(), { Path: anc, Access: null, Source: `implied by ${path}` });
      }
      const key = path.toLowerCase();
      if (plan.has(key)) {
        if (access !== null && access !== undefined) { const e = plan.get(key); e.Access = resolveAccess(access); e.Source = source; }
      } else {
        plan.set(key, { Path: path, Access: resolveAccess(access), Source: source });
      }
    };

    arr(cfg.folders).forEach((f) => add(f.path, get(f, 'access', null), 'folders'));

    const tpl = cfg.clientTemplate || {};
    const tplAccess = get(tpl, 'access', null);
    const tplSubs = arr(get(tpl, 'subfolders', []));
    const lockedNames = arr(get(tpl, 'lockedFolderNames', [])).map((s) => s.toLowerCase());
    const lockedAccess = {};
    tplSubs.forEach((s) => {
      const leaf = normPath(s.path).split('/').pop().toLowerCase();
      if (lockedNames.includes(leaf) && get(s, 'access', null)) lockedAccess[leaf] = s.access;
    });

    const clients = arr(cfg.clients).slice();
    const known = clients.map((c) => c.name.toLowerCase());
    extraClients.forEach((x) => { if (x && !known.includes(x.toLowerCase())) { clients.push({ name: x }); known.push(x.toLowerCase()); } });

    clients.forEach((c) => {
      const base = `Clients/${c.name}`;
      add(base, get(c, 'access', tplAccess), `client ${c.name}`);
      tplSubs.forEach((s) => add(`${base}/${s.path}`, get(s, 'access', null), 'clientTemplate'));
      arr(get(c, 'folders', [])).forEach((cf) => add(`${base}/${cf.path}`, get(cf, 'access', null), `client ${c.name}`));
    });

    for (const item of plan.values()) {
      const leaf = item.Path.split('/').pop().toLowerCase();
      if (/^Clients\/[^/]+\/.+/i.test(item.Path) && lockedAccess[leaf]) {
        item.Access = resolveAccess(lockedAccess[leaf]);
        item.Source = `locked (${leaf})`;
      }
    }

    return [...plan.values()]
      .map((item, order) => ({ item, depth: item.Path.split('/').length, order }))
      .sort((a, b) => a.depth - b.depth || a.order - b.order)
      .map((x) => x.item);
  }

  function validate(cfg, opts) {
    opts = opts || {};
    const errs = [];
    for (const s of ['tenant', 'site', 'groups', 'people', 'folders', 'clientTemplate', 'clients', 'libraryAccess']) {
      if (!(s in cfg)) errs.push(`Missing required section '${s}'.`);
    }
    if (errs.length) return errs;

    if (!/^[a-z0-9][a-z0-9-]{0,62}$/.test(get(cfg.tenant, 'name', ''))) errs.push("tenant.name must be the tenant prefix only (e.g. 'sagamorealliance' from sagamorealliance.sharepoint.com).");
    const cid = String(get(cfg.tenant, 'clientId', ''));
    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(cid)) errs.push('tenant.clientId must be a GUID. Run scripts/Register-SagamoreApp.ps1 to create one.');
    else if (/^0{8}-0{4}-0{4}-0{4}-0{12}$/.test(cid) && !opts.allowPlaceholderClientId) errs.push('tenant.clientId is still the placeholder. Run scripts/Register-SagamoreApp.ps1 first.');

    if (!/^[A-Za-z0-9][A-Za-z0-9-]{0,62}$/.test(get(cfg.site, 'alias', ''))) errs.push('site.alias may contain only letters, digits and hyphens.');
    if (!get(cfg.site, 'title', '')) errs.push('site.title is required.');
    if (!get(cfg.site, 'library', '')) errs.push('site.library is required (normally "Shared Documents").');
    arr(get(cfg.site, 'adminUpns', [])).forEach((a) => { if (!UPN.test(String(a))) errs.push(`site.adminUpns: '${a}' is not a valid UPN.`); });
    if (!SHARING.includes(get(cfg.site, 'externalSharing', 'ExistingExternalUserSharingOnly'))) errs.push(`site.externalSharing must be one of: ${SHARING.join(', ')}.`);

    GROUP_KEYS.forEach((k) => { if (!get(cfg.groups, k, '')) errs.push(`groups.${k} must name an Entra ID security group.`); });

    const upns = {};
    let owners = 0;
    arr(cfg.people).forEach((p) => {
      const upn = String(get(p, 'upn', ''));
      if (!UPN.test(upn)) { errs.push(`people: '${upn}' is not a valid UPN.`); return; }
      if (upns[upn.toLowerCase()]) errs.push(`people: '${upn}' is listed twice.`);
      upns[upn.toLowerCase()] = true;
      const role = get(p, 'role', '');
      if (!GROUP_KEYS.includes(role)) errs.push(`people: '${upn}' has role '${role}'; use Owner, Member or Support.`);
      if (role === 'Owner') owners++;
    });
    if (owners < 1) errs.push('people: at least one person must have role Owner.');

    const checkAccess = (access, where) => {
      if (access === null || access === undefined) return;
      if (!isObj(access)) { errs.push(`${where}.access must be an object of principal -> role.`); return; }
      Object.keys(access).forEach((principal) => {
        if (!GROUP_KEYS.includes(principal)) {
          if (!UPN.test(principal)) errs.push(`${where}.access: '${principal}' is not Owner/Member/Support or a UPN.`);
          else if (!upns[principal.toLowerCase()]) errs.push(`${where}.access: '${principal}' is not listed in people.`);
        }
        if (!VALID_ROLES.includes(access[principal])) errs.push(`${where}.access: role '${access[principal]}' for '${principal}' is not one of: ${VALID_ROLES.join(', ')}.`);
      });
    };

    const checkSilo = (access, where) => {
      if (!isObj(access)) return;
      Object.keys(access).forEach((principal) => {
        if (principal === 'Member' || principal === 'Support') errs.push(`${where}.access: '${principal}' would give every ${principal} access and breaks the client silo. Invite people by name instead.`);
      });
    };

    checkAccess(cfg.libraryAccess, 'libraryAccess');
    if (!get(cfg.libraryAccess, 'Owner', null)) errs.push('libraryAccess must grant Owner access.');

    const seen = {};
    arr(cfg.folders).forEach((f) => {
      const p = normPath(get(f, 'path', ''));
      if (!p) { errs.push('folders: an entry has an empty path.'); return; }
      if (INVALID_PATH.test(p)) errs.push(`folders: '${p}' contains a character SharePoint does not allow.`);
      if (seen[p.toLowerCase()]) errs.push(`folders: '${p}' is listed twice.`);
      seen[p.toLowerCase()] = true;
      checkAccess(get(f, 'access', null), `folders['${p}']`);
    });

    checkAccess(get(cfg.clientTemplate, 'access', null), 'clientTemplate');
    checkSilo(get(cfg.clientTemplate, 'access', null), 'clientTemplate');
    arr(get(cfg.clientTemplate, 'subfolders', [])).forEach((s) => {
      checkAccess(get(s, 'access', null), `clientTemplate.subfolders['${s.path}']`);
      checkSilo(get(s, 'access', null), `clientTemplate.subfolders['${s.path}']`);
    });

    const names = {};
    arr(cfg.clients).forEach((c) => {
      const name = String(get(c, 'name', ''));
      if (!name) { errs.push('clients: an entry has no name.'); return; }
      if (INVALID_PATH.test(name) || name.includes('/')) errs.push(`clients: '${name}' contains a character SharePoint does not allow.`);
      if (names[name.toLowerCase()]) errs.push(`clients: '${name}' is listed twice.`);
      names[name.toLowerCase()] = true;
      checkAccess(get(c, 'access', null), `clients['${name}']`);
      checkSilo(get(c, 'access', null), `clients['${name}']`);
      arr(get(c, 'folders', [])).forEach((cf) => {
        const cp = normPath(get(cf, 'path', ''));
        if (!cp) errs.push(`clients['${name}'].folders: an entry has an empty path.`);
        else if (INVALID_PATH.test(cp)) errs.push(`clients['${name}'].folders: '${cp}' contains a character SharePoint does not allow.`);
        checkAccess(get(cf, 'access', null), `clients['${name}'].folders['${cp}']`);
        checkSilo(get(cf, 'access', null), `clients['${name}'].folders['${cp}']`);
      });
    });
    return errs;
  }

  /* Effective access of one person on one planned folder (union of their UPN entry and their role group). */
  function effectiveAccess(cfg, plan, path, person) {
    const admins = arr(get(cfg.site, 'adminUpns', [])).map((s) => s.toLowerCase());
    if (person.role === 'Owner' || admins.includes(person.upn.toLowerCase())) return { role: 'Full Control', from: 'owner' };
    const byPath = new Map(plan.map((p) => [p.Path.toLowerCase(), p]));
    let cur = normPath(path);
    let from = null;
    let access = null;
    while (cur) {
      const e = byPath.get(cur.toLowerCase());
      if (e && e.Access) { access = e.Access; from = e.Path; break; }
      const i = cur.lastIndexOf('/');
      cur = i < 0 ? '' : cur.substring(0, i);
    }
    if (!access) { access = cfg.libraryAccess || {}; from = '(library)'; }
    const keys = Object.keys(access);
    const roles = keys.filter((k) => k.toLowerCase() === person.upn.toLowerCase() || k === person.role).map((k) => access[k]);
    if (!roles.length) return { role: null, from };
    roles.sort((a, b) => (ROLE_RANK[b] || 0) - (ROLE_RANK[a] || 0));
    return { role: roles[0], from };
  }

  const api = { VALID_ROLES, ROLE_RANK, GROUP_KEYS, normPath, folderName, channelName, channelPlan, resolveAccess, folderPlan, validate, effectiveAccess };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.SagamorePlan = api;
})(typeof self !== 'undefined' ? self : this);
