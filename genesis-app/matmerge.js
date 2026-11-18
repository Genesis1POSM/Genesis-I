/* Materiais — junção por linha e permissão por departamento, aplicada NO SERVIDOR.
   Cada requisição tem `upd` (ms da última edição). Quem salva manda a lista inteira, mas só as linhas
   que a pessoa pode editar (o seu departamento; todos para Almoxarifado e administrador) são aceitas;
   as demais ficam como estão no banco. Exclusões viajam em `matDeleted` (ids). */
export const BORDO_DEPTS = ["Marinha", "Convés", "Máquinas", "Elétrica", "Guindaste"];
export const MAT_DEPTS = ["Máquinas", "Guindaste", "Marinha", "SMS", "Convés", "Base Prio", "Almoxarifado"];

export function acharUsuario(users, nome, senha) {
  if (!nome || !senha) return null;
  const n = String(nome).trim().toLowerCase();
  return (users || []).find((u) => String(u.username || "").trim().toLowerCase() === n && String(u.password) === String(senha)) || null;
}
const ehAdmin = (u) => !!u && (u.admin === true || (u.admin === undefined && String(u.username || "").toLowerCase() === "lethicia"));
export function podeEditarDepto(u, depto) {
  if (!u) return false;
  if (ehAdmin(u)) return true;
  const ds = (Array.isArray(u.matDepts) ? u.matDepts : u.matDept ? [u.matDept] : []).filter((d) => MAT_DEPTS.includes(d) || d === "Elétrica");
  return ds.includes("Almoxarifado") || ds.includes(depto);
}

export function mesclarLista(atual, atualDel, recebido, recebidoDel, pode) {
  const cur = new Map((atual || []).map((r) => [r.id, r]));
  const del = new Set(atualDel || []);
  for (const id of recebidoDel || []) {
    const e = cur.get(id);
    if (e && pode(e.departamento)) { cur.delete(id); del.add(id); }
  }
  for (const r of recebido || []) {
    if (!r || !r.id || del.has(r.id)) continue;
    const e = cur.get(r.id);
    if (!e) { if (pode(r.departamento)) cur.set(r.id, r); continue; }
    if ((r.upd || 0) > (e.upd || 0) && pode(e.departamento) && pode(r.departamento)) cur.set(r.id, r);
  }
  return { rows: [...cur.values()], deleted: [...del] };
}

export const mesclarMateriais = (a, ad, r, rd, u) => mesclarLista(a, ad, r, rd, (d) => podeEditarDepto(u, d));
/* Registro de Bordo: administrador edita tudo; demais, só os departamentos atribuídos (Almoxarifado é um departamento comum aqui) */
export function podeEditarBordo(u, depto) {
  if (!u) return false;
  if (ehAdmin(u)) return true;
  const ds = (Array.isArray(u.matDepts) ? u.matDepts : u.matDept ? [u.matDept] : []).filter((d) => MAT_DEPTS.includes(d) || d === "Elétrica");
  return ds.includes(depto);
}
/* validação de registro sem Job: só o administrador pode marcar "Validado"/"Recusado"; os demais só deixam "Em validação" */
const VAL = ["valStatus", "valPor", "valEm", "valNota"];
export function mesclarBordo(a, ad, r, rd, u) {
  const adm = ehAdmin(u);
  const antes = new Map((a || []).map((x) => [x.id, x]));
  const rec = (r || []).map((x) => {
    if (adm || !x) return x;
    const e = antes.get(x.id);
    if (x.valStatus === "Validado" || x.valStatus === "Recusado") {
      if (e && e.valStatus === x.valStatus) { const y = { ...x }; for (const k of VAL) y[k] = e[k]; return y; }
      const y = { ...x, valStatus: "Em validação" }; for (const k of VAL.slice(1)) y[k] = ""; return y;
    }
    return x;
  });
  return mesclarLista(a, ad, rec, rd, (d) => podeEditarBordo(u, d));
}
