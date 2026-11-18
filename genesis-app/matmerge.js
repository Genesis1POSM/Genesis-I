/* Materiais — junção por linha e permissão por departamento, aplicada NO SERVIDOR.
   Cada requisição tem `upd` (ms da última edição). Quem salva manda a lista inteira, mas só as linhas
   que a pessoa pode editar (o seu departamento; todos para Almoxarifado e administrador) são aceitas;
   as demais ficam como estão no banco. Exclusões viajam em `matDeleted` (ids). */
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
  if (u.role === "viewer") return false;
  const ds = (Array.isArray(u.matDepts) ? u.matDepts : u.matDept ? [u.matDept] : []).filter((d) => MAT_DEPTS.includes(d));
  return ds.includes("Almoxarifado") || ds.includes(depto);
}

export function mesclarMateriais(atual, atualDel, recebido, recebidoDel, usuario) {
  const cur = new Map((atual || []).map((r) => [r.id, r]));
  const del = new Set(atualDel || []);
  for (const id of recebidoDel || []) {
    const e = cur.get(id);
    if (e && podeEditarDepto(usuario, e.departamento)) { cur.delete(id); del.add(id); }
  }
  for (const r of recebido || []) {
    if (!r || !r.id || del.has(r.id)) continue;
    const e = cur.get(r.id);
    if (!e) { if (podeEditarDepto(usuario, r.departamento)) cur.set(r.id, r); continue; }
    if ((r.upd || 0) > (e.upd || 0) && podeEditarDepto(usuario, e.departamento) && podeEditarDepto(usuario, r.departamento)) cur.set(r.id, r);
  }
  return { rows: [...cur.values()], deleted: [...del] };
}
