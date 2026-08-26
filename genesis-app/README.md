# GENESIS I — Maintenance & Port Call

Site com backend (Node/Express) + banco de dados (PostgreSQL), pronto para deploy no Render.
Tudo que qualquer pessoa fizer (serviços, materiais, pagamentos, usuários, Port Calls) fica salvo
no banco e aparece igual para todo mundo que acessar o link.

## Rodar localmente

```bash
npm install
npm run build      # gera a pasta dist/ com o frontend
npm start           # sobe o servidor em http://localhost:3000
```

Sem definir `DATABASE_URL`, o servidor roda em modo de teste com os dados só na memória
(reinicia zerado a cada `npm start` de novo) — bom para conferir se está tudo funcionando
antes de configurar o banco de verdade.

Para desenvolvimento com hot-reload do frontend: `npm run dev` (abre em localhost:5173 e
já reencaminha as chamadas `/api` para o servidor em localhost:3000 — rode `npm start`
em outro terminal ao mesmo tempo).

## Deploy no Render — COM banco de dados compartilhado

### 1. Criar o banco de dados
No painel do Render: **New +** → **PostgreSQL**
- Dê um nome (ex: `genesis-db`)
- Plano Free está ok para começar
- Depois de criado, copie a **Internal Database URL** (Render mostra na página do banco)

### 2. Subir o código para o GitHub
```bash
cd genesis-app
git init
git add .
git commit -m "Primeira versão do GENESIS I"
git branch -M main
git remote add origin https://github.com/SEU-USUARIO/genesis-i.git
git push -u origin main
```
(crie o repositório vazio em github.com antes do `git push`)

### 3. Criar o Web Service
No painel do Render: **New +** → **Web Service**
- Conecte o repositório que você acabou de subir
- **Build Command:** `npm install && npm run build`
- **Start Command:** `npm start`
- Em **Environment Variables**, adicione:
  - `DATABASE_URL` = a Internal Database URL que você copiou no passo 1
- Clique em **Create Web Service**

Em alguns minutos você recebe uma URL pública (tipo `https://genesis-i.onrender.com`) —
esse é o link que Lethicia, Davi e Mariana vão usar. Tudo que qualquer um deles fizer
fica salvo no banco e aparece pros outros dois também.

Toda vez que você der `git push`, o Render rebuilda e republica automaticamente.

## Como funciona a persistência

- O app guarda todos os dados (serviços, materiais, pagamentos, Port Calls, usuários) num
  único registro JSON dentro de uma tabela `app_state` no PostgreSQL.
- Ao abrir o site, ele busca esse registro (`GET /api/state`) e preenche a tela.
- A cada alteração (editar uma célula, adicionar uma linha, etc.), o app espera cerca de
  0,7 segundo de inatividade e salva tudo de volta no banco (`PUT /api/state`) — assim não
  fica mandando uma requisição a cada letra digitada.
- Se o servidor cair ou a internet da pessoa cair no meio de uma edição, aparece um aviso
  discreto avisando que a alteração ficou só local até a conexão voltar.

## Sobre o login

O login (Lethicia/Davi/Mariana) agora também é salvo no banco — se alguém mudar a própria
senha pela aba Configurações, isso vale para todo mundo que acessar depois. Ainda assim,
é uma separação de acesso simples (sem criptografia de senha, sem sessão segura de servidor),
adequada para uso interno da equipe, não para dados sigilosos críticos.
