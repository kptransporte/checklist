# KP CHECKLIST

Check-list de rota da frota da **KP Transportes Ltda.** App leve para celular Android, que funciona sem internet e envia os dados sozinho quando o sinal volta.

- **Motorista:** abre o link, instala na tela inicial, informa nome e CPF uma vez e faz saída e retorno passo a passo, com fotos.
- **Escritório:** na tela de entrada, "Sou do escritório". Cada pessoa tem nome e senha próprios, cadastrados pelo Dev no app.
  - **Dev:** tudo, mais o cadastro de pessoas e o registro de acessos.
  - **Administrativo:** vê tudo, valida, reprova e apaga (o check-list vai para a aba "Lixeira").
  - **Diretoria:** vê tudo.
  - Todo login, validação, exclusão e mudança de usuário fica na aba "Registro de acessos".

## Como funciona

```
Celular (app)  --guarda no aparelho-->  fila offline  --quando tem internet-->  Google Apps Script
                                                                                   ├─ Planilha "Checklists" (uma linha por viagem)
                                                                                   └─ Drive: pasta "KP CHECKLIST - Fotos" (por mês / viagem)
```

## Instalação (uma vez, no computador)

### 1. Planilha e servidor no Google
1. Entre na conta Google da empresa e abra **sheets.new**. Dê o nome **KP CHECKLIST**.
2. Menu **Extensões → Apps Script**.
3. Apague o que estiver no editor e cole todo o conteúdo de [`servidor/Code.gs`](servidor/Code.gs).
4. Na linha `SENHA_DEV_INICIAL`, troque pela senha do Dev (mínimo 6 caracteres). As outras pessoas o Dev cadastra pelo app.
5. Clique em **Salvar**. Na barra de cima, escolha a função **configurar** e clique em **Executar**.
   Autorize: *Revisar permissões → sua conta → Avançado → Acessar KP CHECKLIST → Permitir*.
6. **Implantar → Nova implantação →** engrenagem **→ App da Web**.
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
7. Clique em **Implantar** e copie o **URL do app da Web** (termina em `/exec`).
8. Cole esse URL em `config.js`, no campo `API_URL`.

> Se mudar o `Code.gs` depois, use **Implantar → Gerenciar implantações → editar → Nova versão** para manter o mesmo URL.

### 2. Publicar o app
O app é publicado pelo GitHub Pages a partir do branch **`gh-pages`** (já ativo).
Endereço para os motoristas: **https://kptransporte.github.io/checklist/**

Toda alteração precisa ir para `main` e `gh-pages`.

## No celular do motorista (Android)
1. Abra o link no **Chrome**.
2. Toque em **"Instalar o app neste celular"** ou em **⋮ → Instalar app / Adicionar à tela inicial**.
3. Pronto: o ícone **KP Checklist** fica na tela inicial.

## Segurança e LGPD
- O CPF dos motoristas fica na planilha e só quem tem acesso à planilha e à senha do administrativo vê.
- Compartilhe a planilha e a pasta de fotos só com quem precisa.
- O código público deste repositório não contém dados de ninguém. Os dados ficam no Google da empresa.

## Arquivos
| Arquivo | O que é |
|---|---|
| `index.html` | O app inteiro (motorista + administrativo) |
| `sync.js` | Fila offline (guarda no celular e envia depois) |
| `sw.js` | Faz o app abrir sem internet e enviar em segundo plano |
| `config.js` | Endereço do servidor Google |
| `servidor/Code.gs` | Servidor (colar no Apps Script da planilha) |
