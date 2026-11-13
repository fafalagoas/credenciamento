# Central de Credenciamento de Imprensa · FAF

Sistema de credenciamento de imprensa da Federação Alagoana de Futebol.

**Formulário público:** https://fafalagoas.github.io/credenciamento/

## Como funciona

| Parte | Onde roda | O que faz |
|---|---|---|
| Site público (`index.html`) | GitHub Pages | Lista competições e jogos e recebe a inscrição |
| Backend (`apps-script/`) | Google Apps Script (conta contato@futeboldealagoas.net) | Valida, grava na planilha, envia e-mails, painel (`?p=admin`) e check-in (`?p=checkin`) |
| Base de dados | Planilha "Central de Credenciamento FAF" | Respostas, Partidas, Competicoes, Checkins, Historico |
| Jogos e escudos | Supabase do Urano, só leitura (`supabase/`) | Fonte oficial dos jogos, sincronizada todo dia |

Fluxo: inscrição no site → linha PENDENTE na aba **Respostas** → aprovação pelo painel ou pela coluna STATUS → e-mail automático → check-in no estádio.

## Configuração por competição

Aba **Competicoes**: `ID_URANO`, `NOME_EXIBICAO`, `MODELO` (`PROFISSIONAL` ou `BASE`), `AJUSTES` (ex.: `+selfie, -telefone`), `ATIVA`.
Jogos novos chegam do Urano fechados (`INSCRICOES_ABERTAS = NÃO`). Para abrir, mude para `SIM` na aba **Partidas**.

## Segurança e LGPD

- Só duas ações são públicas na API: `partidasParaFormulario` e `registrarInscricao`. Painel e check-in exigem senha (`PIN_ADMIN`, `PIN_CHECKIN`, nas Propriedades do script). Nenhuma senha está neste repositório.
- E-mails só saem com a propriedade `MODO_EMAIL = PRODUCAO`. Envio com falha fica como `INCERTO — CONFERIR` e não é repetido automaticamente.
- A chave do Supabase no código é a **publishable** (pública) e só executa as funções de leitura de `supabase/credenciamento_leitura.sql`.
- Fotos ficam em pasta privada do Drive, chegam sem metadados (ex.: GPS) e são apagadas 30 dias após o último jogo da pessoa.
- A rotina diária roda pelo gatilho `rotinaDiaria` (3h–4h), que recusa chamadas que não venham do próprio gatilho.
- A aba **Historico** registra quem alterou o quê, sem gravar o conteúdo do CPF.

## Publicar mudanças

1. **Backend:** alterar no Apps Script e criar uma **nova versão** em Implantar → Gerenciar implantações. O link `/exec` não muda. Manter este repositório igual ao Apps Script.
2. **Site:** alterar `index.html` aqui. O GitHub Pages atualiza sozinho.
   O `index.html` é o `apps-script/Formulario.html` com uma "ponte" que troca `google.script.run` por chamadas à API. Ao mudar o formulário, regere o site a partir dele.
