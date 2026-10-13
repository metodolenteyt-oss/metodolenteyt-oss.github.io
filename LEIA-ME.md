# Método Lente — captura de leads (custo zero)

Fluxo: **QR no fim do vídeo / link na descrição → página da aula → nome, e-mail, WhatsApp → PDF na hora + e-mail → lead na planilha.**

| Peça | Onde | Custo |
|---|---|---|
| Página de captura (este repositório) | GitHub Pages | grátis |
| Planilha de leads + e-mail automático | Google Sheets + Apps Script | grátis (100 e-mails/dia) |
| QR code + tela final 4K | `tools/nova-aula.sh` (no Mac) | grátis |

## 1. Planilha + Apps Script (5 min, uma vez só)
1. Crie uma planilha em sheets.google.com chamada **Leads Método Lente**.
2. Menu **Extensões → Apps Script**. Apague o conteúdo e cole `apps-script/Code.gs`. Salve.
3. **Implantar → Nova implantação → Tipo: App da Web**.
   - Executar como: **Eu** · Quem pode acessar: **Qualquer pessoa** → Implantar → autorize com sua conta.
4. Copie a URL que termina em `/exec` e cole em `aulas.js`, no campo `APPS_SCRIPT_URL`.
5. Teste: abra a URL `/exec` no navegador → deve aparecer "Método Lente — captura ativa ✔".

> Se alterar o `Code.gs` depois: **Implantar → Gerenciar implantações → editar → Nova versão** (a URL continua a mesma).

## 2. Publicar a página (GitHub Pages)
1. Crie o repositório público **metodolenteyt-oss.github.io** na sua conta do GitHub e envie esta pasta (`git push`).
2. No repositório: **Settings → Pages → Branch: main / root → Save**.
3. Em ~1 min o site fica em `https://metodolenteyt-oss.github.io/`.
   Se o endereço for diferente, atualize `tools/site-url.txt` e `SITE_URL` no `Code.gs`, e gere o QR de novo (passo 3).
4. Preencha os campos amarelos de `privacidade.html` (nome/CPF ou CNPJ, e-mail de contato, data).

## 3. Para cada vídeo novo
```sh
tools/nova-aula.sh 02 /caminho/mapa-mental-aula-02.png /caminho/mapa-mental-aula-02.pdf
```
Gera: `mapas/aula-02-xxxx.pdf`, `previews/aula-02.jpg` (prévia borrada), `qr/aula-02.png` e
`endcard/tela-final-aula-02-4k.png` + `.mp4` (20 s). Depois:
1. Em `aulas.js`, copie o bloco `"01"` como `"02"` e troque textos e o caminho do PDF que o script mostrou.
2. `git add . && git commit -m "aula 02" && git push`.
3. No vídeo: coloque a tela final nos últimos 15–20 s. A Profe Lu fala:
   *"Aponte a câmera do celular para o QR code — ou, se estiver no celular, toque no link da descrição — e baixe grátis o mapa mental desta aula."*
4. Na descrição e no comentário fixado: `https://metodolenteyt-oss.github.io/?a=02&o=yt`

## Como ler a planilha
- Uma linha por download; `aula` diz de qual vídeo veio e `origem` diz se foi QR (`qr`) ou descrição (`yt`).
- `retorno = sim`: a pessoa já era lead e baixou outra aula (sinal de interesse alto → prioridade na venda do curso).
- Para contar leads únicos: **Dados → Limpeza de dados → Remover duplicatas** (coluna e-mail) numa cópia da aba.
- Para a venda do curso: filtre `consent_whatsapp = sim` antes de mandar mensagem no WhatsApp.

## LGPD — o essencial
- Só use os dados para o que está na política (mapa + conteúdos/ofertas do Método Lente).
- Pedido de exclusão ("SAIR" ou e-mail): apague as linhas da pessoa em até 15 dias.
- A caixa de idade existe porque muitos vestibulandos são menores. Vale uma revisão jurídica antes de anunciar o curso.

## Observações
- O PDF fica num endereço com sufixo aleatório: quem tiver o link direto consegue baixar. Para um material gratuito isso é aceitável.
- A cota do Gmail é de 100 e-mails/dia na conta gratuita. Se passar disso, o lead é salvo e baixa o PDF na hora, mas não recebe o e-mail.
- Domínio próprio (ex.: metodolente.com.br, cerca de R$ 40/ano no registro.br) deixa a URL curta e mais confiável; o GitHub Pages aceita domínio próprio.
