# Landing page — Unit Mail

Landing page em PHP puro (página única, sem framework, sem banco de dados)
para hospedar em `https://agenciaunit.com/unit-mail`.

## Requisitos

- PHP 7.4+ (funciona em qualquer hospedagem compartilhada)
- Apache com `mod_rewrite` **opcional**
- Extensão `openssl` (para a chamada HTTPS à API do GitHub)

## Instalação

1. Envie todo o conteúdo desta pasta para `public_html/unit-mail/` (ou a
   pasta equivalente do seu domínio) via FTP/cPanel.
2. Dê permissão de escrita na pasta `cache/` (755 costuma bastar) — ela
   guarda a resposta da API do GitHub por 1 hora.
3. Abra `https://agenciaunit.com/unit-mail`.

## Configuração (`config.php`)

| Chave | O que faz |
|---|---|
| `base_url` | Caminho do site no domínio. Se mover para a raiz, mude para `/`. |
| `paypal_url` | Link de doação do PayPal (hosted button). Atual: `https://www.paypal.com/donate/?hosted_button_id=X95ATUT9TXGNW`. Para trocar, gere outro em paypal.com/donate e cole aqui. |
| `github_repo` | Repositório usado para os botões de download. |
| `cache_ttl` | Tempo do cache da API do GitHub (segundos). |

## Downloads

Os botões **Windows** e **Linux** apontam direto para o arquivo da
**última release publicada no GitHub** (Setup `.exe` e AppImage,
respectivamente). Ao publicar uma versão nova, o site passa a servir o
link novo sozinho — **nenhuma alteração no site é necessária**.
Se a API do GitHub estiver fora, os botões caem para a página de releases
como fallback.

## Idioma

- Padrão: português; `?lang=en` exibe em inglês (botão EN/PT no menu).

## Estrutura

```
unit-mail/
├── index.php        → página única (hero animado + recursos + download + doação)
├── config.php       → base_url, paypal_url, repo
├── lang.php         → textos pt-BR / en-US
├── .htaccess        → fallback de rotas
├── inc/github.php   → última release via API + cache
├── assets/          → css, js, ícone
└── cache/           → cache da API (precisa de escrita)
```
