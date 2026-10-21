<?php
/**
 * Unit Mail — landing page (página única).
 * ?lang=en alterna o idioma. Downloads apontam direto para o GitHub:
 * a API devolve o instalador da ÚLTIMA release, então o site nunca
 * precisa ser editado quando sai uma versão nova.
 */
if (!function_exists('str_ends_with')) {
    function str_ends_with(string $haystack, string $needle): bool
    {
        return $needle === '' || substr($haystack, -strlen($needle)) === $needle;
    }
}

$cfg = require __DIR__ . '/config.php';
require __DIR__ . '/inc/github.php';
require __DIR__ . '/lang.php';

// base_url 'auto' → detecta pela URL real (funciona em /unit-mail,
// na raiz do domínio e em qualquer subpasta de teste local).
$base = $cfg['base_url'] === 'auto'
    ? rtrim(str_replace('\\', '/', dirname($_SERVER['SCRIPT_NAME'])), '/') . '/'
    : rtrim($cfg['base_url'], '/') . '/';
$assets = $base . 'assets/';

// Cache-busting: anexa a data de modificação do arquivo (?v=…) — o
// navegador sempre baixa CSS/JS novos quando eles mudam no servidor.
$cssV = (int) @filemtime(__DIR__ . '/assets/css/style.css');
$jsV  = (int) @filemtime(__DIR__ . '/assets/js/app.js');
$langUrl = $base . '?lang=' . ($lang === 'en' ? 'pt' : 'en');

// Botões levam à página da release mais recente no GitHub — nunca
// precisa editar o site quando sai versão nova.
$release = ghLatestRelease($cfg) ?? [];
$ver     = $release['version'] ?? '';
$relUrl  = $release['url'] ?? ($cfg['releases_url'] . '/latest');
$dlWin   = $relUrl;
$dlLin   = $relUrl;
?>
<!doctype html>
<html lang="<?= $lang === 'en' ? 'en' : 'pt-BR' ?>">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><?= htmlspecialchars($cfg['app_name']) ?> — <?= htmlspecialchars(t('hero_title')) ?></title>
<meta name="description" content="<?= htmlspecialchars(t('hero_sub')) ?>">
<meta property="og:title" content="Unit Mail">
<meta property="og:description" content="<?= htmlspecialchars(t('hero_sub')) ?>">
<meta property="og:type" content="website">
<meta property="og:image" content="<?= htmlspecialchars($assets) ?>img/icon.png">
<link rel="icon" type="image/png" href="<?= htmlspecialchars($assets) ?>img/icon.png">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<link rel="stylesheet" href="<?= htmlspecialchars($assets) ?>css/style.css?v=<?= $cssV ?>">
</head>
<body>

<header class="nav">
  <div class="wrap nav-in">
    <a class="brand" href="<?= htmlspecialchars($base) ?>">
      <img src="<?= htmlspecialchars($assets) ?>img/icon.png" alt="Unit Mail" width="30" height="30">
      <span>Unit Mail</span><em class="beta">Beta</em>
    </a>
    <nav class="menu">
      <a href="#recursos"><?= t('nav_features') ?></a>
      <a href="#download"><?= t('nav_download') ?></a>
      <a class="lang" href="<?= htmlspecialchars($langUrl) ?>"><?= $lang === 'en' ? 'PT' : 'EN' ?></a>
    </nav>
    <a class="btn btn-donate" href="<?= htmlspecialchars($cfg['paypal_url']) ?>" target="_blank" rel="noopener">♥ <?= t('nav_donate') ?></a>
    <button class="burger" id="burger" aria-label="menu">☰</button>
  </div>
</header>

<!-- Hero -->
<section class="hero">
  <div class="wrap hero-in">
    <div class="hero-txt">
      <span class="badge"><?= t('hero_badge') ?></span>
      <h1><?= t('hero_title') ?></h1>
      <p><?= t('hero_sub') ?></p>
      <div class="hero-cta">
        <a class="btn btn-primary" href="#download">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M12 3v12m0 0 4.5-4.5M12 15l-4.5-4.5M4 21h16"/></svg>
          <?= t('hero_cta') ?>
        </a>
        <a class="btn btn-ghost" href="https://github.com/<?= htmlspecialchars($cfg['github_repo']) ?>" target="_blank" rel="noopener">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M12 .5C5.65.5.5 5.65.5 12c0 5.08 3.29 9.39 7.86 10.91.58.11.79-.25.79-.55v-2.15c-3.2.7-3.87-1.36-3.87-1.36-.52-1.33-1.28-1.68-1.28-1.68-1.04-.71.08-.7.08-.7 1.15.08 1.76 1.19 1.76 1.19 1.03 1.76 2.7 1.25 3.36.96.1-.75.4-1.25.72-1.54-2.55-.29-5.24-1.28-5.24-5.68 0-1.26.45-2.28 1.19-3.09-.12-.29-.52-1.46.11-3.05 0 0 .97-.31 3.17-1.18a11 11 0 0 1 5.78 0c2.2-1.49 3.16-1.18 3.16-1.18.63 1.59.24 2.76.12 3.05.74.81 1.18 1.83 1.18 3.09 0 4.41-2.69 5.38-5.25 5.67.41.35.77 1.05.77 2.12v3.15c0 .3.21.67.8.55A11.51 11.51 0 0 0 23.5 12C23.5 5.65 18.35.5 12 .5z"/></svg>
          <?= t('hero_gh') ?>
        </a>
      </div>
      <p class="hero-note"><?= $ver ? sprintf(t('hero_note'), htmlspecialchars($ver)) : t('hero_badge') ?></p>
    </div>

    <!-- Mockup animado do app -->
    <div class="hero-app" aria-hidden="true">
      <div class="appwin">
        <div class="appwin-bar"><i></i><i></i><i></i><span>Unit Mail</span></div>
        <div class="appwin-body">
          <aside>
            <b><?= t('mock_inbox') ?><em class="nbadge"></em></b>
            <span><?= t('mock_sent') ?></span>
            <span><?= t('mock_drafts') ?></span>
            <span><?= t('mock_deleted') ?></span>
            <span><?= t('mock_junk') ?></span>
          </aside>
          <div class="list">
            <div class="msg m1"><b><?= t('mock_subject1') ?></b><span><?= t('mock_from1') ?></span></div>
            <div class="msg m2"><b><?= t('mock_subject2') ?></b><span><?= t('mock_from2') ?></span></div>
            <div class="msg m3"><b><?= t('mock_subject3') ?></b><span><?= t('mock_from3') ?></span></div>
          </div>
          <div class="read">
            <div class="read-bar"></div>
            <div class="read-line w80"></div>
            <div class="read-line w60"></div>
            <div class="read-line w70"></div>
          </div>
        </div>
        <div class="fly">✉</div>
        <div class="toast">📬 <?= t('mock_newmail') ?></div>
      </div>
    </div>
  </div>
</section>

<!-- Recursos -->
<section class="feats" id="recursos">
  <div class="wrap">
    <h2 class="sec-title"><?= t('feat_title') ?></h2>
    <p class="sec-sub"><?= t('feat_sub') ?></p>
    <div class="feat-grid">
      <article class="feat"><i>📂</i><h3><?= t('f_pst_t') ?></h3><p><?= t('f_pst_d') ?></p></article>
      <article class="feat"><i>📨</i><h3><?= t('f_acc_t') ?></h3><p><?= t('f_acc_d') ?></p></article>
      <article class="feat"><i>📅</i><h3><?= t('f_cal_t') ?></h3><p><?= t('f_cal_d') ?></p></article>
      <article class="feat"><i>👥</i><h3><?= t('f_org_t') ?></h3><p><?= t('f_org_d') ?></p></article>
      <article class="feat"><i>🛡️</i><h3><?= t('f_rule_t') ?></h3><p><?= t('f_rule_d') ?></p></article>
      <article class="feat"><i>✍️</i><h3><?= t('f_draft_t') ?></h3><p><?= t('f_draft_d') ?></p></article>
      <article class="feat"><i>🌐</i><h3><?= t('f_lang_t') ?></h3><p><?= t('f_lang_d') ?></p></article>
      <article class="feat"><i>🔓</i><h3><?= t('f_free_t') ?></h3><p><?= t('f_free_d') ?></p></article>
    </div>
  </div>
</section>

<!-- Download — links diretos do GitHub, sempre a última release -->
<section class="dl" id="download">
  <div class="wrap">
    <h2 class="sec-title"><?= t('dl_title') ?></h2>
    <p class="sec-sub"><?= t('dl_sub') ?></p>
    <div class="dl-grid">
      <a class="dl-card" href="<?= htmlspecialchars($dlWin) ?>" target="_blank" rel="noopener">
        <svg width="34" height="34" viewBox="0 0 24 24" fill="currentColor"><path d="M3 5.5 10.5 4.4v7.1H3zM11.5 4.2 21 3v8.5h-9.5zM3 12.4h7.5v7.1L3 18.4zM11.5 12.4H21V21l-9.5-1.3z"/></svg>
        <div><b><?= t('dl_windows') ?></b><span><?= t('dl_win_sub') ?></span></div>
        <em><?= t('dl_get') ?> →</em>
      </a>
      <a class="dl-card" href="<?= htmlspecialchars($dlLin) ?>" target="_blank" rel="noopener">
        <svg width="34" height="34" viewBox="0 0 24 24" fill="currentColor"><path d="M12.5 2c1.9 0 3.2 1.5 3.2 3.3 0 1-.3 1.9-.8 2.7 1.6 1.2 2.8 3 3.3 5.2.2.9-.1 1.3-.9 1.2-.7-.1-1.1-.8-1.3-1.6-.3-1.2-.9-2.4-1.9-3.3 1.6 2.3 2.1 5.6 1.6 8.5 1.4.4 2.8 1 3.6 2 .7.9.3 1.9-.7 1.9h-11c-1 0-1.5-1.1-.7-2 .7-.8 2.1-1.5 3.4-1.9-.5-2.9 0-6.2 1.6-8.5-1 .9-1.6 2.1-1.9 3.3-.2.8-.6 1.5-1.3 1.6-.8.1-1.1-.3-.9-1.2.5-2.2 1.7-4 3.3-5.2-.5-.8-.8-1.7-.8-2.7C9.3 3.5 10.6 2 12.5 2z"/></svg>
        <div><b><?= t('dl_linux') ?></b><span><?= t('dl_lin_sub') ?></span></div>
        <em><?= t('dl_get') ?> →</em>
      </a>
    </div>
    <p class="center dl-note">
      <?= $ver ? sprintf(t('hero_note'), htmlspecialchars($ver)) . ' · ' : '' ?>
      <a href="<?= htmlspecialchars($cfg['releases_url']) ?>" target="_blank" rel="noopener"><?= t('dl_all') ?> →</a>
    </p>
  </div>
</section>

<!-- Doação -->
<section class="donate-cta" id="doacao">
  <div class="wrap">
    <h2><?= t('donate_title') ?></h2>
    <p><?= t('donate_sub') ?></p>
    <a class="btn btn-paypal" href="<?= htmlspecialchars($cfg['paypal_url']) ?>" target="_blank" rel="noopener">
      <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor"><path d="M20.067 8.478c.492.88.556 2.014.3 3.327-.74 3.806-3.276 5.12-6.514 5.12h-.5a.805.805 0 0 0-.794.68l-.04.22-.63 3.993-.03.17a.804.804 0 0 1-.794.68H7.72a.483.483 0 0 1-.477-.558L9.788 8.48a.806.806 0 0 1 .793-.681h4.916c.786 0 1.512-.053 2.172-.13.32-.037.654-.08.976-.127.243-.036.508-.064.766-.098z"/></svg>
      <?= t('donate_btn') ?>
    </a>
  </div>
</section>

<footer class="foot">
  <div class="wrap foot-in">
    <span>© <?= date('Y') ?> Unit Mail · GPL-3.0</span>
    <nav>
      <a href="<?= htmlspecialchars($cfg['releases_url']) ?>" target="_blank" rel="noopener"><?= t('footer_src') ?></a>
      <a href="<?= htmlspecialchars($cfg['issues_url']) ?>" target="_blank" rel="noopener"><?= t('footer_issues') ?></a>
    </nav>
    <span class="made"><?= t('footer_made') ?></span>
  </div>
</footer>

<script src="<?= htmlspecialchars($assets) ?>js/app.js?v=<?= $jsV ?>" defer></script>
</body>
</html>
