<?php
/**
 * Configuração do site Unit Mail.
 *
 * PAYPAL_URL: troque pelo seu link de doação do PayPal quando tiver
 * (ex.: https://www.paypal.com/donate?hosted_button_id=XXXXXX).
 */
return [
    'app_name'     => 'Unit Mail',
    'tagline'      => 'Seu email, do seu jeito.',
    'base_url'     => 'auto',                  // 'auto' detecta pelo SCRIPT_NAME (ou ex.: '/unit-mail/')
    'github_repo'  => 'unitwebinfo-oss/unit-mail',
    'releases_url' => 'https://github.com/unitwebinfo-oss/unit-mail/releases',
    'issues_url'   => 'https://github.com/unitwebinfo-oss/unit-mail/issues',
    // ⬇️ Link de doação do PayPal (hosted button)
    'paypal_url'   => 'https://www.paypal.com/donate/?hosted_button_id=X95ATUT9TXGNW',
    'cache_ttl'    => 3600,                    // cache da API do GitHub (segundos)
];
