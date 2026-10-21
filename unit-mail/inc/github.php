<?php
/**
 * Busca a última release no GitHub com cache em arquivo — evita bater no
 * rate-limit da API a cada visita.
 */
function ghLatestRelease(array $cfg): ?array
{
    $cacheFile = __DIR__ . '/../cache/release.json';
    if (is_file($cacheFile) && time() - filemtime($cacheFile) < ($cfg['cache_ttl'] ?? 3600)) {
        $cached = json_decode((string) file_get_contents($cacheFile), true);
        if (is_array($cached)) return $cached;
    }
    $ctx = stream_context_create([
        'http' => [
            'method'  => 'GET',
            'timeout' => 8,
            'header'  => "User-Agent: unit-mail-site\r\nAccept: application/vnd.github+json\r\n",
        ],
    ]);
    $url  = 'https://api.github.com/repos/' . $cfg['github_repo'] . '/releases/latest';
    $raw  = @file_get_contents($url, false, $ctx);
    $data = $raw ? json_decode($raw, true) : null;
    if (!is_array($data) || empty($data['tag_name'])) return null;

    $out = [
        'version' => $data['tag_name'],
        'url'     => $data['html_url'] ?? $cfg['releases_url'],
        'date'    => $data['published_at'] ?? '',
        'assets'  => [],
    ];
    foreach (($data['assets'] ?? []) as $a) {
        $out['assets'][] = [
            'name' => $a['name'],
            'url'  => $a['browser_download_url'],
            'size' => $a['size'] ?? 0,
        ];
    }
    @file_put_contents($cacheFile, json_encode($out));
    return $out;
}

function fmtSize(int $bytes): string
{
    if ($bytes <= 0) return '';
    $units = ['B', 'KB', 'MB', 'GB'];
    $i = (int) floor(log($bytes, 1024));
    return round($bytes / pow(1024, $i), 1) . ' ' . $units[$i];
}
